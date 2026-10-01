from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any
import hashlib
import json
import os
import time
import uuid

from src.excel.retry_policy import (
    clasificar_error_persistencia,
    ejecutar_con_reintentos,
)


VERSION_JOURNAL = 2


class EstadoPersistenciaExcel(str, Enum):
    APLICADO = "APLICADO"
    ENCOLADO = "ENCOLADO"
    NO_PERSISTIDO = "NO_PERSISTIDO"


@dataclass(frozen=True)
class ResultadoPersistenciaExcel:
    estado: EstadoPersistenciaExcel
    id_orden: str
    detalle: str
    ruta_journal: Path | None = None
    intentos: int = 1
    codigo_error: str | None = None
    recuperable: bool | None = None

    @property
    def aplicado(self) -> bool:
        return (
            self.estado
            == EstadoPersistenciaExcel.APLICADO
        )

    @property
    def encolado(self) -> bool:
        return (
            self.estado
            == EstadoPersistenciaExcel.ENCOLADO
        )


def _ahora_iso() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


def _normalizar_id(valor: Any) -> str:
    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def _ruta_resuelta(
    ruta_excel: str | Path,
) -> Path:
    return Path(
        ruta_excel
    ).expanduser().resolve()


def _directorio_journals() -> Path:
    """
    Directorio persistente del journal.

    En tests puede sobrescribirse mediante:
        RPA_CBN_PENDING_DIR

    En Windows usa LOCALAPPDATA para que la cola sobreviva
    al cierre de la aplicación y no ensucie la carpeta del Excel.
    """
    configurado = os.getenv(
        "RPA_CBN_PENDING_DIR",
        "",
    ).strip()

    if configurado:
        base = Path(
            configurado
        ).expanduser()
    else:
        local_app_data = os.getenv(
            "LOCALAPPDATA",
            "",
        ).strip()

        if local_app_data:
            base = (
                Path(local_app_data)
                / "RPA_CBN"
                / "pending_excel"
            )
        else:
            base = (
                Path.home()
                / ".rpa_cbn"
                / "pending_excel"
            )

    base.mkdir(
        parents=True,
        exist_ok=True,
    )

    return base


def _clave_excel(
    ruta_excel: str | Path,
) -> str:
    ruta = _ruta_resuelta(
        ruta_excel
    )

    texto = str(
        ruta
    ).casefold()

    return hashlib.sha256(
        texto.encode("utf-8")
    ).hexdigest()[:20]


def ruta_journal_excel(
    ruta_excel: str | Path,
) -> Path:
    ruta = _ruta_resuelta(
        ruta_excel
    )

    clave = _clave_excel(
        ruta
    )

    nombre_seguro = (
        ruta.stem
        .replace(" ", "_")
        .replace(".", "_")
    )

    return (
        _directorio_journals()
        / f"{nombre_seguro}_{clave}.json"
    )


def _ruta_lock_journal(
    ruta_excel: str | Path,
) -> Path:
    journal = ruta_journal_excel(
        ruta_excel
    )

    return journal.with_suffix(
        journal.suffix + ".lock"
    )


@contextmanager
def _journal_lock(
    ruta_excel: str | Path,
    timeout_segundos: float = 10.0,
):
    ruta_lock = _ruta_lock_journal(
        ruta_excel
    )

    inicio = time.monotonic()
    fd = None

    while True:
        try:
            fd = os.open(
                str(ruta_lock),
                os.O_CREAT
                | os.O_EXCL
                | os.O_WRONLY,
            )

            contenido = (
                f"pid={os.getpid()} "
                f"time={time.time()}"
            )

            os.write(
                fd,
                contenido.encode("utf-8"),
            )

            break

        except FileExistsError:
            try:
                antiguedad = (
                    time.time()
                    - ruta_lock.stat().st_mtime
                )

                if antiguedad > 120:
                    ruta_lock.unlink()
                    continue

            except FileNotFoundError:
                continue

            if (
                time.monotonic() - inicio
                >= timeout_segundos
            ):
                raise TimeoutError(
                    f"No se pudo obtener lock "
                    f"del journal: {ruta_lock}"
                )

            time.sleep(0.05)

    try:
        yield

    finally:
        if fd is not None:
            try:
                os.close(fd)
            except Exception:
                pass

        try:
            ruta_lock.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass


def _journal_vacio(
    ruta_excel: str | Path,
) -> dict[str, Any]:
    return {
        "version": VERSION_JOURNAL,
        "target_excel": str(
            _ruta_resuelta(ruta_excel)
        ),
        "updates": {},
        "failed_updates": {},
    }


def _leer_journal_sin_lock(
    ruta_excel: str | Path,
) -> dict[str, Any]:
    ruta_journal = ruta_journal_excel(
        ruta_excel
    )

    if not ruta_journal.exists():
        return _journal_vacio(
            ruta_excel
        )

    try:
        contenido = json.loads(
            ruta_journal.read_text(
                encoding="utf-8"
            )
        )

    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Journal Excel inv?lido: "
            f"{ruta_journal}"
        ) from exc

    if not isinstance(
        contenido,
        dict,
    ):
        raise RuntimeError(
            f"Formato inv?lido de journal: "
            f"{ruta_journal}"
        )

    updates = contenido.get(
        "updates"
    )

    if not isinstance(
        updates,
        dict,
    ):
        raise RuntimeError(
            f"Journal sin diccionario "
            f"'updates': {ruta_journal}"
        )

    # Compatibilidad con journals CP7 / schema v1.
    failed_updates = contenido.get(
        "failed_updates",
        {},
    )

    if not isinstance(
        failed_updates,
        dict,
    ):
        raise RuntimeError(
            f"Journal con 'failed_updates' inv?lido: "
            f"{ruta_journal}"
        )

    contenido["failed_updates"] = (
        failed_updates
    )

    contenido.setdefault(
        "version",
        1,
    )

    return contenido


def _guardar_journal_sin_lock(
    ruta_excel: str | Path,
    journal: dict[str, Any],
) -> None:
    ruta_journal = ruta_journal_excel(
        ruta_excel
    )

    updates = journal.get(
        "updates",
        {},
    )

    failed_updates = journal.get(
        "failed_updates",
        {},
    )

    if (
        not updates
        and not failed_updates
    ):
        try:
            ruta_journal.unlink()
        except FileNotFoundError:
            pass

        return

    temporal = ruta_journal.with_name(
        f".{ruta_journal.name}."
        f"{os.getpid()}."
        f"{uuid.uuid4().hex}.tmp"
    )

    try:
        contenido = json.dumps(
            journal,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )

        with temporal.open(
            "w",
            encoding="utf-8",
            newline="\n",
        ) as archivo:
            archivo.write(
                contenido
            )
            archivo.flush()
            os.fsync(
                archivo.fileno()
            )

        os.replace(
            temporal,
            ruta_journal,
        )

    finally:
        try:
            temporal.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass


def obtener_actualizaciones_pendientes(
    ruta_excel: str | Path,
) -> list[dict[str, Any]]:
    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        actualizaciones = [
            dict(valor)
            for valor
            in journal["updates"].values()
        ]

    return sorted(
        actualizaciones,
        key=lambda item: (
            str(
                item.get(
                    "created_at",
                    "",
                )
            ),
            str(
                item.get(
                    "id_orden",
                    "",
                )
            ),
        ),
    )


def obtener_actualizaciones_fallidas(
    ruta_excel: str | Path,
) -> list[dict[str, Any]]:
    """
    Devuelve operaciones preservadas como evidencia,
    pero excluidas del retry autom?tico.
    """
    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        fallidas = [
            dict(valor)
            for valor
            in journal[
                "failed_updates"
            ].values()
        ]

    return sorted(
        fallidas,
        key=lambda item: (
            str(
                item.get(
                    "sync_failed_at",
                    "",
                )
            ),
            str(
                item.get(
                    "id_orden",
                    "",
                )
            ),
        ),
    )


def hay_actualizaciones_pendientes(
    ruta_excel: str | Path,
) -> bool:
    return bool(
        obtener_actualizaciones_pendientes(
            ruta_excel
        )
    )


def encolar_actualizacion_excel(
    ruta_excel: str | Path,
    id_orden: Any,
    *,
    estado_rpa: int,
    resumen: str | None,
    detalle_error: str = "",
    intentos_realizados: int = 0,
) -> Path:
    if estado_rpa not in (
        0,
        1,
        2,
    ):
        raise ValueError(
            "estado_rpa solo puede ser "
            "0, 1 o 2."
        )

    id_normalizado = _normalizar_id(
        id_orden
    )

    if not id_normalizado:
        raise ValueError(
            "id_orden no puede estar vac?o."
        )

    ahora = _ahora_iso()

    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        anterior = (
            journal["updates"].get(
                id_normalizado
            )
        )

        created_at = (
            anterior.get("created_at")
            if isinstance(anterior, dict)
            else ahora
        )

        intentos_previos = (
            int(
                anterior.get(
                    "intentos",
                    0,
                )
            )
            if isinstance(anterior, dict)
            else 0
        )

        intentos = (
            intentos_previos
            + max(
                0,
                int(intentos_realizados),
            )
        )

        journal["updates"][
            id_normalizado
        ] = {
            "id_orden": id_normalizado,
            "estado_rpa": estado_rpa,
            "resumen": (
                None
                if resumen is None
                else str(resumen).strip()
            ),
            "created_at": created_at,
            "updated_at": ahora,
            "intentos": intentos,
            "ultimo_error": str(
                detalle_error or ""
            ),
        }

        # Una actualizaci?n nueva/reactualizada revive
        # el ID aunque una versi?n anterior hubiese sido
        # aislada como no recuperable.
        journal[
            "failed_updates"
        ].pop(
            id_normalizado,
            None,
        )

        journal["version"] = (
            VERSION_JOURNAL
        )

        journal["target_excel"] = str(
            _ruta_resuelta(
                ruta_excel
            )
        )

        _guardar_journal_sin_lock(
            ruta_excel,
            journal,
        )

    return ruta_journal_excel(
        ruta_excel
    )


def _eliminar_actualizacion_si_version(
    ruta_excel: str | Path,
    id_orden: str,
    updated_at: str,
) -> bool:
    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        actual = journal[
            "updates"
        ].get(
            id_orden
        )

        if not isinstance(
            actual,
            dict,
        ):
            return False

        if (
            str(
                actual.get(
                    "updated_at",
                    "",
                )
            )
            != updated_at
        ):
            # Otro proceso generó una versión más nueva.
            # No debemos borrarla.
            return False

        del journal[
            "updates"
        ][
            id_orden
        ]

        _guardar_journal_sin_lock(
            ruta_excel,
            journal,
        )

        return True


def _registrar_intento_fallido(
    ruta_excel: str | Path,
    id_orden: str,
    updated_at: str,
    error: Exception,
    *,
    incremento: int = 1,
    codigo_error: str | None = None,
) -> None:
    try:
        with _journal_lock(
            ruta_excel
        ):
            journal = (
                _leer_journal_sin_lock(
                    ruta_excel
                )
            )

            actual = journal[
                "updates"
            ].get(
                id_orden
            )

            if not isinstance(
                actual,
                dict,
            ):
                return

            if (
                str(
                    actual.get(
                        "updated_at",
                        "",
                    )
                )
                != updated_at
            ):
                return

            actual["intentos"] = (
                int(
                    actual.get(
                        "intentos",
                        0,
                    )
                )
                + max(
                    1,
                    int(incremento),
                )
            )

            actual[
                "ultimo_error"
            ] = str(error)

            if codigo_error:
                actual[
                    "ultimo_codigo_error"
                ] = str(codigo_error)

            journal["version"] = (
                VERSION_JOURNAL
            )

            _guardar_journal_sin_lock(
                ruta_excel,
                journal,
            )

    except Exception:
        # El error original de sincronizaci?n es m?s importante
        # que un fallo secundario actualizando metadata.
        pass


def _mover_actualizacion_a_fallidas_si_version(
    ruta_excel: str | Path,
    id_orden: str,
    updated_at: str,
    *,
    error: Exception,
    codigo_error: str,
    detalle_error: str,
    intentos_realizados: int,
) -> bool:
    """
    Mueve una actualizaci?n fuera de la cola autom?tica sin
    perder su contenido ni la evidencia del error.
    """
    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        actual = journal[
            "updates"
        ].get(
            id_orden
        )

        if not isinstance(
            actual,
            dict,
        ):
            return False

        if (
            str(
                actual.get(
                    "updated_at",
                    "",
                )
            )
            != updated_at
        ):
            return False

        fallida = dict(
            actual
        )

        fallida["intentos"] = (
            int(
                fallida.get(
                    "intentos",
                    0,
                )
            )
            + max(
                1,
                int(intentos_realizados),
            )
        )

        fallida[
            "sync_failed_at"
        ] = _ahora_iso()

        fallida[
            "sync_error_code"
        ] = str(codigo_error)

        fallida[
            "sync_error_detail"
        ] = str(detalle_error)

        fallida[
            "sync_error_type"
        ] = type(error).__name__

        fallida[
            "sync_recoverable"
        ] = False

        del journal[
            "updates"
        ][
            id_orden
        ]

        journal[
            "failed_updates"
        ][
            id_orden
        ] = fallida

        journal["version"] = (
            VERSION_JOURNAL
        )

        _guardar_journal_sin_lock(
            ruta_excel,
            journal,
        )

        return True


def persistir_o_encolar_resultado(
    ruta_excel: str | Path,
    id_orden: Any,
    *,
    estado_rpa: int,
    resumen: str | None,
) -> ResultadoPersistenciaExcel:
    """
    Intenta persistir inmediatamente en DATA.xlsx.

    Pol?tica CP8:
    - errores recuperables se reintentan brevemente;
    - si persisten, se guardan en journal;
    - errores l?gicos/estructurales NO se encolan;
    - ning?n error de persistencia cambia el resultado
      ya obtenido en el portal.
    """
    id_normalizado = _normalizar_id(
        id_orden
    )

    from src.excel.state_manager import (
        actualizar_resultado_orden,
    )

    def operacion() -> int:
        return actualizar_resultado_orden(
            ruta_excel,
            id_normalizado,
            estado_rpa=estado_rpa,
            resumen=resumen,
        )

    intento = ejecutar_con_reintentos(
        operacion
    )

    if intento.exito:
        return ResultadoPersistenciaExcel(
            estado=(
                EstadoPersistenciaExcel.APLICADO
            ),
            id_orden=id_normalizado,
            detalle=(
                "Resultado aplicado "
                "directamente al Excel."
            ),
            ruta_journal=None,
            intentos=intento.intentos,
            codigo_error=None,
            recuperable=None,
        )

    error_excel = intento.error
    clasificacion = intento.clasificacion

    if (
        error_excel is None
        or clasificacion is None
    ):
        return ResultadoPersistenciaExcel(
            estado=(
                EstadoPersistenciaExcel.NO_PERSISTIDO
            ),
            id_orden=id_normalizado,
            detalle=(
                "Fall? la persistencia y no fue posible "
                "clasificar el error."
            ),
            ruta_journal=None,
            intentos=intento.intentos,
            codigo_error=(
                "ERROR_CLASIFICACION"
            ),
            recuperable=False,
        )

    # --------------------------------------------------------
    # Error l?gico/estructural: NO contaminar el journal.
    # --------------------------------------------------------

    if not clasificacion.recuperable:
        return ResultadoPersistenciaExcel(
            estado=(
                EstadoPersistenciaExcel.NO_PERSISTIDO
            ),
            id_orden=id_normalizado,
            detalle=(
                "La operaci?n del portal ya termin?, "
                "pero Excel rechaz? la persistencia por "
                "un error NO recuperable autom?ticamente. "
                f"[{clasificacion.codigo}] "
                f"{clasificacion.detalle}"
            ),
            ruta_journal=None,
            intentos=intento.intentos,
            codigo_error=(
                clasificacion.codigo
            ),
            recuperable=False,
        )

    # --------------------------------------------------------
    # Error transitorio que sobrevivi? a los retries.
    # Journal persistente.
    # --------------------------------------------------------

    try:
        journal = encolar_actualizacion_excel(
            ruta_excel,
            id_normalizado,
            estado_rpa=estado_rpa,
            resumen=resumen,
            detalle_error=(
                f"[{clasificacion.codigo}] "
                f"{clasificacion.detalle}"
            ),
            intentos_realizados=(
                intento.intentos
            ),
        )

        return ResultadoPersistenciaExcel(
            estado=(
                EstadoPersistenciaExcel.ENCOLADO
            ),
            id_orden=id_normalizado,
            detalle=(
                "No se pudo actualizar el Excel despu?s "
                f"de {intento.intentos} intento(s). "
                "El error es recuperable y el resultado "
                "qued? guardado en la cola persistente. "
                f"[{clasificacion.codigo}] "
                f"{clasificacion.detalle}"
            ),
            ruta_journal=journal,
            intentos=intento.intentos,
            codigo_error=(
                clasificacion.codigo
            ),
            recuperable=True,
        )

    except Exception as error_journal:
        return ResultadoPersistenciaExcel(
            estado=(
                EstadoPersistenciaExcel.NO_PERSISTIDO
            ),
            id_orden=id_normalizado,
            detalle=(
                "La operaci?n del portal ya termin?, "
                "pero no fue posible actualizar Excel "
                "ni guardar el journal. "
                f"Error Excel: {error_excel}. "
                f"Error journal: {error_journal}"
            ),
            ruta_journal=None,
            intentos=intento.intentos,
            codigo_error=(
                clasificacion.codigo
            ),
            recuperable=True,
        )


def sincronizar_actualizaciones_pendientes(
    ruta_excel: str | Path,
) -> dict[str, Any]:
    """
    Intenta aplicar la cola sobre la versi?n ACTUAL del Excel.

    CP8:
    - bloqueo temporal -> retry corto y permanece pendiente;
    - error l?gico -> se mueve a failed_updates;
    - un bloqueo global del archivo evita intentar in?tilmente
      todas las entradas restantes;
    - una entrada aislada no vuelve a reintentarse
      autom?ticamente.
    """
    pendientes = (
        obtener_actualizaciones_pendientes(
            ruta_excel
        )
    )

    resultado = {
        "total": len(
            pendientes
        ),
        "aplicadas": 0,
        "fallidas": 0,
        "recuperables": 0,
        "aisladas": 0,
        "restantes": len(
            pendientes
        ),
        "detalles": [],
    }

    if not pendientes:
        return resultado

    from src.excel.state_manager import (
        actualizar_resultado_orden,
    )

    codigos_bloqueo_global = {
        "EXCEL_OCUPADO",
        "WINDOWS_FILE_LOCK",
        "RPA_EXCEL_LOCK_TIMEOUT",
    }

    for indice, actualizacion in enumerate(
        pendientes
    ):
        id_orden = str(
            actualizacion["id_orden"]
        )

        version = str(
            actualizacion.get(
                "updated_at",
                "",
            )
        )

        def operacion() -> int:
            return actualizar_resultado_orden(
                ruta_excel,
                id_orden,
                estado_rpa=int(
                    actualizacion[
                        "estado_rpa"
                    ]
                ),
                resumen=actualizacion.get(
                    "resumen"
                ),
            )

        intento = ejecutar_con_reintentos(
            operacion
        )

        if intento.exito:
            eliminado = (
                _eliminar_actualizacion_si_version(
                    ruta_excel,
                    id_orden,
                    version,
                )
            )

            if eliminado:
                resultado[
                    "aplicadas"
                ] += 1

                resultado[
                    "detalles"
                ].append(
                    {
                        "id_orden": id_orden,
                        "estado": "APLICADO",
                        "detalle": "",
                        "intentos": (
                            intento.intentos
                        ),
                    }
                )

            continue

        error = (
            intento.error
            or RuntimeError(
                "Error de sincronizaci?n "
                "sin excepci?n disponible."
            )
        )

        clasificacion = (
            intento.clasificacion
            or clasificar_error_persistencia(
                error
            )
        )

        resultado[
            "fallidas"
        ] += 1

        # ----------------------------------------------------
        # RECUPERABLE
        # ----------------------------------------------------

        if clasificacion.recuperable:
            resultado[
                "recuperables"
            ] += 1

            _registrar_intento_fallido(
                ruta_excel,
                id_orden,
                version,
                error,
                incremento=(
                    intento.intentos
                ),
                codigo_error=(
                    clasificacion.codigo
                ),
            )

            resultado[
                "detalles"
            ].append(
                {
                    "id_orden": id_orden,
                    "estado": (
                        "PENDIENTE_RECUPERABLE"
                    ),
                    "codigo": (
                        clasificacion.codigo
                    ),
                    "detalle": (
                        clasificacion.detalle
                    ),
                    "intentos": (
                        intento.intentos
                    ),
                }
            )

            # Un bloqueo a nivel de archivo afectar? a todas
            # las operaciones. No tiene sentido esperar tres
            # veces por cada ID.
            if (
                clasificacion.codigo
                in codigos_bloqueo_global
            ):
                for restante in (
                    pendientes[
                        indice + 1:
                    ]
                ):
                    resultado[
                        "detalles"
                    ].append(
                        {
                            "id_orden": str(
                                restante[
                                    "id_orden"
                                ]
                            ),
                            "estado": (
                                "OMITIDO_BLOQUEO_GLOBAL"
                            ),
                            "codigo": (
                                clasificacion.codigo
                            ),
                            "detalle": (
                                "No se intent? porque "
                                "el archivo completo "
                                "contin?a bloqueado."
                            ),
                            "intentos": 0,
                        }
                    )

                break

            continue

        # ----------------------------------------------------
        # NO RECUPERABLE
        # ----------------------------------------------------

        movida = (
            _mover_actualizacion_a_fallidas_si_version(
                ruta_excel,
                id_orden,
                version,
                error=error,
                codigo_error=(
                    clasificacion.codigo
                ),
                detalle_error=(
                    clasificacion.detalle
                ),
                intentos_realizados=(
                    intento.intentos
                ),
            )
        )

        if movida:
            resultado[
                "aisladas"
            ] += 1

        resultado[
            "detalles"
        ].append(
            {
                "id_orden": id_orden,
                "estado": (
                    "AISLADO_NO_RECUPERABLE"
                    if movida
                    else "NO_RECUPERABLE_NO_MOVIDO"
                ),
                "codigo": (
                    clasificacion.codigo
                ),
                "detalle": (
                    clasificacion.detalle
                ),
                "intentos": (
                    intento.intentos
                ),
            }
        )

    resultado[
        "restantes"
    ] = len(
        obtener_actualizaciones_pendientes(
            ruta_excel
        )
    )

    return resultado


def aplicar_actualizaciones_pendientes_a_snapshot(
    ruta_snapshot: str | Path,
    ruta_excel_original: str | Path,
) -> dict[str, Any]:
    """
    Aplica el journal SOLO sobre el snapshot de ejecución.

    Esto es fundamental si DATA.xlsx sigue abierto:
    el snapshot podría contener ESTADO_RPA=0 aunque CBN ya
    haya procesado la orden en una ejecución anterior.

    El journal NO se elimina aquí porque aún no se ha
    sincronizado con el Excel original.
    """
    pendientes = (
        obtener_actualizaciones_pendientes(
            ruta_excel_original
        )
    )

    resultado = {
        "total": len(
            pendientes
        ),
        "aplicadas": 0,
        "fallidas": 0,
        "detalles": [],
    }

    if not pendientes:
        return resultado

    from src.excel.state_manager import (
        actualizar_resultado_orden,
    )

    for actualizacion in pendientes:
        id_orden = str(
            actualizacion[
                "id_orden"
            ]
        )

        try:
            actualizar_resultado_orden(
                ruta_snapshot,
                id_orden,
                estado_rpa=int(
                    actualizacion[
                        "estado_rpa"
                    ]
                ),
                resumen=actualizacion.get(
                    "resumen"
                ),
            )

            resultado[
                "aplicadas"
            ] += 1

            resultado[
                "detalles"
            ].append(
                {
                    "id_orden": id_orden,
                    "estado": "OVERLAY_OK",
                    "detalle": "",
                }
            )

        except Exception as error:
            resultado[
                "fallidas"
            ] += 1

            resultado[
                "detalles"
            ].append(
                {
                    "id_orden": id_orden,
                    "estado": "OVERLAY_ERROR",
                    "detalle": str(
                        error
                    ),
                }
            )

    return resultado
