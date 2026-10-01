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


VERSION_JOURNAL = 1


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

    @property
    def aplicado(self) -> bool:
        return self.estado == EstadoPersistenciaExcel.APLICADO

    @property
    def encolado(self) -> bool:
        return self.estado == EstadoPersistenciaExcel.ENCOLADO


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
            f"Journal Excel inválido: "
            f"{ruta_journal}"
        ) from exc

    if not isinstance(
        contenido,
        dict,
    ):
        raise RuntimeError(
            f"Formato inválido de journal: "
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

    if not updates:
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
            "id_orden no puede estar vacío."
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

        intentos = (
            int(
                anterior.get(
                    "intentos",
                    0,
                )
            )
            if isinstance(anterior, dict)
            else 0
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
                + 1
            )

            actual[
                "ultimo_error"
            ] = str(error)

            _guardar_journal_sin_lock(
                ruta_excel,
                journal,
            )

    except Exception:
        # El error original de sincronización es más importante
        # que un fallo secundario actualizando metadata.
        pass


def persistir_o_encolar_resultado(
    ruta_excel: str | Path,
    id_orden: Any,
    *,
    estado_rpa: int,
    resumen: str | None,
) -> ResultadoPersistenciaExcel:
    """
    Intenta persistir inmediatamente en DATA.xlsx.

    Si falla la persistencia, guarda un journal persistente.

    IMPORTANTE:
    Esta función NO propaga errores de persistencia al flujo del portal.
    """
    id_normalizado = _normalizar_id(
        id_orden
    )

    try:
        from src.excel.state_manager import (
            actualizar_resultado_orden,
        )

        actualizar_resultado_orden(
            ruta_excel,
            id_normalizado,
            estado_rpa=estado_rpa,
            resumen=resumen,
        )

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
        )

    except Exception as error_excel:
        try:
            journal = (
                encolar_actualizacion_excel(
                    ruta_excel,
                    id_normalizado,
                    estado_rpa=estado_rpa,
                    resumen=resumen,
                    detalle_error=str(
                        error_excel
                    ),
                )
            )

            return ResultadoPersistenciaExcel(
                estado=(
                    EstadoPersistenciaExcel.ENCOLADO
                ),
                id_orden=id_normalizado,
                detalle=(
                    "No se pudo actualizar "
                    "el Excel inmediatamente. "
                    "El resultado quedó "
                    "guardado en la cola persistente. "
                    f"Motivo: {error_excel}"
                ),
                ruta_journal=journal,
            )

        except Exception as error_journal:
            return ResultadoPersistenciaExcel(
                estado=(
                    EstadoPersistenciaExcel.NO_PERSISTIDO
                ),
                id_orden=id_normalizado,
                detalle=(
                    "La operación del portal ya terminó, "
                    "pero no fue posible actualizar Excel "
                    "ni guardar el journal. "
                    f"Error Excel: {error_excel}. "
                    f"Error journal: {error_journal}"
                ),
                ruta_journal=None,
            )


def sincronizar_actualizaciones_pendientes(
    ruta_excel: str | Path,
) -> dict[str, Any]:
    """
    Intenta aplicar todas las actualizaciones pendientes
    sobre la versión ACTUAL del Excel.

    Una actualización eliminada del journal significa que
    ya fue aplicada exitosamente.

    Si la app se cierra después del save pero antes de eliminar
    la entrada, la siguiente ejecución repetirá el mismo valor:
    la operación es idempotente.
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

    for actualizacion in pendientes:
        id_orden = str(
            actualizacion["id_orden"]
        )

        version = str(
            actualizacion.get(
                "updated_at",
                "",
            )
        )

        try:
            actualizar_resultado_orden(
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

        except Exception as error:
            resultado[
                "fallidas"
            ] += 1

            resultado[
                "detalles"
            ].append(
                {
                    "id_orden": id_orden,
                    "estado": "PENDIENTE",
                    "detalle": str(
                        error
                    ),
                }
            )

            _registrar_intento_fallido(
                ruta_excel,
                id_orden,
                version,
                error,
            )

            # Si el Excel está ocupado, normalmente todas
            # las siguientes actualizaciones fallarán igual.
            # No eliminamos ninguna.
            continue

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
