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


VERSION_JOURNAL = 3


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



class ReintentoManualNoPermitidoError(
    RuntimeError
):
    """
    Impide un reintento manual cuando existe evidencia
    persistente de que la orden no es segura para reprocesar.
    """
    pass

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

        except (
            FileExistsError,
            PermissionError,
        ) as error_lock:
            # En Windows, dos procesos/threads compitiendo
            # por O_CREAT | O_EXCL pueden producir tanto
            # FileExistsError como PermissionError.
            #
            # Ambos representan contencion transitoria del
            # lock y deben seguir la misma politica de espera.
            try:
                antiguedad = (
                    time.time()
                    - ruta_lock.stat().st_mtime
                )

                if antiguedad > 120:
                    try:
                        ruta_lock.unlink()

                    except FileNotFoundError:
                        continue

                    except PermissionError:
                        # Otro proceso todavia mantiene
                        # acceso al lock. No se roba.
                        pass

                    else:
                        continue

            except FileNotFoundError:
                # Carrera entre comprobacion y eliminacion.
                # Volvemos a intentar inmediatamente.
                continue

            except PermissionError:
                # Windows puede impedir temporalmente incluso
                # consultar metadata del archivo bloqueado.
                pass

            if (
                time.monotonic() - inicio
                >= timeout_segundos
            ):
                raise TimeoutError(
                    f"No se pudo obtener lock "
                    f"del journal: {ruta_lock}"
                ) from error_lock

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
        "inflight": {},
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
            f"Journal con 'failed_updates' inválido: "
            f"{ruta_journal}"
        )

    contenido["failed_updates"] = (
        failed_updates
    )

    inflight = contenido.get(
        "inflight",
        {},
    )

    if not isinstance(
        inflight,
        dict,
    ):
        raise RuntimeError(
            f"Journal con 'inflight' invalido: "
            f"{ruta_journal}"
        )

    contenido["inflight"] = inflight

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

    inflight = journal.get(
        "inflight",
        {},
    )

    if (
        not updates
        and not failed_updates
        and not inflight
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


def obtener_estado_journal(
    ruta_excel: str | Path,
) -> dict[str, list[dict[str, Any]]]:
    """
    Lee updates, failed_updates e inflight bajo un unico lock.
    """
    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        pendientes = [
            dict(valor)
            for valor
            in journal["updates"].values()
        ]

        fallidas = [
            dict(valor)
            for valor
            in journal[
                "failed_updates"
            ].values()
        ]

        inflight = [
            dict(valor)
            for valor
            in journal[
                "inflight"
            ].values()
        ]

    pendientes = sorted(
        pendientes,
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

    fallidas = sorted(
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

    inflight = sorted(
        inflight,
        key=lambda item: (
            str(
                item.get(
                    "started_at",
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

    return {
        "pendientes": pendientes,
        "fallidas": fallidas,
        "inflight": inflight,
    }



def preparar_reintento_manual(
    ruta_excel: str | Path,
    id_orden: Any,
) -> int:
    """
    Prepara una orden fallida para un reintento explícito.

    Antes del cambio 2 -> 0 verifica que el ID no tenga:
    - resultado pendiente de sincronización;
    - failed_update;
    - inflight.

    Estas condiciones significan que CBN pudo haber realizado
    una operación cuyo estado todavía no está reconciliado.
    """
    id_normalizado = _normalizar_id(
        id_orden
    )

    if not id_normalizado:
        raise ValueError(
            "id_orden no puede estar vacío."
        )

    estado = obtener_estado_journal(
        ruta_excel
    )

    bloqueos = (
        (
            "pendientes",
            (
                "tiene un resultado pendiente de "
                "sincronización con Excel"
            ),
        ),
        (
            "fallidas",
            (
                "requiere revisión manual por un "
                "fallo de sincronización"
            ),
        ),
        (
            "inflight",
            (
                "tiene una ejecución anterior "
                "con resultado incierto"
            ),
        ),
    )

    for coleccion, motivo in bloqueos:
        for item in estado.get(
            coleccion,
            [],
        ):
            if (
                _normalizar_id(
                    item.get(
                        "id_orden"
                    )
                )
                != id_normalizado
            ):
                continue

            raise (
                ReintentoManualNoPermitidoError(
                    "No es seguro reintentar "
                    f"ID_ORDEN={id_normalizado}: "
                    f"{motivo}. "
                    "Revise el estado de sincronización "
                    "antes de reprocesar la orden."
                )
            )

    from src.excel.state_manager import (
        reiniciar_orden_fallida,
    )

    return reiniciar_orden_fallida(
        ruta_excel,
        id_normalizado,
    )

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
    pero excluidas del retry automático.
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


def obtener_ordenes_inflight(
    ruta_excel: str | Path,
) -> list[dict[str, Any]]:
    return obtener_estado_journal(
        ruta_excel
    )["inflight"]


def registrar_orden_inflight(
    ruta_excel: str | Path,
    id_orden: Any,
    *,
    version_esperada: dict[str, Any] | None = None,
) -> str:
    """
    Registra durablemente que una orden va a entrar a CBN.

    Mientras exista esta marca, esa orden no debe volver
    a tratarse automaticamente como pendiente.
    """
    id_normalizado = _normalizar_id(
        id_orden
    )

    if not id_normalizado:
        raise ValueError(
            "id_orden no puede estar vacio."
        )

    ahora = _ahora_iso()
    inflight_id = uuid.uuid4().hex

    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        if (
            id_normalizado
            in journal["updates"]
        ):
            raise RuntimeError(
                "La orden ya tiene un resultado "
                "pendiente de sincronizacion. "
                f"ID_ORDEN={id_normalizado}."
            )

        if (
            id_normalizado
            in journal["failed_updates"]
        ):
            raise RuntimeError(
                "La orden ya requiere revision manual. "
                f"ID_ORDEN={id_normalizado}."
            )

        if (
            id_normalizado
            in journal["inflight"]
        ):
            raise RuntimeError(
                "La orden ya esta marcada como inflight. "
                f"ID_ORDEN={id_normalizado}."
            )

        journal["inflight"][
            id_normalizado
        ] = {
            "id_orden": id_normalizado,
            "inflight_id": inflight_id,
            "started_at": ahora,
            "updated_at": ahora,
            "pid": os.getpid(),
            "version_esperada": (
                dict(version_esperada)
                if version_esperada is not None
                else None
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

    return inflight_id


def cerrar_orden_inflight(
    ruta_excel: str | Path,
    id_orden: Any,
    inflight_id: str,
) -> bool:
    """
    Elimina inflight solo si pertenece exactamente
    a la instancia que lo creo.
    """
    id_normalizado = _normalizar_id(
        id_orden
    )

    token = str(
        inflight_id or ""
    ).strip()

    if (
        not id_normalizado
        or not token
    ):
        return False

    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        actual = journal[
            "inflight"
        ].get(
            id_normalizado
        )

        if not isinstance(
            actual,
            dict,
        ):
            return False

        if (
            str(
                actual.get(
                    "inflight_id",
                    "",
                )
            )
            != token
        ):
            return False

        del journal[
            "inflight"
        ][
            id_normalizado
        ]

        _guardar_journal_sin_lock(
            ruta_excel,
            journal,
        )

    return True


def _depurar_inflight_resueltos(
    ruta_excel: str | Path,
) -> int:
    """
    Recuperacion de crash:

    Si existen simultaneamente inflight y un resultado durable
    para el mismo ID en updates/failed_updates, inflight ya es
    redundante y puede eliminarse.
    """
    eliminados = 0

    with _journal_lock(
        ruta_excel
    ):
        journal = _leer_journal_sin_lock(
            ruta_excel
        )

        for id_orden in list(
            journal["inflight"].keys()
        ):
            if (
                id_orden
                in journal["updates"]
                or id_orden
                in journal["failed_updates"]
            ):
                del journal[
                    "inflight"
                ][
                    id_orden
                ]

                eliminados += 1

        if eliminados:
            _guardar_journal_sin_lock(
                ruta_excel,
                journal,
            )

    return eliminados


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
    version_esperada: dict[str, Any] | None = None,
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
            "version_esperada": (
                dict(version_esperada)
                if version_esperada is not None
                else None
            ),
        }

        # Una actualización nueva/reactualizada revive
        # el ID aunque una versión anterior hubiese sido
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
        # El error original de sincronización es más importante
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
    Mueve una actualización fuera de la cola automática sin
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


def registrar_actualizacion_fallida_excel(
    ruta_excel: str | Path,
    id_orden: Any,
    *,
    estado_rpa: int,
    resumen: str | None,
    codigo_error: str,
    detalle_error: str,
    tipo_error: str,
    intentos_realizados: int = 1,
    version_esperada: dict[str, Any] | None = None,
) -> Path:
    """
    Conserva de forma durable un resultado que CBN ya produjo,
    pero que no puede reflejarse automáticamente en Excel.

    La entrada queda fuera del retry automático y obliga a
    revisión manual antes de una nueva ejecución.
    """
    if estado_rpa not in (
        0,
        1,
        2,
    ):
        raise ValueError(
            "estado_rpa solo puede ser 0, 1 o 2."
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

        activa_anterior = (
            journal[
                "updates"
            ].pop(
                id_normalizado,
                None,
            )
        )

        fallida_anterior = (
            journal[
                "failed_updates"
            ].get(
                id_normalizado
            )
        )

        anterior = (
            activa_anterior
            if isinstance(
                activa_anterior,
                dict,
            )
            else (
                fallida_anterior
                if isinstance(
                    fallida_anterior,
                    dict,
                )
                else None
            )
        )

        created_at = (
            str(
                anterior.get(
                    "created_at",
                    ahora,
                )
            )
            if anterior is not None
            else ahora
        )

        intentos_previos = (
            int(
                anterior.get(
                    "intentos",
                    0,
                )
            )
            if anterior is not None
            else 0
        )

        intentos = (
            intentos_previos
            + max(
                1,
                int(
                    intentos_realizados
                ),
            )
        )

        journal[
            "failed_updates"
        ][
            id_normalizado
        ] = {
            "id_orden": id_normalizado,
            "estado_rpa": estado_rpa,
            "resumen": (
                None
                if resumen is None
                else str(
                    resumen
                ).strip()
            ),
            "created_at": created_at,
            "updated_at": ahora,
            "intentos": intentos,
            "ultimo_error": str(
                detalle_error or ""
            ),
            "sync_failed_at": ahora,
            "sync_error_code": str(
                codigo_error
            ),
            "sync_error_detail": str(
                detalle_error
            ),
            "sync_error_type": str(
                tipo_error
            ),
            "sync_recoverable": False,
            "version_esperada": (
                dict(version_esperada)
                if version_esperada is not None
                else (
                    anterior.get(
                        "version_esperada"
                    )
                    if isinstance(
                        anterior,
                        dict,
                    )
                    else None
                )
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


def persistir_o_encolar_resultado(
    ruta_excel: str | Path,
    id_orden: Any,
    *,
    estado_rpa: int,
    resumen: str | None,
    version_esperada: dict[str, Any] | None = None,
) -> ResultadoPersistenciaExcel:
    """
    Intenta persistir inmediatamente en DATA.xlsx.

    Política CP8:
    - errores recuperables se reintentan brevemente;
    - si persisten, se guardan en journal;
    - errores lógicos/estructurales no entran al retry automático; se aíslan en failed_updates;
    - ningún error de persistencia cambia el resultado
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
            version_esperada=version_esperada,
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
                "Falló la persistencia y no fue posible "
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
    # Error lógico/estructural: preservar en failed_updates sin incorporarlo al retry automático.
    # --------------------------------------------------------

    if not clasificacion.recuperable:
        try:
            journal = (
                registrar_actualizacion_fallida_excel(
                    ruta_excel,
                    id_normalizado,
                    estado_rpa=estado_rpa,
                    resumen=resumen,
                    codigo_error=(
                        clasificacion.codigo
                    ),
                    detalle_error=(
                        clasificacion.detalle
                    ),
                    tipo_error=(
                        type(
                            error_excel
                        ).__name__
                    ),
                    intentos_realizados=(
                        intento.intentos
                    ),
                    version_esperada=version_esperada,
                )
            )

            return ResultadoPersistenciaExcel(
                estado=(
                    EstadoPersistenciaExcel.NO_PERSISTIDO
                ),
                id_orden=id_normalizado,
                detalle=(
                    "La operación del portal ya terminó, "
                    "pero Excel rechazó la persistencia por "
                    "un error NO recuperable automáticamente. "
                    "El resultado quedó protegido en "
                    "failed_updates y requiere revisión manual. "
                    f"[{clasificacion.codigo}] "
                    f"{clasificacion.detalle}"
                ),
                ruta_journal=journal,
                intentos=intento.intentos,
                codigo_error=(
                    clasificacion.codigo
                ),
                recuperable=False,
            )

        except Exception as error_journal:
            return ResultadoPersistenciaExcel(
                estado=(
                    EstadoPersistenciaExcel.NO_PERSISTIDO
                ),
                id_orden=id_normalizado,
                detalle=(
                    "La operación del portal ya terminó y "
                    "Excel rechazó la persistencia. Además, "
                    "no fue posible conservar el resultado "
                    "en failed_updates. "
                    f"Error Excel: {error_excel}. "
                    f"Error journal: {error_journal}"
                ),
                ruta_journal=None,
                intentos=intento.intentos,
                codigo_error=(
                    clasificacion.codigo
                ),
                recuperable=False,
            )

    # --------------------------------------------------------
    # Error transitorio que sobrevivió a los retries.
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
            version_esperada=version_esperada,
        )

        return ResultadoPersistenciaExcel(
            estado=(
                EstadoPersistenciaExcel.ENCOLADO
            ),
            id_orden=id_normalizado,
            detalle=(
                "No se pudo actualizar el Excel después "
                f"de {intento.intentos} intento(s). "
                "El error es recuperable y el resultado "
                "quedó guardado en la cola persistente. "
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
                "La operación del portal ya terminó, "
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
    Intenta aplicar la cola sobre la versión ACTUAL del Excel.

    CP8:
    - bloqueo temporal -> retry corto y permanece pendiente;
    - error lógico -> se mueve a failed_updates;
    - un bloqueo global del archivo evita intentar inútilmente
      todas las entradas restantes;
    - una entrada aislada no vuelve a reintentarse
      automáticamente.
    """
    _depurar_inflight_resueltos(
        ruta_excel
    )

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
                version_esperada=actualizacion.get(
                    "version_esperada"
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
                "Error de sincronización "
                "sin excepción disponible."
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

            # Un bloqueo a nivel de archivo afectará a todas
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
                                "No se intentó porque "
                                "el archivo completo "
                                "continúa bloqueado."
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

    Un failed_update representa una operación que CBN ya pudo
    haber ejecutado, pero cuyo estado no puede reconciliarse
    automáticamente con Excel.

    En ese caso NO se intenta continuar: el runner debe detener
    la ejecución antes de abrir CBN para evitar reprocesamiento.
    """
    estado_journal = obtener_estado_journal(
        ruta_excel_original
    )

    pendientes = estado_journal[
        "pendientes"
    ]

    fallidas_persistentes = (
        estado_journal[
            "fallidas"
        ]
    )

    inflight_persistentes = (
        estado_journal.get(
            "inflight",
            [],
        )
    )

    resultado = {
        "total": len(
            pendientes
        ),
        "aplicadas": 0,
        "fallidas": len(
            fallidas_persistentes
        ),
        "bloqueadas": len(
            fallidas_persistentes
        ),
        "inflight": len(
            inflight_persistentes
        ),
        "omitidas_inflight": 0,
        "detalles": [],
    }

    if fallidas_persistentes:
        for actualizacion in (
            fallidas_persistentes
        ):
            resultado[
                "detalles"
            ].append(
                {
                    "id_orden": str(
                        actualizacion.get(
                            "id_orden",
                            "",
                        )
                    ),
                    "estado": (
                        "FAILED_UPDATE_REQUIERE_REVISION"
                    ),
                    "detalle": str(
                        actualizacion.get(
                            "sync_error_detail",
                            "",
                        )
                    ),
                }
            )

        return resultado

    if inflight_persistentes:
        from src.excel.state_manager import (
            actualizar_resultado_orden,
        )

        for actualizacion in inflight_persistentes:
            id_orden = str(
                actualizacion.get(
                    "id_orden",
                    "",
                )
            )

            try:
                actualizar_resultado_orden(
                    ruta_snapshot,
                    id_orden,
                    estado_rpa=2,
                    resumen=None,
                )

                resultado[
                    "omitidas_inflight"
                ] += 1

                resultado[
                    "detalles"
                ].append(
                    {
                        "id_orden": id_orden,
                        "estado": (
                            "INFLIGHT_OMITIDO"
                        ),
                        "detalle": (
                            "La orden queda fuera de pendientes "
                            "hasta resolver su ejecucion anterior."
                        ),
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
                        "estado": (
                            "INFLIGHT_OVERLAY_ERROR"
                        ),
                        "detalle": str(error),
                    }
                )

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
                version_esperada=actualizacion.get(
                    "version_esperada"
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
