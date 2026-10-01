from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Callable
import time

from src.excel.excel_transaction import (
    ExcelAccesoError,
    ExcelArchivoCambioConcurrenteError,
    ExcelArchivoOcupadoError,
    ExcelPersistenciaError,
)


WINERROR_SHARING_VIOLATION = 32
WINERROR_LOCK_VIOLATION = 33

INTENTOS_PERSISTENCIA_DEFAULT = 3
ESPERAS_PERSISTENCIA_DEFAULT = (
    0.15,
    0.35,
)


class TipoErrorPersistencia(str, Enum):
    RECUPERABLE = "RECUPERABLE"
    NO_RECUPERABLE = "NO_RECUPERABLE"


@dataclass(frozen=True)
class ClasificacionErrorPersistencia:
    tipo: TipoErrorPersistencia
    codigo: str
    detalle: str
    winerror: int | None = None

    @property
    def recuperable(self) -> bool:
        return (
            self.tipo
            == TipoErrorPersistencia.RECUPERABLE
        )


@dataclass(frozen=True)
class ResultadoReintento:
    exito: bool
    intentos: int
    valor: Any = None
    error: Exception | None = None
    clasificacion: (
        ClasificacionErrorPersistencia
        | None
    ) = None


def _cadena_errores(
    error: Exception,
) -> list[Exception]:
    """
    Recorre wrappers comunes sin entrar en ciclos.

    Algunos errores del RPA conservan el original en
    `error_original`; Python también puede usar
    __cause__ / __context__.
    """
    resultado: list[Exception] = []
    vistos: set[int] = set()

    actual: Exception | None = error

    while (
        actual is not None
        and id(actual) not in vistos
    ):
        vistos.add(
            id(actual)
        )

        resultado.append(
            actual
        )

        siguiente = getattr(
            actual,
            "error_original",
            None,
        )

        if not isinstance(
            siguiente,
            Exception,
        ):
            siguiente = (
                actual.__cause__
                or actual.__context__
            )

        if not isinstance(
            siguiente,
            Exception,
        ):
            siguiente = None

        actual = siguiente

    return resultado


def clasificar_error_persistencia(
    error: Exception,
) -> ClasificacionErrorPersistencia:
    """
    Clasifica exclusivamente errores producidos al
    persistir resultados del RPA en Excel.

    Recuperables:
    - Excel abierto/bloqueado.
    - WinError 32 / 33.
    - timeout del lock interno del RPA.

    El resto se considera no recuperable automáticamente:
    estructura inválida, ID inexistente, permisos,
    archivo inexistente, corrupción, bugs, etc.
    """
    cadena = _cadena_errores(
        error
    )

    # --------------------------------------------------------
    # Primero buscamos evidencia fuerte de bloqueo temporal.
    # --------------------------------------------------------

    for actual in cadena:
        if isinstance(
            actual,
            ExcelArchivoCambioConcurrenteError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.RECUPERABLE
                ),
                codigo="EXCEL_CAMBIO_CONCURRENTE",
                detalle=str(actual),
                winerror=None,
            )

        if isinstance(
            actual,
            ExcelArchivoOcupadoError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.RECUPERABLE
                ),
                codigo="EXCEL_OCUPADO",
                detalle=str(actual),
                winerror=getattr(
                    actual,
                    "winerror",
                    None,
                ),
            )

        winerror = getattr(
            actual,
            "winerror",
            None,
        )

        if winerror in (
            WINERROR_SHARING_VIOLATION,
            WINERROR_LOCK_VIOLATION,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.RECUPERABLE
                ),
                codigo="WINDOWS_FILE_LOCK",
                detalle=str(actual),
                winerror=winerror,
            )

        if isinstance(
            actual,
            TimeoutError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.RECUPERABLE
                ),
                codigo="RPA_EXCEL_LOCK_TIMEOUT",
                detalle=str(actual),
                winerror=None,
            )

    # --------------------------------------------------------
    # No recuperables automáticamente.
    # --------------------------------------------------------

    for actual in cadena:
        if isinstance(
            actual,
            FileNotFoundError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="EXCEL_NO_EXISTE",
                detalle=str(actual),
            )

        if isinstance(
            actual,
            ExcelAccesoError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="EXCEL_ERROR_ACCESO",
                detalle=str(actual),
                winerror=getattr(
                    actual,
                    "winerror",
                    None,
                ),
            )

        if isinstance(
            actual,
            ExcelPersistenciaError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="EXCEL_PERSISTENCIA_INVALIDA",
                detalle=str(actual),
            )

        if isinstance(
            actual,
            ValueError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="DATOS_EXCEL_INVALIDOS",
                detalle=str(actual),
            )

        if isinstance(
            actual,
            KeyError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="ESTRUCTURA_EXCEL_INVALIDA",
                detalle=str(actual),
            )

        if isinstance(
            actual,
            PermissionError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="PERMISO_DENEGADO",
                detalle=str(actual),
                winerror=getattr(
                    actual,
                    "winerror",
                    None,
                ),
            )

        if isinstance(
            actual,
            OSError,
        ):
            return ClasificacionErrorPersistencia(
                tipo=(
                    TipoErrorPersistencia.NO_RECUPERABLE
                ),
                codigo="ERROR_SISTEMA_ARCHIVOS",
                detalle=str(actual),
                winerror=getattr(
                    actual,
                    "winerror",
                    None,
                ),
            )

    return ClasificacionErrorPersistencia(
        tipo=(
            TipoErrorPersistencia.NO_RECUPERABLE
        ),
        codigo="ERROR_PERSISTENCIA_DESCONOCIDO",
        detalle=str(error),
    )


def ejecutar_con_reintentos(
    operacion: Callable[[], Any],
    *,
    intentos_maximos: int = (
        INTENTOS_PERSISTENCIA_DEFAULT
    ),
    esperas: tuple[float, ...] = (
        ESPERAS_PERSISTENCIA_DEFAULT
    ),
) -> ResultadoReintento:
    """
    Ejecuta una operación y solo reintenta cuando el error
    está clasificado explícitamente como recuperable.

    Un error lógico no duerme ni se repite.
    """
    if intentos_maximos < 1:
        raise ValueError(
            "intentos_maximos debe ser >= 1."
        )

    ultimo_error: Exception | None = None
    ultima_clasificacion = None

    for intento in range(
        1,
        intentos_maximos + 1,
    ):
        try:
            valor = operacion()

            return ResultadoReintento(
                exito=True,
                intentos=intento,
                valor=valor,
            )

        except Exception as error:
            ultimo_error = error

            clasificacion = (
                clasificar_error_persistencia(
                    error
                )
            )

            ultima_clasificacion = (
                clasificacion
            )

            if not clasificacion.recuperable:
                return ResultadoReintento(
                    exito=False,
                    intentos=intento,
                    error=error,
                    clasificacion=clasificacion,
                )

            if intento >= intentos_maximos:
                break

            if esperas:
                indice_espera = min(
                    intento - 1,
                    len(esperas) - 1,
                )

                espera = max(
                    0.0,
                    float(
                        esperas[
                            indice_espera
                        ]
                    ),
                )

                if espera:
                    time.sleep(
                        espera
                    )

    return ResultadoReintento(
        exito=False,
        intentos=intentos_maximos,
        error=ultimo_error,
        clasificacion=ultima_clasificacion,
    )
