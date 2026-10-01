from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
import ctypes
from ctypes import wintypes
import os


class EstadoAccesoExcel(str, Enum):
    DISPONIBLE = "DISPONIBLE"
    OCUPADO = "OCUPADO"
    ERROR_ACCESO = "ERROR_ACCESO"
    NO_EXISTE = "NO_EXISTE"


@dataclass(frozen=True)
class DiagnosticoAccesoExcel:
    ruta: Path
    estado: EstadoAccesoExcel
    puede_escribir: bool
    lock_office_detectado: bool
    detalle: str
    winerror: int | None = None

    @property
    def ocupado(self) -> bool:
        return self.estado == EstadoAccesoExcel.OCUPADO


# Windows API -------------------------------------------------

_GENERIC_WRITE = 0x40000000

_FILE_SHARE_READ = 0x00000001
_FILE_SHARE_WRITE = 0x00000002
_FILE_SHARE_DELETE = 0x00000004

_OPEN_EXISTING = 3
_FILE_ATTRIBUTE_NORMAL = 0x00000080

_ERROR_FILE_NOT_FOUND = 2
_ERROR_PATH_NOT_FOUND = 3
_ERROR_ACCESS_DENIED = 5
_ERROR_SHARING_VIOLATION = 32
_ERROR_LOCK_VIOLATION = 33


def _ruta_lock_office(ruta: Path) -> Path:
    """
    Excel suele crear un archivo auxiliar:

        DATA.xlsx
        ~$DATA.xlsx

    Su existencia es únicamente una señal auxiliar.
    Puede quedar obsoleto, por lo que NO determina por sí sola
    que el archivo esté bloqueado.
    """
    return ruta.with_name(f"~${ruta.name}")


def _probar_escritura_windows(
    ruta: Path,
) -> tuple[bool, int | None]:
    """
    Intenta obtener acceso de escritura al archivo utilizando CreateFileW.

    Se permiten todos los SHARE modes desde nuestro lado para no bloquear
    otros procesos. Si otro proceso (por ejemplo Excel) abrió el archivo
    sin compartir escritura, Windows devuelve ERROR_SHARING_VIOLATION.
    """
    kernel32 = ctypes.WinDLL(
        "kernel32",
        use_last_error=True,
    )

    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [
        wintypes.HANDLE,
    ]
    close_handle.restype = wintypes.BOOL

    invalid_handle = ctypes.c_void_p(-1).value

    handle = create_file(
        str(ruta),
        _GENERIC_WRITE,
        (
            _FILE_SHARE_READ
            | _FILE_SHARE_WRITE
            | _FILE_SHARE_DELETE
        ),
        None,
        _OPEN_EXISTING,
        _FILE_ATTRIBUTE_NORMAL,
        None,
    )

    handle_value = ctypes.cast(
        handle,
        ctypes.c_void_p,
    ).value

    if handle_value == invalid_handle:
        error = ctypes.get_last_error()
        return False, error

    try:
        return True, None
    finally:
        close_handle(handle)


def _probar_escritura_fallback(
    ruta: Path,
) -> tuple[bool, int | None]:
    """
    Fallback para entornos no Windows.

    RPA-CBN se ejecuta normalmente en Windows, pero mantenemos
    comportamiento razonable para tests/imports multiplataforma.
    """
    try:
        with ruta.open("r+b"):
            pass

        return True, None

    except PermissionError as exc:
        return False, getattr(
            exc,
            "winerror",
            None,
        )

    except OSError as exc:
        return False, getattr(
            exc,
            "winerror",
            None,
        )


def diagnosticar_acceso_excel(
    ruta_excel: str | Path,
) -> DiagnosticoAccesoExcel:
    """
    Diagnostica si el archivo puede ser escrito en este momento.

    IMPORTANTE:
    - No modifica el archivo.
    - No crea locks propios.
    - No cierra Excel.
    - El archivo ~$ de Office es solo información auxiliar.
    """
    ruta = Path(
        ruta_excel
    ).expanduser()

    try:
        ruta = ruta.resolve()
    except OSError:
        ruta = ruta.absolute()

    lock_office = _ruta_lock_office(
        ruta
    ).exists()

    if not ruta.exists():
        return DiagnosticoAccesoExcel(
            ruta=ruta,
            estado=EstadoAccesoExcel.NO_EXISTE,
            puede_escribir=False,
            lock_office_detectado=lock_office,
            detalle="El archivo no existe.",
        )

    if not ruta.is_file():
        return DiagnosticoAccesoExcel(
            ruta=ruta,
            estado=EstadoAccesoExcel.ERROR_ACCESO,
            puede_escribir=False,
            lock_office_detectado=lock_office,
            detalle="La ruta existe, pero no es un archivo.",
        )

    if os.name == "nt":
        disponible, winerror = (
            _probar_escritura_windows(
                ruta
            )
        )
    else:
        disponible, winerror = (
            _probar_escritura_fallback(
                ruta
            )
        )

    if disponible:
        detalle = (
            "Windows permite acceso de escritura."
        )

        if lock_office:
            detalle += (
                " Existe un archivo ~$ de Office, "
                "pero no está impidiendo la escritura."
            )

        return DiagnosticoAccesoExcel(
            ruta=ruta,
            estado=EstadoAccesoExcel.DISPONIBLE,
            puede_escribir=True,
            lock_office_detectado=lock_office,
            detalle=detalle,
        )

    if winerror in (
        _ERROR_SHARING_VIOLATION,
        _ERROR_LOCK_VIOLATION,
    ):
        return DiagnosticoAccesoExcel(
            ruta=ruta,
            estado=EstadoAccesoExcel.OCUPADO,
            puede_escribir=False,
            lock_office_detectado=lock_office,
            detalle=(
                "Otro proceso mantiene el archivo abierto "
                "sin permitir escritura."
            ),
            winerror=winerror,
        )

    if winerror in (
        _ERROR_FILE_NOT_FOUND,
        _ERROR_PATH_NOT_FOUND,
    ):
        return DiagnosticoAccesoExcel(
            ruta=ruta,
            estado=EstadoAccesoExcel.NO_EXISTE,
            puede_escribir=False,
            lock_office_detectado=lock_office,
            detalle=(
                "El archivo dejó de existir durante "
                "la comprobación."
            ),
            winerror=winerror,
        )

    if winerror == _ERROR_ACCESS_DENIED:
        detalle = (
            "Windows rechazó el acceso de escritura. "
            "Puede tratarse de permisos insuficientes "
            "o de una restricción del archivo."
        )
    else:
        detalle = (
            "No fue posible obtener acceso de escritura "
            "al archivo."
        )

    return DiagnosticoAccesoExcel(
        ruta=ruta,
        estado=EstadoAccesoExcel.ERROR_ACCESO,
        puede_escribir=False,
        lock_office_detectado=lock_office,
        detalle=detalle,
        winerror=winerror,
    )


def archivo_excel_ocupado(
    ruta_excel: str | Path,
) -> bool:
    return diagnosticar_acceso_excel(
        ruta_excel
    ).ocupado
