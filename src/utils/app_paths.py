from __future__ import annotations

from datetime import datetime
from pathlib import Path
import os
import re
import sys
import threading
import uuid


_LOCK = threading.RLock()
_DIRECTORIO_EJECUCION: Path | None = None
_ESCRITURA_VALIDADA: set[Path] = set()


def _limpiar_nombre(texto: str) -> str:
    texto = re.sub(r"[^A-Za-z0-9._-]+", "_", str(texto or "").strip())
    texto = texto.strip("._-")
    return texto or "ejecucion"


def obtener_directorio_programa() -> Path:
    """
    Devuelve la carpeta portable de la aplicación.

    - Ejecutable PyInstaller: carpeta que contiene RPA_CBN.exe.
    - Código fuente: raíz del proyecto RPA_CBN.

    No depende del directorio actual de PowerShell ni del acceso directo.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    # app_paths.py está en <proyecto>/src/utils/app_paths.py
    return Path(__file__).resolve().parents[2]


def _resolver_ruta_portable(valor: str | Path) -> Path:
    ruta = Path(valor).expanduser()

    if not ruta.is_absolute():
        ruta = obtener_directorio_programa() / ruta

    return ruta.resolve()


def validar_permisos_escritura(ruta: str | Path | None = None) -> Path:
    """
    Comprueba que la aplicación pueda crear archivos sin permisos de administrador.
    """
    directorio = _resolver_ruta_portable(ruta or obtener_directorio_programa())
    directorio.mkdir(parents=True, exist_ok=True)

    with _LOCK:
        if directorio in _ESCRITURA_VALIDADA:
            return directorio

        prueba = directorio / f".rpa_cbn_write_test_{uuid.uuid4().hex}.tmp"

        try:
            prueba.write_text("ok", encoding="utf-8")
            prueba.unlink()
        except Exception as error:
            raise PermissionError(
                "RPA CBN no tiene permisos para escribir en su propia carpeta: "
                f"{directorio}. Mueva la carpeta completa a Documentos, Escritorio "
                "o cualquier ubicación dentro de C:\\Users\\SU_USUARIO y vuelva a intentar."
            ) from error

        _ESCRITURA_VALIDADA.add(directorio)

    return directorio


def obtener_directorio_app() -> Path:
    """
    Raíz de datos en modo portable.

    Por defecto es la misma carpeta del proyecto o del ejecutable. Se mantiene
    RPA_APP_HOME como opción técnica; si es relativa también se resuelve desde
    la carpeta de la aplicación.
    """
    personalizado = os.getenv("RPA_APP_HOME", "").strip()
    ruta = _resolver_ruta_portable(personalizado) if personalizado else obtener_directorio_programa()
    return validar_permisos_escritura(ruta)


def obtener_directorio_datos() -> Path:
    ruta = obtener_directorio_app() / "data"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def obtener_directorio_ejecuciones() -> Path:
    ruta = obtener_directorio_app() / "ejecuciones"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def crear_directorio_ejecucion(ruta_excel: str | Path) -> Path:
    ruta_excel = Path(ruta_excel)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_nombre = _limpiar_nombre(ruta_excel.stem)

    raiz = obtener_directorio_ejecuciones()
    candidato = raiz / f"{timestamp}_{base_nombre}"
    contador = 2

    while candidato.exists():
        candidato = raiz / f"{timestamp}_{base_nombre}_{contador}"
        contador += 1

    candidato.mkdir(parents=True, exist_ok=False)
    (candidato / "errores").mkdir(parents=True, exist_ok=True)
    (candidato / "temporal").mkdir(parents=True, exist_ok=True)

    return candidato.resolve()


def configurar_directorio_ejecucion(ruta: str | Path | None) -> None:
    global _DIRECTORIO_EJECUCION

    with _LOCK:
        if ruta is None:
            _DIRECTORIO_EJECUCION = None
            return

        directorio = Path(ruta).expanduser().resolve()
        directorio.mkdir(parents=True, exist_ok=True)
        _DIRECTORIO_EJECUCION = directorio


def obtener_directorio_ejecucion() -> Path | None:
    with _LOCK:
        return _DIRECTORIO_EJECUCION


def obtener_directorio_salida_compatibilidad() -> Path:
    """
    Durante una ejecución devuelve su carpeta única.
    Fuera del runner utiliza <carpeta_app>/outputs.
    """
    actual = obtener_directorio_ejecucion()

    if actual is not None:
        return actual

    ruta = obtener_directorio_app() / "outputs"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def obtener_directorio_logs_compatibilidad() -> Path:
    ruta = obtener_directorio_app() / "logs"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def obtener_directorio_temporal() -> Path | None:
    actual = obtener_directorio_ejecucion()

    if actual is None:
        return None

    ruta = actual / "temporal"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def obtener_directorio_perfil(
    navegador: str = "msedge",
    worker_id: str | int | None = None,
) -> Path:
    navegador = str(navegador or "msedge").strip().lower()
    navegador = "msedge" if navegador in {"edge", "microsoft edge"} else navegador
    navegador = "chrome" if navegador in {"google chrome"} else navegador
    navegador = _limpiar_nombre(navegador)

    if worker_id is None:
        worker_id = os.getenv("RPA_WORKER_ID", "").strip() or None

    sufijo = f"_worker_{worker_id}" if worker_id is not None else ""
    ruta = obtener_directorio_app() / "perfiles" / f"{navegador}{sufijo}"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta.resolve()


def obtener_ruta_configuracion() -> Path:
    ruta = obtener_directorio_app() / "configuracion" / "config.json"
    ruta.parent.mkdir(parents=True, exist_ok=True)
    return ruta
