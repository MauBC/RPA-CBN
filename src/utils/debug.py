from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
import re
import traceback

from src.utils.app_paths import (
    obtener_directorio_ejecucion,
    obtener_directorio_salida_compatibilidad,
    obtener_directorio_logs_compatibilidad,
)


def _limpiar_contexto(contexto: str) -> str:
    texto = re.sub(r"[^A-Za-z0-9._-]+", "_", str(contexto or "error"))
    return texto.strip("._-") or "error"


def preparar_carpetas() -> None:
    ejecucion = obtener_directorio_ejecucion()

    if ejecucion is not None:
        (ejecucion / "errores").mkdir(parents=True, exist_ok=True)
        (ejecucion / "temporal").mkdir(parents=True, exist_ok=True)
        return

    salida = obtener_directorio_salida_compatibilidad()
    (salida / "screenshots").mkdir(parents=True, exist_ok=True)
    (salida / "html_debug").mkdir(parents=True, exist_ok=True)
    obtener_directorio_logs_compatibilidad()


def obtener_timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _directorio_error(contexto: str, timestamp: str) -> Path:
    ejecucion = obtener_directorio_ejecucion()
    contexto = _limpiar_contexto(contexto)

    if ejecucion is not None:
        ruta = ejecucion / "errores" / f"{contexto}_{timestamp}"
        ruta.mkdir(parents=True, exist_ok=True)
        return ruta

    # Compatibilidad para scripts antiguos ejecutados fuera del runner.
    ruta = obtener_directorio_salida_compatibilidad() / "errores" / f"{contexto}_{timestamp}"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def guardar_evidencia_error(
    page: Any | None,
    error: Exception,
    contexto: str = "error",
) -> dict[str, Path]:
    preparar_carpetas()

    timestamp = obtener_timestamp()
    directorio = _directorio_error(contexto, timestamp)
    rutas: dict[str, Path] = {}

    ruta_log = directorio / "error.log"
    ruta_log.write_text(
        "ERROR:\n"
        f"{repr(error)}\n\n"
        "TRACEBACK:\n"
        f"{traceback.format_exc()}\n",
        encoding="utf-8",
    )
    rutas["log"] = ruta_log

    if page is None:
        return rutas

    try:
        ruta_screenshot = directorio / "captura.png"
        page.screenshot(path=str(ruta_screenshot), full_page=True)
        rutas["screenshot"] = ruta_screenshot
    except Exception:
        pass

    try:
        ruta_html = directorio / "pagina.html"
        ruta_html.write_text(page.content(), encoding="utf-8")
        rutas["html"] = ruta_html
    except Exception:
        pass

    return rutas
