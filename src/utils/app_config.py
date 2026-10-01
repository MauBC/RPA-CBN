from __future__ import annotations

from pathlib import Path
from typing import Any
import json

from src.utils.app_paths import obtener_ruta_configuracion


CONFIG_DEFAULT: dict[str, Any] = {
    "ultimo_excel": "",
    "navegador": "msedge",
    "mostrar_navegador": True,
}


def cargar_configuracion() -> dict[str, Any]:
    ruta = obtener_ruta_configuracion()

    if not ruta.exists():
        return CONFIG_DEFAULT.copy()

    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))
    except Exception:
        return CONFIG_DEFAULT.copy()

    config = CONFIG_DEFAULT.copy()

    if isinstance(datos, dict):
        config.update(datos)

    return config


def guardar_configuracion(config: dict[str, Any]) -> Path:
    ruta = obtener_ruta_configuracion()
    ruta.parent.mkdir(parents=True, exist_ok=True)

    temporal = ruta.with_suffix(".tmp")
    temporal.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporal.replace(ruta)

    return ruta
