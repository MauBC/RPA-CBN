from __future__ import annotations

from pathlib import Path
import os
import sys

from src.core.runner import ejecutar_rpa
from src.utils.app_paths import obtener_directorio_datos


def _ruta_excel_consola() -> Path:
    valor = os.getenv("RPA_EXCEL", "").strip()

    if valor:
        ruta = Path(valor).expanduser()

        if not ruta.is_absolute():
            ruta = obtener_directorio_datos().parent / ruta

        return ruta.resolve()

    return (obtener_directorio_datos() / "DATA.xlsx").resolve()


RUTA_EXCEL = _ruta_excel_consola()
NAVEGADOR = os.getenv("RPA_NAVEGADOR", "msedge")


def main() -> int:
    resultado = ejecutar_rpa(
        ruta_excel=RUTA_EXCEL,
        navegador=NAVEGADOR,
    )

    if resultado.estado in {"COMPLETADA", "SIN_PENDIENTES", "CANCELADA"}:
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
