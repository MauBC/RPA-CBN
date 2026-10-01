from __future__ import annotations

from pathlib import Path
import os

from playwright.sync_api import sync_playwright

from src.utils.app_paths import obtener_directorio_app, obtener_directorio_perfil


URL_SISTEMA = "https://cbntech.net/"


def _env_bool(nombre: str, default: bool = False) -> bool:
    valor = os.getenv(nombre)

    if valor is None:
        return default

    return str(valor).strip().lower() in ("1", "true", "si", "sí", "yes", "y")


def _env_int(nombre: str, default: int) -> int:
    valor = os.getenv(nombre)

    if valor is None or not str(valor).strip():
        return default

    try:
        return int(str(valor).strip())
    except Exception:
        return default


def _normalizar_navegador(navegador: str | None) -> str:
    valor = str(
        navegador
        or os.getenv("RPA_NAVEGADOR", "")
        or "msedge"
    ).strip().lower()

    equivalencias = {
        "edge": "msedge",
        "microsoft edge": "msedge",
        "google chrome": "chrome",
    }

    return equivalencias.get(valor, valor)


def _carpeta_perfil(navegador: str) -> Path:
    perfil_env = os.getenv("RPA_PROFILE", "").strip()

    if perfil_env:
        ruta = Path(perfil_env).expanduser()

        if not ruta.is_absolute():
            ruta = obtener_directorio_app() / ruta

        ruta = ruta.resolve()
        ruta.mkdir(parents=True, exist_ok=True)
        return ruta

    return obtener_directorio_perfil(navegador=navegador)


def abrir_contexto(navegador: str | None = None):
    """
    Abre un perfil persistente independiente del perfil personal del usuario.

    Orden predeterminado:
    1. Microsoft Edge.
    2. Google Chrome como respaldo.

    Variables compatibles:
    - RPA_HEADLESS=1
    - RPA_PROFILE=<ruta>
    - RPA_SLOW_MO=0
    - RPA_NAVEGADOR=msedge|chrome
    """
    solicitado = _normalizar_navegador(navegador)

    candidatos = [solicitado]

    for respaldo in ("msedge", "chrome"):
        if respaldo not in candidatos:
            candidatos.append(respaldo)

    headless = _env_bool("RPA_HEADLESS", default=False)
    slow_mo_default = 0 if headless else 300
    slow_mo = _env_int("RPA_SLOW_MO", slow_mo_default)

    playwright = sync_playwright().start()
    errores: list[str] = []

    for canal in candidatos:
        carpeta_perfil = _carpeta_perfil(canal)

        print(
            f"Intentando navegador={canal} | headless={headless} | "
            f"perfil={carpeta_perfil} | slow_mo={slow_mo}"
        )

        try:
            contexto = playwright.chromium.launch_persistent_context(
                user_data_dir=str(carpeta_perfil),
                channel=canal,
                headless=headless,
                slow_mo=slow_mo,
                viewport={
                    "width": 1366,
                    "height": 768,
                },
            )

            page = contexto.pages[0] if contexto.pages else contexto.new_page()
            print(f"Navegador iniciado correctamente: {canal}")

            return playwright, contexto, page

        except Exception as error:
            errores.append(f"{canal}: {error}")
            print(f"No se pudo abrir {canal}. Se intentará otro navegador.")

    try:
        playwright.stop()
    except Exception:
        pass

    raise RuntimeError(
        "No se pudo abrir Microsoft Edge ni Google Chrome con Playwright.\n"
        + "\n".join(errores)
    )
