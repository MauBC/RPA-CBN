from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT_DIR))

from src.browser.session import URL_SISTEMA, abrir_contexto
from src.browser.auth import esperar_login_manual
from src.utils.debug import preparar_carpetas, guardar_evidencia_error


def main() -> None:
    preparar_carpetas()

    playwright = None
    contexto = None
    page = None

    try:
        playwright, contexto, page = abrir_contexto()

        page.goto(
            URL_SISTEMA,
            wait_until="domcontentloaded",
            timeout=60_000,
        )

        esperar_login_manual(page)

        print("Inspector listo.")
        print("Usa Playwright Inspector para grabar o revisar selectores.")
        print("Cuando termines, copia los pasos generados y pasamelos.")

        page.pause()

    except Exception as error:
        print("Error en modo inspeccion.")
        print(error)

        rutas = guardar_evidencia_error(page, error, "error_inspeccion")

        if "log" in rutas:
            print(f"Log guardado en: {rutas['log']}")
        if "screenshot" in rutas:
            print(f"Screenshot guardado en: {rutas['screenshot']}")
        if "html" in rutas:
            print(f"HTML guardado en: {rutas['html']}")

        raise

    finally:
        if contexto is not None:
            try:
                contexto.close()
            except Exception:
                pass

        if playwright is not None:
            try:
                playwright.stop()
            except Exception:
                pass


if __name__ == "__main__":
    main()
