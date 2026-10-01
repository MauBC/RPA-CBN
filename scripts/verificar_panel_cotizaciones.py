from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]

if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.browser.session import URL_SISTEMA, abrir_contexto
from src.browser.auth import esperar_login_manual
from src.excel.reader import leer_ordenes_excel
from src.flows.panel_cotizaciones import verificar_cotizacion_en_panel
from src.utils.debug import preparar_carpetas, guardar_evidencia_error


RUTA_EXCEL = ROOT_DIR / "data" / "DATA.xlsx"


def main() -> None:
    preparar_carpetas()

    playwright = None
    contexto = None
    page = None

    try:
        ordenes = leer_ordenes_excel(RUTA_EXCEL)

        if not ordenes:
            raise RuntimeError("No hay ordenes en DATA.xlsx.")

        orden = ordenes[0]

        print(f"Probando verificacion en panel para ID_ORDEN={orden.id_orden}")
        print(f"Nombre de envio: {orden.texto}")

        playwright, contexto, page = abrir_contexto()

        page.goto(
            URL_SISTEMA,
            wait_until="domcontentloaded",
            timeout=60_000,
        )

        esperar_login_manual(page)

        info = verificar_cotizacion_en_panel(page, orden)

        print("")
        print("Verificacion OK")
        print(f"N° de Cotizacion: {info.numero_cotizacion}")
        print(f"Estado CBN: {info.estado_cbn}")

    except Exception as error:
        print("")
        print("Error verificando Panel de cotizaciones.")
        print(error)

        rutas = guardar_evidencia_error(page, error, "verificar_panel_cotizaciones")

        if "log" in rutas:
            print(f"Log guardado en: {rutas['log']}")
        if "screenshot" in rutas:
            print(f"Screenshot guardado en: {rutas['screenshot']}")
        if "html" in rutas:
            print(f"HTML guardado en: {rutas['html']}")

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
