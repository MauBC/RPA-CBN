from playwright.sync_api import Page, expect

from src.utils.rpa_errors import paso_rpa


BOTONES_CONFIRMACION = [
    "ACEPTAR",
    "CONFIRMAR",
    "SÍ",
    "SI",
    "OK",
]


def _click_confirmacion_si_aparece(page: Page) -> None:
    for nombre in BOTONES_CONFIRMACION:
        try:
            boton = page.get_by_role("button", name=nombre, exact=True).last
            expect(boton).to_be_visible(timeout=2_000)
            expect(boton).to_be_enabled(timeout=2_000)
            boton.click()
            print(f"Confirmacion aceptada: {nombre}")
            page.wait_for_timeout(1_000)
            return
        except Exception:
            continue


def enviar_cotizacion(page: Page) -> None:
    with paso_rpa("enviar cotizacion"):
        boton_enviar = page.get_by_role("button", name="ENVIAR", exact=True).last

        expect(boton_enviar).to_be_visible(timeout=25_000)
        expect(boton_enviar).to_be_enabled(timeout=25_000)

        boton_enviar.scroll_into_view_if_needed()
        boton_enviar.click()

        print("Boton ENVIAR presionado.")

        _click_confirmacion_si_aparece(page)

        try:
            page.wait_for_load_state("networkidle", timeout=30_000)
        except Exception:
            pass

        page.wait_for_timeout(2_000)

        print("Cotizacion enviada o procesada por CBN.")
