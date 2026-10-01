from datetime import date
from pathlib import Path

from playwright.sync_api import Page, expect

from src.utils.rpa_errors import paso_rpa


def obtener_fecha_hoy_texto() -> str:
    return date.today().strftime("%d/%m/%Y")


def obtener_dia_hoy_texto() -> str:
    return str(date.today().day)


def obtener_bloque_fecha_cierre(page: Page):
    return page.locator(
        "xpath=//div[contains(@class,'field')][.//label[contains(normalize-space(),'Fecha de cierre')]]"
    ).first


def seleccionar_fecha_cierre_hoy(page: Page) -> str:
    fecha_texto = obtener_fecha_hoy_texto()
    dia_texto = obtener_dia_hoy_texto()

    with paso_rpa("seleccionar fecha de cierre"):
        bloque_fecha = obtener_bloque_fecha_cierre(page)
        expect(bloque_fecha).to_be_visible(timeout=20_000)

        campo_fecha = bloque_fecha.locator("input").first
        boton_calendario = bloque_fecha.locator("button[aria-label='Choose Date']").first

        expect(campo_fecha).to_be_visible(timeout=20_000)
        expect(boton_calendario).to_be_visible(timeout=20_000)

        boton_calendario.click()

        calendario = page.locator(".p-datepicker").last
        expect(calendario).to_be_visible(timeout=20_000)

        dia = calendario.locator(
            "xpath=.//td[not(contains(@class,'p-datepicker-other-month')) "
            "and not(contains(@class,'p-disabled'))]"
            f"//span[normalize-space()='{dia_texto}']"
        ).first

        expect(dia).to_be_visible(timeout=20_000)
        dia.click()

        try:
            expect(campo_fecha).to_have_value(fecha_texto, timeout=10_000)
        except Exception:
            campo_fecha.click()
            campo_fecha.fill(fecha_texto)
            campo_fecha.press("Tab")
            expect(campo_fecha).to_have_value(fecha_texto, timeout=10_000)

    return fecha_texto


def abrir_modal_adjuntar(page: Page) -> None:
    with paso_rpa("abrir modal adjuntar archivo"):
        boton_adjuntar = page.get_by_role("button", name="ADJUNTAR")
        expect(boton_adjuntar).to_be_visible(timeout=20_000)
        boton_adjuntar.click()

        input_archivo = page.locator("input#fileDropRef")
        expect(input_archivo).to_be_attached(timeout=20_000)


def subir_archivos(page: Page, rutas_archivos: list[Path]) -> None:
    with paso_rpa("subir archivos adjuntos"):
        if not rutas_archivos:
            print("No hay archivos adjuntos para subir. Se omite este paso.")
            return

        for ruta in rutas_archivos:
            if not ruta.exists():
                raise FileNotFoundError(f"No existe el archivo adjunto: {ruta}")

        input_archivo = page.locator("input#fileDropRef")
        expect(input_archivo).to_be_attached(timeout=20_000)

        input_archivo.set_input_files([str(ruta) for ruta in rutas_archivos])


def continuar_adjunto(page: Page) -> None:
    with paso_rpa("continuar despues de adjuntar archivo"):
        boton_continuar = page.get_by_role("button", name="CONTINUAR")
        expect(boton_continuar).to_be_visible(timeout=25_000)
        boton_continuar.click()


def seleccionar_fecha_y_adjuntar(page: Page, rutas_archivos: list[Path]) -> None:
    fecha = seleccionar_fecha_cierre_hoy(page)
    print(f"Fecha de cierre colocada: {fecha}")

    abrir_modal_adjuntar(page)
    subir_archivos(page, rutas_archivos)
    continuar_adjunto(page)
