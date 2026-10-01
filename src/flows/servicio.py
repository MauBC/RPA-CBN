from playwright.sync_api import Page, Locator, expect

from src.excel.validators import OrdenCotizacion
from src.utils.rpa_errors import paso_rpa


TIPOS_SERVICIO_WEB = {
    "B": ["BIEN", "BIENES"],
    "S": ["SERVICIO", "SERVICIOS"],
}


def _obtener_panel_lateral(page: Page) -> Page | Locator:
    panel = page.get_by_role("complementary").last

    try:
        expect(panel).to_be_visible(timeout=10_000)
        return panel
    except Exception:
        return page


def _obtener_bloque_tipo(page: Page) -> Locator:
    """
    Ubica especificamente el campo cuyo label es Tipo.
    Esto evita confundirlo con Categoria, porque ambos combos dicen SELECCIONAR.
    """
    panel = _obtener_panel_lateral(page)

    bloque_tipo = panel.locator(
        "xpath=.//div[contains(@class,'field')][.//label[contains(normalize-space(.),'Tipo')]]"
    ).first

    expect(bloque_tipo).to_be_visible(timeout=20_000)
    return bloque_tipo


def _obtener_combo_tipo(page: Page) -> Locator:
    bloque_tipo = _obtener_bloque_tipo(page)

    combo = bloque_tipo.locator("[role='combobox']").first
    expect(combo).to_be_visible(timeout=20_000)

    return combo


def abrir_modal_busqueda_codigo(page: Page) -> None:
    with paso_rpa("abrir modal de busqueda de bien o servicio"):
        boton_busqueda = page.locator("button.p-button-help").first

        try:
            expect(boton_busqueda).to_be_visible(timeout=10_000)
        except Exception:
            boton_busqueda = page.locator(".p-button-help").first
            expect(boton_busqueda).to_be_visible(timeout=20_000)

        boton_busqueda.click()

        # Esperamos el campo Tipo, no cualquier combo SELECCIONAR.
        _obtener_combo_tipo(page)


def seleccionar_tipo_codigo(page: Page, tipo_servicio: str) -> None:
    with paso_rpa("seleccionar tipo bien o servicio"):
        tipo_servicio = tipo_servicio.strip().upper()

        if tipo_servicio not in TIPOS_SERVICIO_WEB:
            raise ValueError(
                f"TIPO_SERVICIO invalido: {tipo_servicio}. Use B para BIEN o S para SERVICIO."
            )

        combo_tipo = _obtener_combo_tipo(page)
        combo_tipo.click()

        ultimo_error: Exception | None = None

        for nombre_opcion in TIPOS_SERVICIO_WEB[tipo_servicio]:
            try:
                opcion = page.locator("[role='option']").filter(has_text=nombre_opcion).last
                expect(opcion).to_be_visible(timeout=10_000)
                opcion.click()
                print(f"Tipo seleccionado: {nombre_opcion}")
                return
            except Exception as error:
                ultimo_error = error

        raise RuntimeError(
            f"No se pudo seleccionar el tipo {tipo_servicio}. "
            f"Opciones intentadas: {TIPOS_SERVICIO_WEB[tipo_servicio]}"
        ) from ultimo_error


def buscar_codigo(page: Page, codigo: str) -> None:
    with paso_rpa("buscar codigo de bien o servicio"):
        panel = _obtener_panel_lateral(page)

        caja_busqueda = panel.get_by_role("textbox").last
        expect(caja_busqueda).to_be_visible(timeout=20_000)

        caja_busqueda.click()
        caja_busqueda.fill(codigo)

        boton_buscar = panel.get_by_role("button", name="BUSCAR").first
        expect(boton_buscar).to_be_visible(timeout=20_000)
        boton_buscar.click()


def _obtener_fila_resultado(page: Page, codigo: str) -> Locator:
    tabla = page.locator(".p-datatable").last

    try:
        expect(tabla).to_be_visible(timeout=20_000)
    except Exception:
        tabla = page.locator("table").last
        expect(tabla).to_be_visible(timeout=20_000)

    fila_codigo = tabla.locator("tbody tr").filter(has_text=codigo).first

    try:
        expect(fila_codigo).to_be_visible(timeout=20_000)
        return fila_codigo
    except Exception:
        # Fallback: normalmente la busqueda por codigo devuelve 1 resultado.
        primera_fila = tabla.locator("tbody tr").first
        expect(primera_fila).to_be_visible(timeout=20_000)
        return primera_fila


def agregar_codigo_en_posiciones(page: Page, codigo: str, cantidad_posiciones: int) -> None:
    with paso_rpa("agregar codigo a posiciones"):
        if cantidad_posiciones < 1:
            raise ValueError("La cantidad de posiciones debe ser mayor o igual a 1.")

        fila = _obtener_fila_resultado(page, codigo)

        boton_agregar = fila.locator("button:has(.fa-plus), button:has(.fas.fa-plus)").first

        try:
            expect(boton_agregar).to_be_visible(timeout=10_000)
        except Exception:
            boton_agregar = fila.locator("button").first
            expect(boton_agregar).to_be_visible(timeout=20_000)

        for indice in range(cantidad_posiciones):
            boton_agregar.click()
            page.wait_for_timeout(300)
            print(f"Codigo {codigo} agregado a posicion {indice + 1}/{cantidad_posiciones}.")


def cerrar_modal_busqueda_codigo(page: Page) -> None:
    with paso_rpa("cerrar modal de busqueda de bien o servicio"):
        boton_cerrar = page.get_by_role("button", name="CERRAR", exact=True).last
        expect(boton_cerrar).to_be_visible(timeout=20_000)
        boton_cerrar.click()


def seleccionar_codigo_bien_servicio(page: Page, orden: OrdenCotizacion) -> None:
    abrir_modal_busqueda_codigo(page)
    seleccionar_tipo_codigo(page, orden.tipo_servicio)
    buscar_codigo(page, orden.servicio)
    agregar_codigo_en_posiciones(page, orden.servicio, orden.posiciones)
    cerrar_modal_busqueda_codigo(page)
