from dataclasses import dataclass
from pathlib import Path

from playwright.sync_api import Page, Locator, expect

from src.excel.validators import OrdenCotizacion
from src.utils.rpa_errors import paso_rpa


TIPOS_IMPUTACION_DETALLE_WEB = {
    "K": "CENTRO DE COSTO",
    "H": "CENTRO DE BENEFICIO",
    "F": "ORDEN INTERNA",
    "I": "ORDEN DE INVERSION",
}


@dataclass(frozen=True)
class DatosTemplateImputacion:
    numero: int
    ruta_template: Path
    tipo_imputacion: str
    tipo_distribucion: str


def _normalizar_texto(texto: str) -> str:
    return " ".join(str(texto).strip().upper().split())




def _tipo_distribucion_template(ruta_template: Path) -> str:
    """
    Regla:
    - Si el nombre contiene '_porcent_' => PORCENTAJE
    - Caso contrario => CANTIDAD
    """
    nombre = str(ruta_template.name).strip().lower()

    if "_porcent_" in nombre:
        return "PORCENTAJE"

    return "CANTIDAD"

def _obtener_templates(orden: OrdenCotizacion) -> list[DatosTemplateImputacion]:
    tipo = str(orden.tipo_imputacion).strip().upper()

    if tipo not in TIPOS_IMPUTACION_DETALLE_WEB:
        raise ValueError(f"TIPO_IMPUTACION invalido: {tipo}. Use K, H, F o I.")

    if orden.posiciones == 1 and orden.ceco:
        if not orden.template:
            return []

        return [
            DatosTemplateImputacion(
                numero=1,
                ruta_template=orden.template,
                tipo_imputacion=tipo,
                tipo_distribucion=_tipo_distribucion_template(orden.template),
            )
        ]

    datos: list[DatosTemplateImputacion] = []

    for indice, posicion in enumerate(orden.posiciones_detalle, start=1):
        if not posicion.template:
            continue

        datos.append(
            DatosTemplateImputacion(
                numero=indice,
                ruta_template=posicion.template,
                tipo_imputacion=tipo,
                tipo_distribucion=_tipo_distribucion_template(posicion.template),
            )
        )

    return datos


def _obtener_tabla_requerimiento(page: Page) -> Locator:
    tabla = page.locator(".table-supplier-panel.p-datatable").last

    try:
        expect(tabla).to_be_visible(timeout=20_000)
        return tabla
    except Exception:
        tabla = page.locator("p-table").filter(has_text="Requerimiento").last
        expect(tabla).to_be_visible(timeout=20_000)
        return tabla


def _obtener_indice_columna(tabla: Locator, nombre_columna: str) -> int:
    buscado = _normalizar_texto(nombre_columna)
    headers = tabla.locator("thead th")
    total = headers.count()

    for indice in range(total):
        try:
            texto_header = _normalizar_texto(headers.nth(indice).inner_text(timeout=1_000))
        except Exception:
            texto_header = ""

        if not texto_header:
            continue

        if texto_header == buscado or texto_header.startswith(buscado):
            return indice

    raise RuntimeError(f"No se encontro la columna '{nombre_columna}' en la tabla de requerimiento.")


def _obtener_celda_por_header(tabla: Locator, fila: Locator, nombre_columna: str) -> Locator:
    indice = _obtener_indice_columna(tabla, nombre_columna)
    celda = fila.locator("td").nth(indice)
    celda.scroll_into_view_if_needed()
    return celda


def _obtener_filas_posiciones(page: Page) -> Locator:
    tabla = _obtener_tabla_requerimiento(page)

    filas = tabla.locator("tbody tr.main-file")

    if filas.count() > 0:
        return filas

    return tabla.locator("tbody tr").filter(has=page.locator("p-tablecheckbox"))


def _obtener_fila_por_numero(page: Page, numero_posicion: int) -> Locator:
    filas = _obtener_filas_posiciones(page)
    fila = filas.nth(numero_posicion - 1)
    expect(fila).to_be_visible(timeout=20_000)
    return fila


def _obtener_boton_detalle(celda: Locator, numero_posicion: int) -> Locator:
    """
    En Detalle de imputación puede aparecer:
    - AGREGAR: cuando todavía no se cargó driver.
    - EDITAR: cuando ya existe driver cargado.
    """
    candidatos = [
        celda.get_by_role("button", name="AGREGAR").first,
        celda.get_by_role("button", name="EDITAR").first,
        celda.locator("button").filter(has_text="AGREGAR").first,
        celda.locator("button").filter(has_text="EDITAR").first,
        celda.locator("button").first,
    ]

    ultimo_error: Exception | None = None

    for boton in candidatos:
        try:
            expect(boton).to_be_visible(timeout=8_000)
            expect(boton).to_be_enabled(timeout=8_000)
            return boton
        except Exception as error:
            ultimo_error = error

    raise RuntimeError(
        f"No se encontro boton AGREGAR/EDITAR en Detalle de imputacion "
        f"para la posicion {numero_posicion}. "
        "Verifica que esa posicion ya tenga imputacion asignada."
    ) from ultimo_error


def _abrir_detalle_imputacion(page: Page, numero_posicion: int) -> Locator:
    with paso_rpa(f"abrir detalle imputacion posicion {numero_posicion}"):
        tabla = _obtener_tabla_requerimiento(page)
        fila = _obtener_fila_por_numero(page, numero_posicion)
        celda = _obtener_celda_por_header(tabla, fila, "Detalle de imputación")

        boton_detalle = _obtener_boton_detalle(celda, numero_posicion)

        texto_boton = ""
        try:
            texto_boton = boton_detalle.inner_text(timeout=1_000).strip()
        except Exception:
            pass

        boton_detalle.scroll_into_view_if_needed()
        boton_detalle.click()

        print(
            f"Detalle de imputacion abierto para posicion {numero_posicion}"
            + (f" con boton {texto_boton}." if texto_boton else ".")
        )

        panel = page.get_by_role("complementary").last
        expect(panel).to_be_visible(timeout=25_000)
        expect(panel.get_by_text("IMPUTACIONES MULTIPLES", exact=False)).to_be_visible(timeout=25_000)

        return panel


def _asegurar_formulario_agregar(page: Page, panel: Locator) -> Locator:
    """
    Si el panel abre en modo listado y aparece un boton AGREGAR interno, lo presiona.
    Si ya abre directo en formulario, no hace nada.
    """
    try:
        if panel.get_by_text("Agregar distribución por archivo", exact=False).first.is_visible():
            return panel
    except Exception:
        pass

    try:
        boton_agregar = panel.get_by_role("button", name="AGREGAR").last
        expect(boton_agregar).to_be_visible(timeout=8_000)
        boton_agregar.click()
        page.wait_for_timeout(700)
    except Exception:
        pass

    panel = page.get_by_role("complementary").last
    expect(panel.get_by_text("IMPUTACIONES MULTIPLES", exact=False)).to_be_visible(timeout=25_000)

    return panel


def _obtener_bloque_por_label(panel: Locator, label: str) -> Locator:
    xpath = (
        f".//div[contains(@class,'field')]"
        f"[.//label[contains(normalize-space(.), '{label}')]]"
    )

    bloque = panel.locator(f"xpath={xpath}").first
    expect(bloque).to_be_visible(timeout=20_000)
    bloque.scroll_into_view_if_needed()

    return bloque


def _activar_distribucion_por_archivo(panel: Locator) -> None:
    bloque = _obtener_bloque_por_label(panel, "Agregar distribución por archivo")

    switch_input = bloque.locator("p-inputswitch input[role='switch']").first
    slider = bloque.locator(".p-inputswitch-slider").first

    expect(slider).to_be_visible(timeout=20_000)

    try:
        checked = bool(
            switch_input.evaluate(
                "(el) => el.checked || el.getAttribute('aria-checked') === 'true'"
            )
        )
    except Exception:
        checked = False

    if not checked:
        slider.click()
        print("Distribucion por archivo activada.")
    else:
        print("Distribucion por archivo ya estaba activada.")


def _abrir_dropdown_por_label(panel: Locator, label: str) -> None:
    bloque = _obtener_bloque_por_label(panel, label)

    trigger = bloque.locator(
        ".p-dropdown-trigger, [role='button'][aria-label='dropdown trigger']"
    ).first

    expect(trigger).to_be_visible(timeout=20_000)
    trigger.click()


def _obtener_panel_dropdown(page: Page) -> Locator:
    panel_dropdown = page.locator(".p-dropdown-panel").last
    expect(panel_dropdown).to_be_visible(timeout=20_000)
    return panel_dropdown


def _seleccionar_opcion_abierta(page: Page, texto: str, nombre_campo: str) -> None:
    ultimo_error: Exception | None = None

    try:
        panel_dropdown = _obtener_panel_dropdown(page)
        opcion = panel_dropdown.locator(".p-dropdown-item, [role='option']").filter(has_text=texto).first
        expect(opcion).to_be_visible(timeout=10_000)
        opcion.click()
        return
    except Exception as error:
        ultimo_error = error

    try:
        opcion = page.get_by_role("option").filter(has_text=texto).first
        expect(opcion).to_be_visible(timeout=20_000)
        opcion.click()
        return
    except Exception as error:
        ultimo_error = error

    raise RuntimeError(
        f"No se encontro la opcion '{texto}' para el campo '{nombre_campo}'."
    ) from ultimo_error


def _seleccionar_dropdown_por_label(page: Page, panel: Locator, label: str, texto_opcion: str) -> None:
    _abrir_dropdown_por_label(panel, label)
    _seleccionar_opcion_abierta(page, texto_opcion, label)
    print(f"{label}: {texto_opcion}")


def _subir_archivo_template(panel: Locator, ruta_template: Path) -> None:
    input_file = panel.locator("p-fileupload input[type='file'], input[type='file'][accept*='.xlsx']").first
    expect(input_file).to_be_attached(timeout=25_000)

    input_file.set_input_files(str(ruta_template))

    print(f"TEMPLATE adjuntado: {ruta_template}")


def _guardar_imputacion_multiple(page: Page, panel: Locator, numero_posicion: int) -> None:
    boton_guardar = panel.get_by_role("button", name="GUARDAR").first

    try:
        expect(boton_guardar).to_be_visible(timeout=20_000)
    except Exception:
        boton_guardar = page.get_by_role("button", name="GUARDAR").last
        expect(boton_guardar).to_be_visible(timeout=20_000)

    expect(boton_guardar).to_be_enabled(timeout=20_000)
    boton_guardar.click()

    print(f"Imputacion multiple guardada para posicion {numero_posicion}.")
    page.wait_for_timeout(1_500)


def _subir_template_una_posicion(page: Page, datos: DatosTemplateImputacion) -> None:
    with paso_rpa(f"subir template imputacion posicion {datos.numero}"):
        panel = _abrir_detalle_imputacion(page, datos.numero)
        panel = _asegurar_formulario_agregar(page, panel)

        _activar_distribucion_por_archivo(panel)

        _seleccionar_dropdown_por_label(page, panel, "Tipo", datos.tipo_distribucion)

        tipo_web = TIPOS_IMPUTACION_DETALLE_WEB[datos.tipo_imputacion]
        _seleccionar_dropdown_por_label(page, panel, "Imputación", tipo_web)

        print(f"Tipo de distribucion detectado para template: {datos.tipo_distribucion}")
        _subir_archivo_template(panel, datos.ruta_template)

        _guardar_imputacion_multiple(page, panel, datos.numero)


def subir_templates_imputacion_multiple(page: Page, orden: OrdenCotizacion) -> None:
    templates = _obtener_templates(orden)

    if not templates:
        print("No hay TEMPLATE de imputacion multiple. Se omite este paso.")
        return

    with paso_rpa("subir templates imputacion multiple"):
        for datos in templates:
            print(f"Subiendo TEMPLATE de imputacion multiple para posicion {datos.numero}...")
            _subir_template_una_posicion(page, datos)

        print("Todos los TEMPLATE de imputacion multiple fueron procesados.")
