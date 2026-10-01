from dataclasses import dataclass
from datetime import date

from playwright.sync_api import Page, Locator, expect

from src.excel.validators import OrdenCotizacion
from src.utils.rpa_errors import paso_rpa


EMPRESA_DEFAULT = "RANSA COMERCIAL SAC"
SEDE_BIEN_DEFAULT = "4000 - RANSA CALLAO ARGENTINA"

TIPOS_IMPUTACION_WEB = {
    "K": {
        "opciones": ["CENTRO DE COSTO (K)"],
        "label_codigo": "Centro de costo",
        "nombre_error": "Centro de costo",
    },
    "H": {
        "opciones": ["CENTRO DE BENEFICIO (H)", "CENTRO DE BENEFICIOS (H)"],
        "label_codigo": "Centro de beneficios",
        "nombre_error": "Centro de beneficios",
    },
    "F": {
        "opciones": ["ORDEN INTERNA (F)"],
        "label_codigo": "Orden interna",
        "nombre_error": "Orden interna",
    },
    "I": {
        "opciones": ["ORDEN DE INVERSION (I)", "ORDEN DE INVERSIÓN (I)"],
        "label_codigo": "Orden de inversión",
        "nombre_error": "Orden de inversión",
    },
}


@dataclass(frozen=True)
class DatosAsignacion:
    numero: int
    cuenta: str
    codigo_imputacion: str
    tipo_imputacion: str


def _fecha_hoy_texto() -> str:
    return date.today().strftime("%d/%m/%Y")


def _obtener_tabla_requerimiento(page: Page) -> Locator:
    tabla = page.locator(".table-supplier-panel.p-datatable").last

    try:
        expect(tabla).to_be_visible(timeout=20_000)
        return tabla
    except Exception:
        tabla = page.locator("p-table").filter(has_text="Requerimiento").last
        expect(tabla).to_be_visible(timeout=20_000)
        return tabla


def _obtener_filas_posiciones(page: Page) -> Locator:
    tabla = _obtener_tabla_requerimiento(page)

    filas = tabla.locator("tbody tr.main-file")

    if filas.count() > 0:
        return filas

    return tabla.locator("tbody tr").filter(has=page.locator("p-tablecheckbox"))


def _obtener_fila_por_indice(page: Page, indice_cero: int) -> Locator:
    filas = _obtener_filas_posiciones(page)
    fila = filas.nth(indice_cero)
    expect(fila).to_be_visible(timeout=20_000)
    return fila


def _obtener_datos_asignacion(orden: OrdenCotizacion) -> list[DatosAsignacion]:
    tipo_imputacion = str(orden.tipo_imputacion).strip().upper()

    if tipo_imputacion not in TIPOS_IMPUTACION_WEB:
        raise ValueError(
            f"TIPO_IMPUTACION invalido: {tipo_imputacion}. Use K, H, F o I."
        )

    if orden.posiciones == 1 and orden.ceco:
        return [
            DatosAsignacion(
                numero=1,
                cuenta=orden.cuenta,
                codigo_imputacion=orden.ceco,
                tipo_imputacion=tipo_imputacion,
            )
        ]

    datos: list[DatosAsignacion] = []

    for indice, posicion in enumerate(orden.posiciones_detalle, start=1):
        datos.append(
            DatosAsignacion(
                numero=indice,
                cuenta=orden.cuenta,
                codigo_imputacion=posicion.ceco,
                tipo_imputacion=tipo_imputacion,
            )
        )

    return datos


def _input_checked(input_checkbox: Locator) -> bool:
    try:
        return bool(input_checkbox.evaluate("(el) => el.checked"))
    except Exception:
        return False


def _deseleccionar_filas_visibles(page: Page) -> None:
    filas = _obtener_filas_posiciones(page)
    total = filas.count()

    for indice in range(total):
        fila = filas.nth(indice)
        input_checkbox = fila.locator("p-tablecheckbox input[type='checkbox']").first

        try:
            if _input_checked(input_checkbox):
                caja = fila.locator("p-tablecheckbox .p-checkbox-box").first
                caja.click()
                page.wait_for_timeout(200)
        except Exception:
            pass


def _seleccionar_fila(page: Page, indice_cero: int, numero_posicion: int) -> None:
    with paso_rpa(f"seleccionar fila posicion {numero_posicion}"):
        _deseleccionar_filas_visibles(page)

        fila = _obtener_fila_por_indice(page, indice_cero)

        checkbox_box = fila.locator("p-tablecheckbox .p-checkbox-box").first
        expect(checkbox_box).to_be_visible(timeout=20_000)

        checkbox_box.scroll_into_view_if_needed()
        checkbox_box.click()

        page.wait_for_timeout(500)

        input_checkbox = fila.locator("p-tablecheckbox input[type='checkbox']").first

        if not _input_checked(input_checkbox):
            raise RuntimeError(f"No se pudo seleccionar la posicion {numero_posicion}.")

        print(f"Fila de posicion {numero_posicion} seleccionada.")


def _abrir_panel_asignar(page: Page, numero_posicion: int) -> Locator:
    with paso_rpa(f"abrir panel asignar posicion {numero_posicion}"):
        boton_asignar = page.locator(
            "button:has(.fa-layer-group), button:has(.fas.fa-layer-group)"
        ).first

        expect(boton_asignar).to_be_visible(timeout=20_000)
        expect(boton_asignar).to_be_enabled(timeout=20_000)

        boton_asignar.click()

        panel = page.get_by_role("complementary").last
        expect(panel).to_be_visible(timeout=25_000)
        expect(panel.get_by_text("DATOS DE REQUERIMIENTO", exact=False)).to_be_visible(timeout=25_000)

        return panel


def _obtener_bloque_por_label(panel: Locator, label: str) -> Locator:
    # IMPORTANTE:
    # Este era el bug. Antes buscaba literalmente {label}.
    xpath = (
        f".//div[contains(@class,'field')]"
        f"[.//label[contains(normalize-space(.), '{label}')]]"
    )

    bloque = panel.locator(f"xpath={xpath}").first

    expect(bloque).to_be_visible(timeout=20_000)
    bloque.scroll_into_view_if_needed()

    return bloque


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


def _llenar_searchbox_si_existe(page: Page, texto: str) -> None:
    try:
        searchbox = page.get_by_role("searchbox").last
        expect(searchbox).to_be_visible(timeout=8_000)
        searchbox.click()
        searchbox.press("Control+A")
        searchbox.fill(texto)
        page.wait_for_timeout(900)
    except Exception:
        pass


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
        opcion = page.get_by_role("option", name=texto, exact=True).last
        expect(opcion).to_be_visible(timeout=20_000)
        opcion.click()
        return
    except Exception as error:
        ultimo_error = error

    raise RuntimeError(
        f"No se encontro la opcion '{texto}' para el campo '{nombre_campo}'."
    ) from ultimo_error


def _seleccionar_dropdown_por_label(
    page: Page,
    panel: Locator,
    label: str,
    texto_opcion: str,
    buscar_texto: str | None = None,
) -> None:
    _abrir_dropdown_por_label(panel, label)

    if buscar_texto:
        _llenar_searchbox_si_existe(page, buscar_texto)

    _seleccionar_opcion_abierta(page, texto_opcion, label)
    print(f"{label}: {texto_opcion}")


def _seleccionar_dropdown_filtrado_por_codigo(
    page: Page,
    panel: Locator,
    label: str,
    codigo: str,
    nombre_error: str,
) -> None:
    _abrir_dropdown_por_label(panel, label)
    _llenar_searchbox_si_existe(page, codigo)

    ultimo_error: Exception | None = None

    try:
        panel_dropdown = _obtener_panel_dropdown(page)

        # Primero intenta opción que contenga el código.
        opcion_codigo = panel_dropdown.locator(".p-dropdown-item, [role='option']").filter(has_text=codigo).first

        try:
            expect(opcion_codigo).to_be_visible(timeout=8_000)
            texto = opcion_codigo.inner_text(timeout=2_000)
            opcion_codigo.click()
            print(f"{label}: {texto}")
            return
        except Exception as error:
            ultimo_error = error

        # Para Cuenta contable, a veces luego de filtrar por código aparece solo la descripción.
        primera_opcion = panel_dropdown.locator(".p-dropdown-item, [role='option']").first
        expect(primera_opcion).to_be_visible(timeout=10_000)

        texto = primera_opcion.inner_text(timeout=2_000).strip()

        if not texto:
            raise RuntimeError("La primera opcion esta vacia.")

        if "no results" in texto.lower() or "sin resultados" in texto.lower():
            raise RuntimeError(texto)

        primera_opcion.click()
        print(f"{label}: {texto}")
        return

    except Exception as error:
        ultimo_error = error

    raise RuntimeError(
        f"No se encontro {nombre_error}: {codigo}. "
        "Revisa que el codigo exista en la web o que este bien escrito en Excel."
    ) from ultimo_error


def _seleccionar_tipo_imputacion(
    page: Page,
    panel: Locator,
    tipo_imputacion: str,
) -> None:
    config = TIPOS_IMPUTACION_WEB[tipo_imputacion]

    _abrir_dropdown_por_label(panel, "Imputación")

    ultimo_error: Exception | None = None

    for opcion in config["opciones"]:
        try:
            _seleccionar_opcion_abierta(page, opcion, "Imputación")
            print(f"Imputacion: {opcion}")
            return
        except Exception as error:
            ultimo_error = error

    raise RuntimeError(
        f"No se pudo seleccionar TIPO_IMPUTACION={tipo_imputacion}. "
        f"Opciones intentadas: {config['opciones']}"
    ) from ultimo_error


def _llenar_fecha_entrega_hoy(panel: Locator) -> None:
    fecha = _fecha_hoy_texto()

    bloque = _obtener_bloque_por_label(panel, "Fecha de entrega")
    campo_fecha = bloque.locator("input").first

    expect(campo_fecha).to_be_visible(timeout=20_000)

    campo_fecha.click()
    campo_fecha.press("Control+A")
    campo_fecha.fill(fecha)
    campo_fecha.press("Tab")

    print(f"Fecha de entrega: {fecha}")


def _asignar_en_panel(
    page: Page,
    panel: Locator,
    datos: DatosAsignacion,
    tipo_servicio: str,
) -> None:
    config = TIPOS_IMPUTACION_WEB[datos.tipo_imputacion]
    tipo_servicio = str(tipo_servicio).strip().upper()

    _seleccionar_dropdown_por_label(
        page=page,
        panel=panel,
        label="Empresa",
        texto_opcion=EMPRESA_DEFAULT,
        buscar_texto=EMPRESA_DEFAULT,
    )

    if tipo_servicio == "B":
        _seleccionar_dropdown_por_label(
            page=page,
            panel=panel,
            label="Sede",
            texto_opcion=SEDE_BIEN_DEFAULT,
            buscar_texto="4000",
        )
        print(f"Sede para BIEN seleccionada: {SEDE_BIEN_DEFAULT}")

    _llenar_fecha_entrega_hoy(panel)

    _seleccionar_tipo_imputacion(
        page=page,
        panel=panel,
        tipo_imputacion=datos.tipo_imputacion,
    )

    _seleccionar_dropdown_filtrado_por_codigo(
        page=page,
        panel=panel,
        label="Cuenta contable",
        codigo=datos.cuenta,
        nombre_error="Cuenta contable",
    )

    _seleccionar_dropdown_filtrado_por_codigo(
        page=page,
        panel=panel,
        label=config["label_codigo"],
        codigo=datos.codigo_imputacion,
        nombre_error=config["nombre_error"],
    )

    boton_asignar = panel.get_by_role("button", name="ASIGNAR").first
    expect(boton_asignar).to_be_visible(timeout=20_000)
    expect(boton_asignar).to_be_enabled(timeout=20_000)

    boton_asignar.click()

    print(f"Asignacion aplicada a posicion {datos.numero}.")
    page.wait_for_timeout(1_200)


def asignar_imputacion_a_posiciones(page: Page, orden: OrdenCotizacion) -> None:
    datos_asignacion = _obtener_datos_asignacion(orden)

    with paso_rpa("asignar imputacion a posiciones"):
        for indice_cero, datos in enumerate(datos_asignacion):
            print(f"Asignando imputacion a posicion {datos.numero}...")

            _seleccionar_fila(page, indice_cero, datos.numero)
            panel = _abrir_panel_asignar(page, datos.numero)
            _asignar_en_panel(
                page=page,
                panel=panel,
                datos=datos,
                tipo_servicio=orden.tipo_servicio,
            )

        print("Todas las posiciones tienen imputacion asignada.")
