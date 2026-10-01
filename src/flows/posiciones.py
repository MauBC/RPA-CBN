from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

from playwright.sync_api import Page, Locator, expect

from src.excel.validators import OrdenCotizacion
from src.utils.rpa_errors import paso_rpa


MONEDAS_WEB = {
    "PEN": ["PEN - S/", "PEN"],
    "USD": ["USD - US$", "USD"],
}


@dataclass(frozen=True)
class DatosPosicionWeb:
    numero: int
    cantidad: Decimal
    especificacion: str
    moneda: str


def _normalizar_texto(texto: str) -> str:
    return " ".join(str(texto).strip().upper().split())


def _a_decimal(valor: Any) -> Decimal:
    if isinstance(valor, Decimal):
        return valor

    texto = str(valor).strip().replace(",", "")
    return Decimal(texto)


def _formatear_importe_web(valor: Any) -> str:
    """
    Formato para PrimeNG inputnumber con locale en-US.
    Ejemplo:
    1735      -> 1,735.00
    900       -> 900.00
    1735.5    -> 1,735.50
    """
    numero = _a_decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{numero:,.2f}"


def _normalizar_moneda(moneda: str) -> str:
    moneda = str(moneda).strip().upper()

    if moneda not in MONEDAS_WEB:
        raise ValueError(f"Moneda no soportada: {moneda}. Use PEN o USD.")

    return moneda



def _obtener_especificacion_posicion_simple(orden: OrdenCotizacion) -> str:
    """
    Para órdenes simples:
    - Ordenes.TEXTO se usa como nombre/cabecera de cotización.
    - Posiciones.TEXTO_BREVE, si existe, se usa como especificación de la posición.
    """
    if orden.posiciones_detalle:
        return orden.posiciones_detalle[0].texto_breve

    texto = str(orden.texto).strip()

    if len(texto) > 40:
        raise ValueError(
            f"La orden {orden.id_orden} tiene TEXTO de {len(texto)} caracteres. "
            "Ese TEXTO sirve como nombre de cotización, pero para la posición "
            "debe existir una fila en hoja Posiciones con TEXTO_BREVE de máximo 40 caracteres."
        )

    return texto

def _obtener_datos_posiciones(orden: OrdenCotizacion) -> list[DatosPosicionWeb]:
    moneda = _normalizar_moneda(orden.moneda)
    tipo_servicio = str(orden.tipo_servicio).strip().upper()

    if tipo_servicio == "B" and orden.posiciones != 1:
        raise ValueError(
            f"La orden {orden.id_orden} es de tipo BIEN y tiene "
            f"{orden.posiciones} posiciones. Por ahora BIEN admite exactamente 1 posición."
        )

    # Caso simple:
    # Ordenes.TEXTO se queda como nombre/cabecera de cotización.
    # Si existe hoja Posiciones para este ID_ORDEN, usamos TEXTO_BREVE como especificación.
    # Si no existe, usamos Ordenes.TEXTO solo si no supera 40 caracteres.
    if orden.posiciones == 1 and orden.ceco:
        especificacion = _obtener_especificacion_posicion_simple(orden)

        return [
            DatosPosicionWeb(
                numero=1,
                cantidad=orden.valor,
                especificacion=especificacion,
                moneda=moneda,
            )
        ]

    # Caso multiple:
    # Cada posicion usa su propio VALOR como cantidad.
    if not orden.posiciones_detalle:
        raise ValueError(
            f"La orden {orden.id_orden} requiere posiciones_detalle, "
            "pero no se encontraron datos en la hoja Posiciones."
        )

    datos: list[DatosPosicionWeb] = []

    for indice, posicion in enumerate(orden.posiciones_detalle, start=1):
        # SERVICIO: la cantidad representa el importe de la posición.
        # BIEN: temporalmente se cotiza una unidad y el importe se coloca
        # después como Valor unitario en la fase de la manito.
        cantidad_web = Decimal("1") if tipo_servicio == "B" else posicion.valor

        datos.append(
            DatosPosicionWeb(
                numero=indice,
                cantidad=cantidad_web,
                especificacion=posicion.texto_breve,
                moneda=moneda,
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

    return tabla.locator("tbody tr").filter(
        has=page.locator("button:has(.pi-pencil), button:has(.pi-check)")
    )


def esperar_posiciones_generadas(page: Page, cantidad_esperada: int) -> None:
    with paso_rpa("esperar posiciones generadas"):
        if cantidad_esperada < 1:
            raise ValueError("La cantidad esperada de posiciones debe ser mayor o igual a 1.")

        filas = _obtener_filas_posiciones(page)
        expect(filas.nth(cantidad_esperada - 1)).to_be_visible(timeout=20_000)

        cantidad_actual = filas.count()

        if cantidad_actual < cantidad_esperada:
            raise RuntimeError(
                f"Se esperaban {cantidad_esperada} posiciones, "
                f"pero solo se encontraron {cantidad_actual}."
            )

        print(f"Posiciones generadas detectadas: {cantidad_actual}")


def _obtener_fila_por_indice(page: Page, indice_cero: int) -> Locator:
    filas = _obtener_filas_posiciones(page)
    fila = filas.nth(indice_cero)
    expect(fila).to_be_visible(timeout=20_000)
    return fila


def _click_editar_fila(fila: Locator, numero_posicion: int) -> None:
    boton_editar = fila.locator("button:has(.pi-pencil), button[icon='pi pi-pencil']").first

    try:
        expect(boton_editar).to_be_visible(timeout=10_000)
    except Exception:
        boton_editar = fila.locator("button.p-button-info.p-button-outlined").first
        expect(boton_editar).to_be_visible(timeout=20_000)

    boton_editar.scroll_into_view_if_needed()
    boton_editar.click()

    print(f"Editando posicion {numero_posicion}...")


def _obtener_fila_en_edicion(page: Page) -> Locator:
    fila = page.locator(
        "xpath=//tr[.//button[.//span[contains(@class,'pi-check')]]]"
    ).last

    expect(fila).to_be_visible(timeout=25_000)
    return fila


def _set_text_input_value(input_locator: Locator, value: str) -> None:
    expect(input_locator).to_be_visible(timeout=20_000)

    input_locator.scroll_into_view_if_needed()
    input_locator.click(force=True)

    input_locator.evaluate(
        """
        (el, value) => {
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        }
        """,
        value,
    )


def _set_inputnumber_value(input_locator: Locator, value: str) -> None:
    """
    Para p-inputnumber hay que escribir como usuario.
    Si solo se cambia con evaluate, Angular/PrimeNG puede no actualizar aria-valuenow/modelo.
    """
    expect(input_locator).to_be_visible(timeout=20_000)

    input_locator.scroll_into_view_if_needed()
    input_locator.click(force=True)
    input_locator.press("Control+A")
    input_locator.press("Backspace")
    input_locator.type(value, delay=25)
    input_locator.press("Tab")


def _llenar_cantidad(tabla: Locator, fila_editando: Locator, cantidad: Decimal) -> None:
    cantidad_texto = _formatear_importe_web(cantidad)

    celda = _obtener_celda_por_header(tabla, fila_editando, "Cantidad")

    campo_cantidad = celda.locator("p-inputnumber input[role='spinbutton']").first

    try:
        expect(campo_cantidad).to_be_visible(timeout=25_000)
    except Exception:
        campo_cantidad = celda.locator("input[role='spinbutton'], input.p-inputnumber-input").first
        expect(campo_cantidad).to_be_visible(timeout=20_000)

    _set_inputnumber_value(campo_cantidad, cantidad_texto)

    print(f"Cantidad colocada: {cantidad_texto}")


def _llenar_especificacion(tabla: Locator, fila_editando: Locator, especificacion: str) -> None:
    celda = _obtener_celda_por_header(tabla, fila_editando, "Especificación")

    campo_especificacion = celda.locator("input.uppercase.invalidate-span").first

    try:
        expect(campo_especificacion).to_be_visible(timeout=25_000)
    except Exception:
        campo_especificacion = celda.locator(
            "input[class*='uppercase'][class*='invalidate-span'], input.p-inputtext"
        ).first
        expect(campo_especificacion).to_be_visible(timeout=20_000)

    _set_text_input_value(campo_especificacion, especificacion)

    print(f"Especificacion colocada: {especificacion}")


def _seleccionar_moneda(tabla: Locator, fila_editando: Locator, page: Page, moneda: str) -> None:
    moneda = _normalizar_moneda(moneda)

    celda = _obtener_celda_por_header(tabla, fila_editando, "Moneda")

    dropdown_moneda = celda.locator("p-dropdown").first
    expect(dropdown_moneda).to_be_visible(timeout=20_000)

    trigger = dropdown_moneda.locator("[role='button'], .p-dropdown-trigger").first
    expect(trigger).to_be_visible(timeout=20_000)

    trigger.click()

    ultimo_error: Exception | None = None

    for texto_opcion in MONEDAS_WEB[moneda]:
        try:
            opcion = page.get_by_role("option", name=texto_opcion, exact=True).last
            expect(opcion).to_be_visible(timeout=10_000)
            opcion.click()
            print(f"Moneda seleccionada: {texto_opcion}")
            return
        except Exception as error:
            ultimo_error = error

    for texto_opcion in MONEDAS_WEB[moneda]:
        try:
            opcion = page.locator(".p-dropdown-item").filter(has_text=texto_opcion).last
            expect(opcion).to_be_visible(timeout=10_000)
            opcion.click()
            print(f"Moneda seleccionada: {texto_opcion}")
            return
        except Exception as error:
            ultimo_error = error

    raise RuntimeError(
        f"No se pudo seleccionar moneda {moneda}. "
        f"Opciones intentadas: {MONEDAS_WEB[moneda]}"
    ) from ultimo_error


def _confirmar_edicion_fila(fila_editando: Locator, numero_posicion: int) -> None:
    boton_check = fila_editando.locator(
        "button:has(.pi-check), "
        "button[icon='pi pi-check'], "
        "button.p-button-success"
    ).first

    expect(boton_check).to_be_visible(timeout=20_000)

    boton_check.scroll_into_view_if_needed()
    boton_check.click()

    print(f"Posicion {numero_posicion} confirmada con check.")


def editar_una_posicion(page: Page, indice_cero: int, datos: DatosPosicionWeb) -> None:
    with paso_rpa(f"editar posicion {datos.numero}"):
        fila = _obtener_fila_por_indice(page, indice_cero)

        _click_editar_fila(fila, datos.numero)

        page.wait_for_timeout(700)

        tabla = _obtener_tabla_requerimiento(page)
        fila_editando = _obtener_fila_en_edicion(page)

        _llenar_cantidad(tabla, fila_editando, datos.cantidad)
        _llenar_especificacion(tabla, fila_editando, datos.especificacion)
        _seleccionar_moneda(tabla, fila_editando, page, datos.moneda)
        _confirmar_edicion_fila(fila_editando, datos.numero)

        page.wait_for_timeout(900)


def editar_posiciones_requerimiento(page: Page, orden: OrdenCotizacion) -> None:
    datos_posiciones = _obtener_datos_posiciones(orden)

    with paso_rpa("editar posiciones del requerimiento"):
        esperar_posiciones_generadas(page, len(datos_posiciones))

        for indice_cero, datos in enumerate(datos_posiciones):
            editar_una_posicion(page, indice_cero, datos)

        print("Todas las posiciones fueron editadas correctamente.")
