from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
import re
import unicodedata

from playwright.sync_api import Page, Locator, expect

from src.excel.validators import OrdenCotizacion
from src.flows.panel_cotizaciones import (
    _ir_panel_cotizaciones,
    _obtener_tabla_panel,
    _buscar_fila_por_nombre_envio,
    _abrir_detalle_proceso,
)
from src.utils.rpa_errors import paso_rpa


@dataclass(frozen=True)
class InfoMandarCotizar:
    numero_cotizacion: str
    estado_cbn: str
    posiciones_cotizadas: int
    omitido: bool
    mensaje: str


MONEDA_WEB = {
    "PEN": "SOLES",
    "SOLES": "SOLES",
    "S": "SOLES",
    "USD": "DÓLAR ESTADOUNIDENSE",
    "DOLAR": "DÓLAR ESTADOUNIDENSE",
    "DOLARES": "DÓLAR ESTADOUNIDENSE",
    "DÓLAR": "DÓLAR ESTADOUNIDENSE",
    "DÓLARES": "DÓLAR ESTADOUNIDENSE",
    "US": "DÓLAR ESTADOUNIDENSE",
    "EUR": "EURO",
    "EURO": "EURO",
}


def _normalizar(texto: str) -> str:
    texto = str(texto or "")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Z0-9 ]+", " ", texto.upper())
    texto = " ".join(texto.strip().split())

    return texto


def _moneda_web_desde_orden(orden: OrdenCotizacion) -> str:
    moneda = _normalizar(getattr(orden, "moneda", ""))

    if moneda in MONEDA_WEB:
        return MONEDA_WEB[moneda]

    raise RuntimeError(
        f"Moneda no soportada para mandar a cotizar: {getattr(orden, 'moneda', '')}"
    )


def _abrir_detalle_proceso_desde_panel(page: Page, orden: OrdenCotizacion):
    _ir_panel_cotizaciones(page)

    _obtener_tabla_panel(page)
    fila, fila_data = _buscar_fila_por_nombre_envio(page, orden)

    numero_cotizacion = str(fila_data.get("numero_cotizacion", "")).strip()
    estado_cbn = str(fila_data.get("estado", "")).strip()

    if not numero_cotizacion:
        raise RuntimeError("No se pudo leer el N° de Cotización desde la fila del panel.")

    if not estado_cbn:
        raise RuntimeError("No se pudo leer el Estado desde la fila del panel.")

    panel = _abrir_detalle_proceso(page, fila)

    return panel, numero_cotizacion, estado_cbn


def _click_manitos_empresas_invitadas(page: Page, panel: Locator) -> None:
    print("Buscando botón de manitos en Empresas invitadas...")

    boton_icono = panel.locator(
        "i.fa-handshake, i.fas.fa-handshake, .fa-handshake"
    ).first

    expect(boton_icono).to_be_visible(timeout=20_000)

    boton_icono.scroll_into_view_if_needed()
    boton_icono.click(force=True)

    print("Botón de manitos presionado.")

    try:
        page.wait_for_load_state("domcontentloaded", timeout=30_000)
    except Exception:
        pass

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    page.wait_for_timeout(2_000)


def _seleccionar_dropdown(page: Page, dropdown: Locator, opcion: str, nombre_campo: str) -> None:
    combobox = dropdown.locator("[role='combobox']").first
    expect(combobox).to_be_visible(timeout=20_000)

    try:
        actual = combobox.inner_text(timeout=2_000)
    except Exception:
        actual = ""

    if _normalizar(opcion) in _normalizar(actual):
        print(f"{nombre_campo}: {opcion} ya estaba seleccionado.")
        return

    combobox.scroll_into_view_if_needed()
    combobox.click()

    opcion_locator = page.get_by_role("option", name=opcion, exact=True).last
    expect(opcion_locator).to_be_visible(timeout=20_000)
    opcion_locator.click()

    print(f"{nombre_campo}: {opcion}")


def _seleccionar_moneda(page: Page, orden: OrdenCotizacion) -> None:
    moneda_web = _moneda_web_desde_orden(orden)

    campo_moneda = page.locator(".field").filter(
        has_text=re.compile(r"\bMoneda\b", re.IGNORECASE)
    ).first

    expect(campo_moneda).to_be_visible(timeout=20_000)

    dropdown = campo_moneda.locator("p-dropdown").first
    expect(dropdown).to_be_visible(timeout=20_000)

    _seleccionar_dropdown(page, dropdown, moneda_web, "Moneda")


def _abrir_dialog_forma_pago(page: Page) -> Locator | None:
    campo_forma_pago = page.locator(".field").filter(
        has_text=re.compile(r"Forma\s+de\s+pago", re.IGNORECASE)
    ).first

    expect(campo_forma_pago).to_be_visible(timeout=20_000)

    texto_campo = campo_forma_pago.inner_text(timeout=3_000)

    if "FACTURA 60" in _normalizar(texto_campo):
        print("Forma de pago: FACTURA 60 DÍAS ya estaba configurada.")
        return None

    boton_editar = campo_forma_pago.locator(
        "button:has(.fa-pen), button:has(.fas.fa-pen), button:has(.p-button-icon)"
    ).last

    expect(boton_editar).to_be_visible(timeout=20_000)
    expect(boton_editar).to_be_enabled(timeout=20_000)

    boton_editar.scroll_into_view_if_needed()
    boton_editar.click()

    dialog = page.get_by_role("dialog").filter(has_text="FORMA DE PAGO").last
    expect(dialog).to_be_visible(timeout=20_000)

    return dialog


def _configurar_forma_pago(page: Page) -> None:
    dialog = _abrir_dialog_forma_pago(page)

    if dialog is None:
        return

    filas = dialog.locator("tbody tr")

    if filas.count() == 0:
        boton_plus = dialog.locator("button:has(.pi-plus)").first
        expect(boton_plus).to_be_visible(timeout=20_000)
        expect(boton_plus).to_be_enabled(timeout=20_000)
        boton_plus.click()

    fila = dialog.locator("tbody tr").first
    expect(fila).to_be_visible(timeout=20_000)

    dropdown = fila.locator("p-dropdown").first
    expect(dropdown).to_be_visible(timeout=20_000)

    _seleccionar_dropdown(page, dropdown, "FACTURA 60 DÍAS", "Forma de pago")

    input_peso = fila.locator("input[role='spinbutton']").first

    if input_peso.count() > 0:
        input_peso.click()
        input_peso.press("Control+A")
        input_peso.fill("100")
        input_peso.evaluate(
            """
            (el) => {
                el.value = '100';
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.blur();
            }
            """
        )

    boton_guardar = dialog.get_by_role("button", name="GUARDAR", exact=True).last
    expect(boton_guardar).to_be_visible(timeout=20_000)
    expect(boton_guardar).to_be_enabled(timeout=20_000)
    boton_guardar.click()

    try:
        expect(dialog).not_to_be_visible(timeout=20_000)
    except Exception:
        page.wait_for_timeout(2_000)

    print("Forma de pago configurada: FACTURA 60 DÍAS al 100%.")


def _obtener_filas_posiciones(page: Page) -> Locator:
    filas = page.locator("tr.main-file")

    expect(filas.first).to_be_visible(timeout=30_000)

    return filas


def _fila_posicion_ya_cotizada(fila: Locator) -> bool:
    """
    Si una posición ya se cotizó antes, alguna columna de cotización
    deja de estar en '-'. Esto ayuda cuando reintentamos desde checkpoint.
    """
    try:
        celdas = fila.locator("td").evaluate_all(
            """
            (tds) => tds.map(td => (td.innerText || '').replace(/\\s+/g, ' ').trim())
            """
        )
    except Exception:
        return False

    if len(celdas) < 9:
        return False

    # En la tabla:
    # 4 = Item cotizado
    # 5 = Cantidad cotizada
    # 6 = Valor c/dscto.
    # 8 = Monto cotizado
    for indice in (4, 5, 6, 8):
        valor = str(celdas[indice] or "").strip()

        if valor and valor != "-":
            return True

    return False


def _click_boton_cotizar_fila(fila: Locator, numero_posicion: int) -> None:
    """
    El botón COTIZAR puede aparecer visualmente solo con hover.
    Si Playwright lo ve hidden, se usa click por JavaScript.
    """
    fila.scroll_into_view_if_needed()

    try:
        fila.hover(timeout=5_000)
    except Exception:
        pass

    try:
        fila.locator("td").last.hover(timeout=3_000)
    except Exception:
        pass

    boton = fila.locator(
        "button[ptooltip='COTIZAR'], "
        "button:has(i.fa-arrow-right), "
        "button:has(.fa-arrow-right)"
    ).last

    expect(boton).to_be_attached(timeout=20_000)

    try:
        expect(boton).to_be_visible(timeout=3_000)
        expect(boton).to_be_enabled(timeout=3_000)
        boton.click()
    except Exception:
        print(
            f"Botón COTIZAR de posición {numero_posicion} está oculto; "
            "se hará click por JavaScript."
        )
        boton.evaluate(
            """
            (el) => {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
            }
            """
        )

    print(f"Botón COTIZAR presionado para posición {numero_posicion}.")


def _obtener_panel_cotizar(page: Page, numero_posicion: int) -> Locator:
    panel = page.locator(".p-sidebar-content").filter(has_text="Valor unitario").last

    try:
        expect(panel).to_be_visible(timeout=20_000)
        return panel
    except Exception:
        panel = page.get_by_role("complementary").filter(has_text="Valor unitario").last
        expect(panel).to_be_visible(timeout=20_000)
        return panel


def _formatear_valor_input(valor) -> str:
    numero = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(numero, "f")


def _setear_inputnumber_formcontrol(
    panel: Locator,
    formcontrolname: str,
    label_regex: str,
    valor,
    nombre_campo: str,
    numero_posicion: int,
) -> None:
    valor_texto = _formatear_valor_input(valor)

    input_numero = panel.locator(
        f"p-inputnumber[formcontrolname='{formcontrolname}'] "
        "input[role='spinbutton']"
    ).first

    if input_numero.count() == 0:
        campo = panel.locator(".field").filter(
            has_text=re.compile(label_regex, re.IGNORECASE)
        ).first

        expect(campo).to_be_visible(timeout=20_000)
        input_numero = campo.locator(
            "input[role='spinbutton'], input.p-inputnumber-input, input"
        ).first

    expect(input_numero).to_be_visible(timeout=20_000)

    input_numero.scroll_into_view_if_needed()
    input_numero.click()
    input_numero.press("Control+A")
    input_numero.press("Backspace")

    try:
        input_numero.fill(valor_texto)
    except Exception:
        input_numero.type(valor_texto, delay=25)

    input_numero.evaluate(
        """
        (el, value) => {
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        }
        """,
        valor_texto,
    )

    print(
        f"{nombre_campo} colocado en posición {numero_posicion}: {valor_texto}"
    )


def _setear_cantidad_bien(panel: Locator, numero_posicion: int) -> None:
    _setear_inputnumber_formcontrol(
        panel=panel,
        formcontrolname="quantity",
        label_regex=r"\bCantidad\b",
        valor=1,
        nombre_campo="Cantidad",
        numero_posicion=numero_posicion,
    )


def _setear_input_valor_unitario(
    panel: Locator,
    numero_posicion: int,
    valor_unitario,
) -> None:
    _setear_inputnumber_formcontrol(
        panel=panel,
        formcontrolname="unitValue",
        label_regex=r"Valor\s+unitario",
        valor=valor_unitario,
        nombre_campo="Valor unitario",
        numero_posicion=numero_posicion,
    )

def _click_guardar_panel_cotizar(panel: Locator, numero_posicion: int) -> None:
    botones = panel.get_by_role("button", name="GUARDAR", exact=True)

    total = botones.count()

    if total == 0:
        raise RuntimeError(
            f"No se encontró botón GUARDAR en panel de cotización para posición {numero_posicion}."
        )

    ultimo_error = None

    for indice in reversed(range(total)):
        boton = botones.nth(indice)

        try:
            expect(boton).to_be_visible(timeout=3_000)
            expect(boton).to_be_enabled(timeout=3_000)
            boton.scroll_into_view_if_needed()
            boton.click()
            return
        except Exception as error:
            ultimo_error = error

    try:
        botones.last.evaluate(
            """
            (el) => {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
            }
            """
        )
        return
    except Exception:
        raise RuntimeError(
            f"No se pudo presionar GUARDAR en panel de cotización "
            f"para posición {numero_posicion}. Último error: {ultimo_error}"
        )


def _colocar_valor_unitario_y_guardar(
    page: Page,
    numero_posicion: int,
    orden: OrdenCotizacion,
) -> None:
    panel = _obtener_panel_cotizar(page, numero_posicion)

    tipo_servicio = str(orden.tipo_servicio).strip().upper()

    if tipo_servicio == "B":
        # Regla temporal:
        # - una sola posición
        # - Cantidad = 1
        # - Valor unitario = suma/valor total de la orden
        _setear_cantidad_bien(panel, numero_posicion)
        _setear_input_valor_unitario(
            panel,
            numero_posicion,
            orden.valor,
        )
    else:
        # Conserva el comportamiento histórico para SERVICIO:
        # la cantidad ya contiene el importe y el valor unitario es 1.
        _setear_input_valor_unitario(
            panel,
            numero_posicion,
            1,
        )

    page.wait_for_timeout(500)

    _click_guardar_panel_cotizar(panel, numero_posicion)

    try:
        expect(panel).not_to_be_visible(timeout=20_000)
    except Exception:
        page.wait_for_timeout(2_000)

    print(f"Posición {numero_posicion} cotizada correctamente.")


def _cotizar_posiciones(page: Page, orden: OrdenCotizacion) -> int:
    filas = _obtener_filas_posiciones(page)
    total_filas = filas.count()

    if total_filas < orden.posiciones:
        raise RuntimeError(
            f"Se esperaban {orden.posiciones} posiciones, pero solo se detectaron {total_filas}."
        )

    posiciones_cotizadas = 0

    for indice in range(orden.posiciones):
        numero_posicion = indice + 1

        print(f"Cotizando posición {numero_posicion}/{orden.posiciones}...")

        filas = _obtener_filas_posiciones(page)
        fila = filas.nth(indice)

        if _fila_posicion_ya_cotizada(fila):
            print(f"Posición {numero_posicion} ya estaba cotizada. Se omite llenado.")
            posiciones_cotizadas += 1
            continue

        _click_boton_cotizar_fila(fila, numero_posicion)
        _colocar_valor_unitario_y_guardar(
            page,
            numero_posicion,
            orden,
        )

        posiciones_cotizadas += 1
        page.wait_for_timeout(800)

    return posiciones_cotizadas


def _click_enviar_principal_cotizacion(page: Page) -> None:
    print("Presionando ENVIAR principal de la cotización...")

    boton_enviar = page.get_by_role("button", name="ENVIAR", exact=True).last

    expect(boton_enviar).to_be_visible(timeout=20_000)
    expect(boton_enviar).to_be_enabled(timeout=20_000)

    boton_enviar.scroll_into_view_if_needed()
    boton_enviar.click()

    page.wait_for_timeout(1_500)


def _obtener_panel_usuarios_proveedor(page: Page) -> Locator:
    panel = page.locator(".p-sidebar-content").filter(
        has_text=re.compile(r"USUARIOS\s+DEL\s+PROVEEDOR", re.IGNORECASE)
    ).last

    try:
        expect(panel).to_be_visible(timeout=20_000)
        return panel
    except Exception:
        panel = page.get_by_role("complementary").filter(
            has_text=re.compile(r"USUARIOS\s+DEL\s+PROVEEDOR", re.IGNORECASE)
        ).last
        expect(panel).to_be_visible(timeout=20_000)
        return panel


def _guardar_usuarios_proveedor(page: Page) -> None:
    print("Guardando usuarios del proveedor...")

    panel = _obtener_panel_usuarios_proveedor(page)

    try:
        radio_checked = panel.locator(".p-radiobutton-checked").count()

        if radio_checked == 0:
            primer_radio = panel.locator("p-tableradiobutton, .p-radiobutton").first
            expect(primer_radio).to_be_visible(timeout=5_000)
            primer_radio.click()
    except Exception:
        pass

    boton_guardar = panel.get_by_role("button", name="GUARDAR", exact=True).last

    expect(boton_guardar).to_be_visible(timeout=20_000)
    expect(boton_guardar).to_be_enabled(timeout=20_000)

    boton_guardar.scroll_into_view_if_needed()
    boton_guardar.click()

    page.wait_for_timeout(1_500)


def _click_enviar_dialogo(page: Page, texto_dialogo: str, descripcion: str) -> None:
    print(f"Confirmando diálogo: {descripcion}...")

    dialog = page.get_by_role("dialog").filter(
        has_text=re.compile(texto_dialogo, re.IGNORECASE)
    ).last

    expect(dialog).to_be_visible(timeout=25_000)

    boton_enviar = dialog.get_by_role("button", name="ENVIAR", exact=True).last

    expect(boton_enviar).to_be_visible(timeout=20_000)
    expect(boton_enviar).to_be_enabled(timeout=20_000)

    boton_enviar.click()

    page.wait_for_timeout(1_500)


def _confirmar_envio_cotizacion(page: Page) -> None:
    _click_enviar_dialogo(page, r"VERIFICACI[ÓO]N|seguro", "VERIFICACIÓN")
    _click_enviar_dialogo(page, r"Confirmaci[oó]n|IMPORTANTE|descalificaci[oó]n", "Confirmación final")

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    page.wait_for_timeout(3_000)

    try:
        expect(page.get_by_text("DETALLE DE PROCESO", exact=False).first).to_be_visible(timeout=30_000)
        print("CBN volvió al DETALLE DE PROCESO.")
    except Exception:
        print("No se confirmó visualmente DETALLE DE PROCESO, pero el envío final fue presionado.")


def _enviar_cotizacion_final(page: Page) -> None:
    print("Iniciando envío final de la cotización...")

    _click_enviar_principal_cotizacion(page)
    _guardar_usuarios_proveedor(page)
    _confirmar_envio_cotizacion(page)

    print("Envío final de cotización completado.")



def _extraer_company_y_quote_id(url: str) -> tuple[str, str, str]:
    """
    Extrae:
    - origin: https://cbntech.net
    - company_id: f829...
    - quote_id: dc7...

    Soporta URLs tipo:
    /home/company/<company>/panel-quotes/quote-client/<quote_id>/...
    /home/company/<company>/panel-quotes/<quote_id>
    /home/company/<company>/panel-quotes/modify-quotes/<quote_id>/duration
    """
    url = str(url or "").strip()

    m_origin = re.match(r"^(https?://[^/]+)", url)
    if not m_origin:
        raise RuntimeError(f"No se pudo extraer origin desde URL: {url}")

    origin = m_origin.group(1)

    m_company = re.search(r"/home/company/([^/]+)", url)
    if not m_company:
        raise RuntimeError(f"No se pudo extraer company_id desde URL: {url}")

    company_id = m_company.group(1)

    patrones_quote = [
        r"/panel-quotes/quote-client/([^/]+)",
        r"/panel-quotes/modify-quotes/([^/]+)",
        r"/panel-quotes/([^/?#]+)",
    ]

    quote_id = ""

    for patron in patrones_quote:
        m_quote = re.search(patron, url)

        if m_quote:
            candidato = m_quote.group(1)

            if candidato not in ("quote-client", "modify-quotes"):
                quote_id = candidato
                break

    if not quote_id:
        raise RuntimeError(f"No se pudo extraer quote_id desde URL: {url}")

    return origin, company_id, quote_id


def _construir_url_duration(url_referencia: str) -> str:
    origin, company_id, quote_id = _extraer_company_y_quote_id(url_referencia)

    return (
        f"{origin}/home/company/{company_id}"
        f"/panel-quotes/modify-quotes/{quote_id}/duration"
    )


def _ir_a_duration_modificacion(page: Page, url_referencia: str) -> None:
    url_duration = _construir_url_duration(url_referencia)

    print(f"Ingresando directo a DURACIÓN: {url_duration}")

    page.goto(
        url_duration,
        wait_until="domcontentloaded",
        timeout=60_000,
    )

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    expect(page.get_by_text("DURACIÓN", exact=False).first).to_be_visible(timeout=30_000)


def _click_cerrar_duration(page: Page) -> None:
    print("Presionando CERRAR en DURACIÓN...")

    boton_cerrar = page.get_by_role("button", name="CERRAR", exact=True).last

    expect(boton_cerrar).to_be_visible(timeout=30_000)
    expect(boton_cerrar).to_be_enabled(timeout=30_000)

    boton_cerrar.scroll_into_view_if_needed()
    boton_cerrar.click()

    page.wait_for_timeout(1_500)


def _confirmar_cierre_si_aparece(page: Page) -> None:
    """
    Por si CBN muestra algún diálogo de confirmación al cerrar.
    No pasa nada si no aparece.
    """
    posibles_botones = ["ACEPTAR", "CONFIRMAR", "SÍ", "SI", "ENVIAR", "GUARDAR"]

    try:
        dialog = page.get_by_role("dialog").last
        expect(dialog).to_be_visible(timeout=5_000)
    except Exception:
        return

    for nombre in posibles_botones:
        try:
            boton = dialog.get_by_role("button", name=nombre, exact=True).last
            expect(boton).to_be_visible(timeout=2_000)
            expect(boton).to_be_enabled(timeout=2_000)
            print(f"Confirmando cierre con botón: {nombre}")
            boton.click()
            page.wait_for_timeout(1_500)
            return
        except Exception:
            continue


def _cerrar_orden_en_duration(page: Page, url_referencia: str) -> None:
    print("Cerrando orden para evitar modificaciones futuras...")

    _ir_a_duration_modificacion(page, url_referencia)
    _click_cerrar_duration(page)
    _confirmar_cierre_si_aparece(page)

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    page.wait_for_timeout(2_000)

    print("Orden cerrada desde DURACIÓN.")




def mandar_a_cotizar_orden(page: Page, orden: OrdenCotizacion) -> InfoMandarCotizar:
    with paso_rpa("mandar a cotizar"):
        print("Abriendo cotización desde Panel de cotizaciones...")
        panel, numero_cotizacion, estado_cbn = _abrir_detalle_proceso_desde_panel(page, orden)

        print(f"N° de Cotización: {numero_cotizacion}")
        print(f"Estado CBN: {estado_cbn}")

        _click_manitos_empresas_invitadas(page, panel)
        url_quote_client = page.url
        print(f"URL de cotización capturada: {url_quote_client}")

        print("Configurando moneda...")
        _seleccionar_moneda(page, orden)

        print("Configurando forma de pago...")
        _configurar_forma_pago(page)

        print("Cotizando posiciones...")
        posiciones_cotizadas = _cotizar_posiciones(page, orden)

        print("Enviando cotización final...")
        _enviar_cotizacion_final(page)

        mensaje = (
            f"Mandar a cotizar completado y enviado. "
            f"Posiciones cotizadas: {posiciones_cotizadas}/{orden.posiciones}."
        )

        print(mensaje)

        return InfoMandarCotizar(
            numero_cotizacion=numero_cotizacion,
            estado_cbn=estado_cbn,
            posiciones_cotizadas=posiciones_cotizadas,
            omitido=False,
            mensaje=mensaje,
        )
