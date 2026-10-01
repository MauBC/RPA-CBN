from __future__ import annotations

from dataclasses import dataclass
import re

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
class InfoAsignarComparativo:
    numero_cotizacion: str
    estado_cbn: str
    url_comparativo: str
    resumen: str
    mensaje: str


def _extraer_company_y_quote_id(url: str) -> tuple[str, str, str]:
    url = str(url or "").strip()

    m_origin = re.match(r"^(https?://[^/]+)", url)
    if not m_origin:
        raise RuntimeError(f"No se pudo extraer origin desde URL: {url}")

    origin = m_origin.group(1)

    m_company = re.search(r"/home/company/([^/]+)", url)
    if not m_company:
        raise RuntimeError(f"No se pudo extraer company_id desde URL: {url}")

    company_id = m_company.group(1)

    patrones = [
        r"/panel-quotes/modify-quotes/([^/]+)",
        r"/panel-quotes/quote-client/([^/]+)",
        r"/panel-quotes/([^/?#]+)",
        r"/comparative/([^/?#]+)",
    ]

    quote_id = ""

    for patron in patrones:
        m_quote = re.search(patron, url)

        if not m_quote:
            continue

        candidato = m_quote.group(1).strip()

        if candidato and candidato not in ("modify-quotes", "quote-client"):
            quote_id = candidato
            break

    if not quote_id:
        raise RuntimeError(f"No se pudo extraer quote_id desde URL: {url}")

    return origin, company_id, quote_id


def _construir_url_comparativo(url_referencia: str) -> str:
    origin, company_id, quote_id = _extraer_company_y_quote_id(url_referencia)

    return f"{origin}/home/company/{company_id}/comparative/{quote_id}"


def _abrir_detalle_desde_panel(page: Page, orden: OrdenCotizacion):
    _ir_panel_cotizaciones(page)
    _obtener_tabla_panel(page)

    fila, fila_data = _buscar_fila_por_nombre_envio(page, orden)

    numero_cotizacion = str(fila_data.get("numero_cotizacion", "")).strip()
    estado_cbn = str(fila_data.get("estado", "")).strip()

    panel = _abrir_detalle_proceso(page, fila)

    return panel, numero_cotizacion, estado_cbn


def _ir_a_comparativo(page: Page, url_comparativo: str) -> None:
    print(f"Ingresando a COMPARATIVO: {url_comparativo}")

    page.goto(
        url_comparativo,
        wait_until="domcontentloaded",
        timeout=60_000,
    )

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    page.wait_for_timeout(2_000)

    expect(page.get_by_role("button", name="ASIGNAR", exact=True).first).to_be_visible(
        timeout=30_000
    )


def _obtener_card_proveedor(page: Page) -> Locator:
    cards = page.locator(".company-options .child-card.card-info")

    total = cards.count()

    if total == 0:
        cards = page.locator(".company-options").filter(has_text=re.compile(r"\S"))
        total = cards.count()

    if total == 0:
        raise RuntimeError("No se encontró tarjeta de proveedor en COMPARATIVO.")

    for i in range(total):
        card = cards.nth(i)

        try:
            if card.is_visible(timeout=1_000):
                return card
        except Exception:
            continue

    return cards.first


def _seleccionar_checkbox_proveedor(page: Page) -> None:
    print("Seleccionando proveedor en comparativo...")

    card = _obtener_card_proveedor(page)

    checkbox_box = card.locator(".p-checkbox-box").first

    expect(checkbox_box).to_be_attached(timeout=20_000)

    try:
        resaltado = checkbox_box.get_attribute("data-p-highlight", timeout=2_000)
    except Exception:
        resaltado = None

    if str(resaltado).lower() == "true":
        print("Proveedor ya estaba seleccionado.")
        return

    try:
        expect(checkbox_box).to_be_visible(timeout=10_000)
        checkbox_box.scroll_into_view_if_needed()
        checkbox_box.click()
    except Exception:
        checkbox_box.evaluate(
            """
            (el) => {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
            }
            """
        )

    page.wait_for_timeout(800)

    print("Proveedor seleccionado.")


def _click_asignar(page: Page) -> None:
    print("Presionando ASIGNAR...")

    boton_asignar = page.get_by_role("button", name="ASIGNAR", exact=True).last

    expect(boton_asignar).to_be_visible(timeout=20_000)
    expect(boton_asignar).to_be_enabled(timeout=20_000)

    boton_asignar.scroll_into_view_if_needed()
    boton_asignar.click()

    page.wait_for_timeout(2_000)


def _confirmar_si_aparece(page: Page) -> None:
    try:
        dialog = page.get_by_role("dialog").last
        expect(dialog).to_be_visible(timeout=5_000)
    except Exception:
        return

    botones_posibles = [
        "ASIGNAR",
        "ACEPTAR",
        "CONFIRMAR",
        "SÍ",
        "SI",
        "GUARDAR",
        "ENVIAR",
    ]

    for nombre in botones_posibles:
        try:
            boton = dialog.get_by_role("button", name=nombre, exact=True).last
            expect(boton).to_be_visible(timeout=2_000)
            expect(boton).to_be_enabled(timeout=2_000)

            print(f"Confirmando comparativo con botón: {nombre}")
            boton.click()
            page.wait_for_timeout(1_500)
            return

        except Exception:
            continue


def _esperar_winner_assignment(page: Page, url_comparativo: str) -> None:
    print("Esperando pantalla winner-assignment...")

    try:
        page.wait_for_url(re.compile(r".*/winner-assignment.*"), timeout=30_000)
    except Exception:
        if "winner-assignment" not in page.url:
            page.goto(
                url_comparativo.rstrip("/") + "/winner-assignment",
                wait_until="domcontentloaded",
                timeout=60_000,
            )

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    expect(page.get_by_text("Empresas asignadas", exact=False).first).to_be_visible(timeout=30_000)

    print(f"Pantalla winner-assignment lista: {page.url}")


def _seleccionar_radio_empresa_asignada(page: Page) -> None:
    print("Seleccionando empresa asignada...")

    fila = page.locator("tbody tr").filter(has_text=re.compile(r"\S")).first
    expect(fila).to_be_visible(timeout=30_000)

    radio_box = fila.locator(".p-radiobutton-box").first
    expect(radio_box).to_be_attached(timeout=20_000)

    try:
        clase = radio_box.get_attribute("class", timeout=2_000) or ""
    except Exception:
        clase = ""

    if "p-highlight" in clase:
        print("Empresa asignada ya estaba seleccionada.")
        return

    try:
        radio_box.scroll_into_view_if_needed()
        radio_box.click()
    except Exception:
        radio_box.evaluate(
            """
            (el) => {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
            }
            """
        )

    page.wait_for_timeout(1_500)

    print("Empresa asignada seleccionada.")


def _seleccionar_razon_regularizacion(page: Page) -> None:
    print("Seleccionando razón: REGULARIZACION...")

    campo_razon = page.locator(".field").filter(
        has_text=re.compile(r"Raz[oó]n", re.IGNORECASE)
    ).first

    expect(campo_razon).to_be_visible(timeout=30_000)

    dropdown = campo_razon.locator("p-dropdown").first
    expect(dropdown).to_be_visible(timeout=20_000)

    combobox = dropdown.locator("[role='combobox']").first
    expect(combobox).to_be_visible(timeout=20_000)

    try:
        actual = combobox.inner_text(timeout=2_000)
    except Exception:
        actual = ""

    if "REGULARIZACION" in str(actual).upper():
        print("Razón REGULARIZACION ya estaba seleccionada.")
        return

    combobox.scroll_into_view_if_needed()
    combobox.click()

    opcion = page.get_by_role("option", name="REGULARIZACION", exact=True).last
    expect(opcion).to_be_visible(timeout=20_000)
    opcion.click()

    page.wait_for_timeout(1_500)

    print("Razón seleccionada: REGULARIZACION.")


def _click_finalizar(page: Page) -> None:
    print("Presionando FINALIZAR...")

    boton_finalizar = page.get_by_role("button", name="FINALIZAR", exact=True).last

    expect(boton_finalizar).to_be_visible(timeout=30_000)
    expect(boton_finalizar).to_be_enabled(timeout=30_000)

    boton_finalizar.scroll_into_view_if_needed()
    boton_finalizar.click()

    page.wait_for_timeout(1_000)


def _confirmar_popup_si(page: Page) -> None:
    print("Confirmando popup con SI...")

    popup = page.locator(".p-confirm-popup, [role='alertdialog']").last
    expect(popup).to_be_visible(timeout=20_000)

    boton_si = popup.get_by_role("button", name=re.compile(r"^S[IÍ]$", re.IGNORECASE)).last

    expect(boton_si).to_be_visible(timeout=20_000)
    expect(boton_si).to_be_enabled(timeout=20_000)

    boton_si.click()

    page.wait_for_timeout(2_000)


def _esperar_resultado_simulacion(page: Page) -> str:
    print("Esperando resultado de simulación...")

    try:
        page.wait_for_url(re.compile(r".*test=true.*"), timeout=30_000)
    except Exception:
        pass

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass

    page.wait_for_timeout(2_000)

    mensaje = page.locator(
        ".message-success, .message-important, .message-error, "
        ".message-warning, .message-danger, .message-info"
    ).first

    expect(mensaje).to_be_visible(timeout=30_000)

    texto = mensaje.inner_text(timeout=5_000)
    texto = re.sub(r"\s+", " ", texto).strip()

    if not texto:
        texto = "Sin texto de resumen visible."

    clase = mensaje.get_attribute("class", timeout=2_000) or ""
    clase_lower = clase.lower()
    texto_lower = texto.lower()

    es_error_cbn = (
        "message-important" in clase_lower
        or "message-error" in clase_lower
        or "message-danger" in clase_lower
        or texto_lower.startswith("error:")
        or "proveedor con vencimiento" in texto_lower
        or "todavía son erróneos" in texto_lower
        or "todavia son erroneos" in texto_lower
    )

    if es_error_cbn:
        resumen = f"ERROR CBN: {texto}"
        print(f"Resumen detectado: {resumen}")
        print("CBN devolvió un error de negocio al final. Se guarda en RESUMEN y no se relanza como error técnico.")
        return resumen

    print(f"Resumen detectado: {texto}")
    return texto

def _finalizar_winner_assignment(page: Page, url_comparativo: str) -> str:
    _esperar_winner_assignment(page, url_comparativo)
    _seleccionar_radio_empresa_asignada(page)
    _seleccionar_razon_regularizacion(page)
    _click_finalizar(page)
    _confirmar_popup_si(page)

    resumen = _esperar_resultado_simulacion(page)

    return resumen


def asignar_comparativo_desde_panel(page: Page, orden: OrdenCotizacion) -> InfoAsignarComparativo:
    with paso_rpa("asignar comparativo"):
        print("Abriendo Panel de cotizaciones...")
        panel, numero_cotizacion, estado_cbn = _abrir_detalle_desde_panel(page, orden)

        print(f"N° de Cotización: {numero_cotizacion}")
        print(f"Estado CBN: {estado_cbn}")
        print(f"URL detalle actual: {page.url}")

        url_comparativo = _construir_url_comparativo(page.url)

        _ir_a_comparativo(page, url_comparativo)
        _seleccionar_checkbox_proveedor(page)
        _click_asignar(page)
        _confirmar_si_aparece(page)

        resumen = _finalizar_winner_assignment(page, url_comparativo)

        if str(resumen).upper().startswith("ERROR CBN"):
            mensaje = f"Comparativo finalizado con observación CBN. Resumen: {resumen}"
        else:
            mensaje = f"Comparativo asignado y finalizado correctamente. Resumen: {resumen}"

        print(mensaje)

        return InfoAsignarComparativo(
            numero_cotizacion=numero_cotizacion,
            estado_cbn=estado_cbn,
            url_comparativo=url_comparativo,
            resumen=resumen,
            mensaje=mensaje,
        )
