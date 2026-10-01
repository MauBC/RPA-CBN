from __future__ import annotations

from dataclasses import dataclass
import re

from playwright.sync_api import Page, Locator, expect, TimeoutError as PlaywrightTimeoutError

from src.excel.validators import OrdenCotizacion
from src.flows.panel_cotizaciones import (
    _ir_panel_cotizaciones,
    _obtener_tabla_panel,
    _buscar_fila_por_nombre_envio,
    _abrir_detalle_proceso,
)
from src.utils.rpa_errors import paso_rpa


@dataclass(frozen=True)
class InfoCerrarCotizacion:
    numero_cotizacion: str
    estado_cbn: str
    url_duration: str
    mensaje: str


def _extraer_company_y_quote_id(url: str) -> tuple[str, str, str]:
    """
    Extrae:
    - origin: https://cbntech.net
    - company_id
    - quote_id

    Soporta:
    /home/company/<company>/panel-quotes/<quote_id>
    /home/company/<company>/panel-quotes/modify-quotes/<quote_id>/duration
    /home/company/<company>/panel-quotes/quote-client/<quote_id>/...
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

    patrones = [
        r"/panel-quotes/modify-quotes/([^/]+)",
        r"/panel-quotes/quote-client/([^/]+)",
        r"/panel-quotes/([^/?#]+)",
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


def _construir_url_duration(url_referencia: str) -> str:
    origin, company_id, quote_id = _extraer_company_y_quote_id(url_referencia)

    return (
        f"{origin}/home/company/{company_id}"
        f"/panel-quotes/modify-quotes/{quote_id}/duration"
    )


def _abrir_detalle_desde_panel(page: Page, orden: OrdenCotizacion):
    _ir_panel_cotizaciones(page)
    _obtener_tabla_panel(page)

    fila, fila_data = _buscar_fila_por_nombre_envio(page, orden)

    numero_cotizacion = str(fila_data.get("numero_cotizacion", "")).strip()
    estado_cbn = str(fila_data.get("estado", "")).strip()

    panel = _abrir_detalle_proceso(page, fila)

    return panel, numero_cotizacion, estado_cbn


def _abrir_modificar_si_no_hay_quote_id(page: Page, panel: Locator) -> Page:
    """
    Normalmente al abrir DETALLE DE PROCESO la URL ya queda:
    /panel-quotes/<quote_id>

    Si no se puede extraer quote_id, usamos MODIFICAR.
    Si CBN abre nueva pestaña, devolvemos esa nueva page.
    Si navega en la misma pestaña, devolvemos la misma page.
    """
    print("No se pudo obtener quote_id desde la URL actual. Intentando con botón MODIFICAR...")

    boton_modificar = panel.get_by_role("button", name="MODIFICAR", exact=True).last

    expect(boton_modificar).to_be_visible(timeout=20_000)
    expect(boton_modificar).to_be_enabled(timeout=20_000)

    try:
        with page.context.expect_page(timeout=15_000) as nueva_page_info:
            boton_modificar.click()

        nueva_page = nueva_page_info.value
        nueva_page.wait_for_load_state("domcontentloaded", timeout=30_000)

        try:
            nueva_page.wait_for_load_state("networkidle", timeout=30_000)
        except Exception:
            pass

        print(f"MODIFICAR abrió nueva pestaña: {nueva_page.url}")
        return nueva_page

    except PlaywrightTimeoutError:
        boton_modificar.click()

        try:
            page.wait_for_load_state("domcontentloaded", timeout=30_000)
        except Exception:
            pass

        try:
            page.wait_for_load_state("networkidle", timeout=30_000)
        except Exception:
            pass

        print(f"MODIFICAR navegó en la misma pestaña: {page.url}")
        return page


def _obtener_url_duration_desde_panel(page: Page, panel: Locator) -> tuple[Page, str]:
    try:
        url_duration = _construir_url_duration(page.url)
        return page, url_duration

    except Exception:
        page_modificar = _abrir_modificar_si_no_hay_quote_id(page, panel)
        url_duration = _construir_url_duration(page_modificar.url)

        return page_modificar, url_duration


def _ir_a_duration(page: Page, url_duration: str) -> None:
    print(f"Ingresando a DURACIÓN: {url_duration}")

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


def _confirmar_si_aparece(page: Page) -> None:
    """
    Por si CBN muestra un modal después de CERRAR.
    Si no aparece, continúa normal.
    """
    try:
        dialog = page.get_by_role("dialog").last
        expect(dialog).to_be_visible(timeout=5_000)
    except Exception:
        return

    botones_posibles = ["ACEPTAR", "CONFIRMAR", "SÍ", "SI", "GUARDAR", "CERRAR"]

    for nombre in botones_posibles:
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


def cerrar_cotizacion_desde_panel(page: Page, orden: OrdenCotizacion) -> InfoCerrarCotizacion:
    with paso_rpa("cerrar cotizacion"):
        print("Abriendo Panel de cotizaciones...")
        panel, numero_cotizacion, estado_cbn = _abrir_detalle_desde_panel(page, orden)

        print(f"N° de Cotización: {numero_cotizacion}")
        print(f"Estado CBN: {estado_cbn}")
        print(f"URL detalle actual: {page.url}")

        page_trabajo, url_duration = _obtener_url_duration_desde_panel(page, panel)

        _ir_a_duration(page_trabajo, url_duration)
        _click_cerrar_duration(page_trabajo)
        _confirmar_si_aparece(page_trabajo)

        try:
            page_trabajo.wait_for_load_state("networkidle", timeout=30_000)
        except Exception:
            pass

        page_trabajo.wait_for_timeout(2_000)

        # Si se abrió una pestaña nueva para modificar, la cerramos al final.
        if page_trabajo is not page:
            try:
                page_trabajo.close()
            except Exception:
                pass

        mensaje = "Cotización cerrada correctamente desde DURACIÓN."

        print(mensaje)

        return InfoCerrarCotizacion(
            numero_cotizacion=numero_cotizacion,
            estado_cbn=estado_cbn,
            url_duration=url_duration,
            mensaje=mensaje,
        )
