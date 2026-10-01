from playwright.sync_api import Page, expect

from src.excel.validators import OrdenCotizacion


def ir_a_panel_requerimientos(page: Page) -> None:
    # Entra al modulo de compras
    page.locator(".layout-tabmenu.relative > .layout-tabmenu-nav > li:nth-child(3) > a").click()

    # Abre gestion de compras
    page.get_by_role("listitem").filter(
        has_text="SIDEBAR.MANAGEMENT_SHOPPING"
    ).get_by_role("link").click()

    # Entra al panel
    page.get_by_role("link", name=" Panel de requerimientos").click()


def abrir_formulario_cotizacion(page: Page) -> None:
    page.get_by_role("button", name="COTIZACION").click()

    caja_texto = page.locator("app-quick-quote").get_by_role("textbox")
    expect(caja_texto).to_be_visible(timeout=20_000)


def llenar_texto_cotizacion(page: Page, texto: str) -> None:
    caja_texto = page.locator("app-quick-quote").get_by_role("textbox")
    expect(caja_texto).to_be_visible(timeout=20_000)

    caja_texto.click()
    caja_texto.fill(texto)


def iniciar_cotizacion_con_texto(page: Page, orden: OrdenCotizacion) -> None:
    ir_a_panel_requerimientos(page)
    abrir_formulario_cotizacion(page)
    llenar_texto_cotizacion(page, orden.texto)