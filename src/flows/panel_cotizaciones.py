from __future__ import annotations

from dataclasses import dataclass
import re
import unicodedata

from playwright.sync_api import Page, Locator, expect

from src.excel.validators import OrdenCotizacion
from src.utils.rpa_errors import paso_rpa


URL_PANEL_COTIZACIONES = (
    "https://cbntech.net/home/company/f829dab0adfb3d2890abff0b49764979/panel-quotes"
)


@dataclass(frozen=True)
class InfoPanelCotizacion:
    numero_cotizacion: str
    estado_cbn: str


def _normalizar(texto: str) -> str:
    texto = str(texto or "")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.replace("°", " ").replace("º", " ")
    texto = re.sub(r"[^A-Z0-9 ]+", " ", texto.upper())
    texto = " ".join(texto.strip().split())

    return texto


def _ir_panel_cotizaciones(page: Page) -> None:
    page.goto(
        URL_PANEL_COTIZACIONES,
        wait_until="domcontentloaded",
        timeout=60_000,
    )

    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except Exception:
        pass

    expect(page.get_by_text("COLUMNAS", exact=False).first).to_be_visible(timeout=30_000)
    expect(page.get_by_text("Nombre de envío", exact=False).first).to_be_visible(timeout=30_000)


def _obtener_tabla_panel(page: Page) -> Locator:
    tabla = page.locator(".table-supplier-panel.p-datatable").last
    expect(tabla).to_be_visible(timeout=30_000)

    return tabla


def _leer_filas_visibles(tabla: Locator) -> list[dict]:
    filas = tabla.locator("tbody tr")

    try:
        datos = filas.evaluate_all(
            """
            (rows) => rows.map((tr, index) => {
                const cells = Array.from(tr.querySelectorAll('td')).map(td =>
                    (td.innerText || '').replace(/\\s+/g, ' ').trim()
                );

                return {
                    index,
                    nombre: cells[1] || '',
                    numero_cotizacion: cells[2] || '',
                    items: cells[3] || '',
                    estado: cells[4] || '',
                    texto_fila: (tr.innerText || '').replace(/\\s+/g, ' ').trim()
                };
            })
            """
        )
    except Exception:
        datos = []

    return datos


def _obtener_wrapper_tabla(tabla: Locator) -> Locator:
    wrapper = tabla.locator(".p-datatable-wrapper").first
    expect(wrapper).to_be_visible(timeout=20_000)

    return wrapper


def _scroll_tabla_al_inicio(tabla: Locator) -> None:
    try:
        wrapper = _obtener_wrapper_tabla(tabla)
        wrapper.evaluate("(el) => { el.scrollTop = 0; el.scrollLeft = 0; }")
    except Exception:
        pass


def _bajar_scroll_tabla(tabla: Locator) -> bool:
    """
    Devuelve True si pudo bajar más.
    Devuelve False si ya llegó al final.
    """
    try:
        wrapper = _obtener_wrapper_tabla(tabla)

        return bool(
            wrapper.evaluate(
                """
                (el) => {
                    const before = el.scrollTop;
                    const step = Math.max(250, Math.floor(el.clientHeight * 0.85));
                    el.scrollTop = el.scrollTop + step;

                    return el.scrollTop > before && (el.scrollTop + el.clientHeight) < (el.scrollHeight - 2);
                }
                """
            )
        )
    except Exception:
        return False


def _clave_fila(fila_data: dict) -> str:
    return "|".join(
        [
            str(fila_data.get("nombre", "")),
            str(fila_data.get("numero_cotizacion", "")),
            str(fila_data.get("estado", "")),
            str(fila_data.get("texto_fila", ""))[:80],
        ]
    )


def _buscar_fila_por_nombre_envio(page: Page, orden: OrdenCotizacion) -> tuple[Locator, dict]:
    tabla = _obtener_tabla_panel(page)
    buscado = _normalizar(orden.texto)

    _scroll_tabla_al_inicio(tabla)
    page.wait_for_timeout(500)

    nombres_vistos: list[str] = []
    claves_vistas: set[str] = set()

    # 25 intentos con scroll alcanza bastante más que 10-12 órdenes.
    # Sigue siendo rápido porque lee los td con JS directo.
    max_intentos_scroll = 25

    for intento in range(max_intentos_scroll):
        datos_filas = _leer_filas_visibles(tabla)

        for fila_data in datos_filas:
            clave = _clave_fila(fila_data)

            if clave in claves_vistas:
                continue

            claves_vistas.add(clave)

            nombre = str(fila_data.get("nombre", "")).strip()

            if nombre and len(nombres_vistos) < 15:
                nombres_vistos.append(nombre)

            nombre_norm = _normalizar(nombre)
            texto_norm = _normalizar(fila_data.get("texto_fila", ""))

            if nombre_norm == buscado or buscado in nombre_norm or buscado in texto_norm:
                fila = tabla.locator("tbody tr").nth(int(fila_data["index"]))
                expect(fila).to_be_visible(timeout=20_000)

                return fila, fila_data

        pudo_bajar = _bajar_scroll_tabla(tabla)
        page.wait_for_timeout(350)

        if not pudo_bajar:
            break

    raise RuntimeError(
        "No se encontró en Panel de cotizaciones la fila con "
        f"Nombre de envío: {orden.texto}. "
        f"Filas revisadas aprox: {len(claves_vistas)}. "
        f"Primeros nombres vistos: {nombres_vistos}"
    )


def _abrir_detalle_proceso(page: Page, fila: Locator) -> Locator:
    boton = fila.locator("td").first.locator("button").first

    expect(boton).to_be_visible(timeout=20_000)
    expect(boton).to_be_enabled(timeout=20_000)

    boton.scroll_into_view_if_needed()
    boton.click()

    panel = page.get_by_role("complementary").last

    try:
        expect(panel).to_be_visible(timeout=25_000)
        expect(panel.get_by_text("DETALLE DE PROCESO", exact=False)).to_be_visible(timeout=25_000)
        return panel
    except Exception:
        panel = page.locator(".p-sidebar-content").filter(has_text="DETALLE DE PROCESO").last
        expect(panel).to_be_visible(timeout=25_000)
        return panel


def _cerrar_panel_si_aparece(panel: Locator) -> None:
    try:
        boton_cerrar = panel.get_by_role("button", name="CERRAR").last
        expect(boton_cerrar).to_be_visible(timeout=2_000)
        boton_cerrar.click()
    except Exception:
        pass


def verificar_cotizacion_en_panel(page: Page, orden: OrdenCotizacion) -> InfoPanelCotizacion:
    with paso_rpa("verificar cotizacion en panel"):
        print("Ingresando a Panel de cotizaciones...")
        _ir_panel_cotizaciones(page)

        print(f"Buscando cotizacion por Nombre de envio: {orden.texto}")
        fila, fila_data = _buscar_fila_por_nombre_envio(page, orden)

        numero_cotizacion = str(fila_data.get("numero_cotizacion", "")).strip()
        estado_cbn = str(fila_data.get("estado", "")).strip()

        if not numero_cotizacion:
            raise RuntimeError(
                "Se encontró la fila, pero no se pudo leer el N° de Cotización en la tabla."
            )

        if not estado_cbn:
            raise RuntimeError(
                "Se encontró la fila, pero no se pudo leer el Estado en la tabla."
            )

        print(f"Fila encontrada: {fila_data.get('nombre', '')}")
        print(f"N° de Cotizacion detectado: {numero_cotizacion}")
        print(f"Estado CBN detectado: {estado_cbn}")

        # Abrimos el panel porque será parte de los siguientes pasos.
        print("Abriendo DETALLE DE PROCESO...")
        panel = _abrir_detalle_proceso(page, fila)
        print("DETALLE DE PROCESO abierto correctamente.")

        _cerrar_panel_si_aparece(panel)

        return InfoPanelCotizacion(
            numero_cotizacion=numero_cotizacion,
            estado_cbn=estado_cbn,
        )
