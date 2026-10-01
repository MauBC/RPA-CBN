from playwright.sync_api import (
    Page,
    expect,
    TimeoutError as PlaywrightTimeoutError,
)

from src.utils.rpa_errors import paso_rpa


TIMEOUT_BOTONES_MS = 20_000
TIMEOUT_PROVEEDOR_MS = 30_000
MAX_INTENTOS_BUSQUEDA = 2


def abrir_modal_empresas(page: Page) -> None:
    with paso_rpa("abrir modal empresas"):
        boton_empresas = page.get_by_role("button", name="EMPRESAS")
        expect(boton_empresas).to_be_visible(timeout=TIMEOUT_BOTONES_MS)
        expect(boton_empresas).to_be_enabled(timeout=TIMEOUT_BOTONES_MS)
        boton_empresas.click(timeout=TIMEOUT_BOTONES_MS)

        buscador = page.get_by_role(
            "textbox",
            name="Buscar por empresa, código",
        )
        expect(buscador).to_be_visible(timeout=TIMEOUT_BOTONES_MS)
        expect(buscador).to_be_editable(timeout=TIMEOUT_BOTONES_MS)


def seleccionar_filtro_codigo_tributario(page: Page) -> None:
    with paso_rpa("seleccionar filtro codigo tributario"):
        dropdown = page.get_by_role(
            "button",
            name="dropdown trigger",
        )

        expect(dropdown).to_be_visible(timeout=TIMEOUT_BOTONES_MS)
        expect(dropdown).to_be_enabled(timeout=TIMEOUT_BOTONES_MS)
        dropdown.click(timeout=TIMEOUT_BOTONES_MS)

        opcion = (
            page
            .get_by_label("Código tributario")
            .get_by_text("Código tributario")
        )

        expect(opcion).to_be_visible(timeout=TIMEOUT_BOTONES_MS)
        opcion.click(timeout=TIMEOUT_BOTONES_MS)


def buscar_empresa_por_codigo(
    page: Page,
    proveedor: str,
) -> None:
    with paso_rpa("buscar empresa por codigo tributario"):
        buscador = page.get_by_role(
            "textbox",
            name="Buscar por empresa, código",
        )

        expect(buscador).to_be_visible(timeout=TIMEOUT_BOTONES_MS)
        expect(buscador).to_be_editable(timeout=TIMEOUT_BOTONES_MS)

        buscador.click(timeout=TIMEOUT_BOTONES_MS)
        buscador.fill("", timeout=TIMEOUT_BOTONES_MS)
        buscador.fill(proveedor, timeout=TIMEOUT_BOTONES_MS)

        boton_buscar = page.locator(".icon > .p-element").first

        expect(boton_buscar).to_be_visible(
            timeout=TIMEOUT_BOTONES_MS
        )
        expect(boton_buscar).to_be_enabled(
            timeout=TIMEOUT_BOTONES_MS
        )

        boton_buscar.click(timeout=TIMEOUT_BOTONES_MS)


def seleccionar_empresa_resultado(
    page: Page,
    proveedor: str,
) -> None:
    with paso_rpa("seleccionar proveedor en resultados"):
        fila = (
            page
            .locator("tbody tr")
            .filter(has_text=proveedor)
            .first
        )

        try:
            expect(fila).to_be_visible(
                timeout=TIMEOUT_PROVEEDOR_MS
            )

            boton_fila = fila.locator("button").first

            expect(boton_fila).to_be_visible(
                timeout=TIMEOUT_BOTONES_MS
            )
            expect(boton_fila).to_be_enabled(
                timeout=TIMEOUT_BOTONES_MS
            )

            boton_fila.click(timeout=TIMEOUT_BOTONES_MS)
            page.wait_for_timeout(1_000)
            return

        except Exception:
            pass

        boton_primer_resultado = (
            page
            .locator("tbody tr button")
            .first
        )

        expect(boton_primer_resultado).to_be_visible(
            timeout=TIMEOUT_PROVEEDOR_MS
        )
        expect(boton_primer_resultado).to_be_enabled(
            timeout=TIMEOUT_BOTONES_MS
        )

        boton_primer_resultado.click(
            timeout=TIMEOUT_BOTONES_MS
        )

        page.wait_for_timeout(1_000)


def _panel_representantes_visible(page: Page):
    """
    Después de elegir una empresa, a veces CBN abre un panel lateral:
    REPRESENTANTES Y CONTACTOS.
    Si aparece, hay que seleccionar representantes y GUARDAR.
    """
    panel = page.get_by_role("complementary").last

    try:
        expect(panel).to_be_visible(timeout=8_000)

        expect(
            panel
            .get_by_text(
                "REPRESENTANTES Y CONTACTOS",
                exact=False,
            )
            .first
        ).to_be_visible(timeout=8_000)

        return panel

    except Exception:
        return None


def _obtener_seccion_representantes(panel):
    xpath = (
        ".//div[contains(@class,'field')]"
        "[.//label[contains("
        "normalize-space(.), "
        "'Representantes disponibles'"
        ")]]"
    )

    seccion = panel.locator(f"xpath={xpath}").first

    expect(seccion).to_be_visible(
        timeout=TIMEOUT_BOTONES_MS
    )

    return seccion


def _checkbox_esta_marcado(caja) -> bool:
    try:
        clase = caja.get_attribute(
            "class",
            timeout=1_000,
        ) or ""

        if "p-highlight" in clase:
            return True

    except Exception:
        pass

    try:
        input_checkbox = caja.locator(
            "xpath=preceding-sibling::*//input | "
            "ancestor::div[contains(@class,'p-checkbox')]//input"
        ).first

        return bool(
            input_checkbox.evaluate("(el) => el.checked")
        )

    except Exception:
        return False


def _seleccionar_todos_representantes(panel) -> None:
    seccion = _obtener_seccion_representantes(panel)

    # Opción principal: checkbox de cabecera, el que marca todos.
    checkbox_header = seccion.locator(
        "p-tableheadercheckbox .p-checkbox-box"
    ).first

    try:
        expect(checkbox_header).to_be_visible(
            timeout=TIMEOUT_BOTONES_MS
        )

        if not _checkbox_esta_marcado(checkbox_header):
            checkbox_header.scroll_into_view_if_needed()
            checkbox_header.click(
                timeout=TIMEOUT_BOTONES_MS
            )

            print(
                "Representantes disponibles: "
                "checkbox de cabecera seleccionado."
            )
        else:
            print(
                "Representantes disponibles: "
                "ya estaban seleccionados."
            )

        return

    except Exception:
        pass

    # Fallback: marcar representantes fila por fila.
    checkboxes = seccion.locator(
        "tbody p-tablecheckbox .p-checkbox-box"
    )

    total = checkboxes.count()

    if total == 0:
        raise RuntimeError(
            "No se encontraron representantes disponibles "
            "para seleccionar."
        )

    for i in range(total):
        caja = checkboxes.nth(i)

        try:
            expect(caja).to_be_visible(timeout=8_000)

            if not _checkbox_esta_marcado(caja):
                caja.scroll_into_view_if_needed()
                caja.click(timeout=TIMEOUT_BOTONES_MS)

                page_wait = (
                    panel.page
                    if hasattr(panel, "page")
                    else None
                )

                if page_wait:
                    page_wait.wait_for_timeout(300)

        except Exception:
            continue

    print(
        "Representantes disponibles seleccionados "
        f"fila por fila: {total}"
    )


def _guardar_representantes(
    page: Page,
    panel,
) -> None:
    boton_guardar = panel.get_by_role(
        "button",
        name="GUARDAR",
        exact=True,
    ).first

    expect(boton_guardar).to_be_visible(
        timeout=TIMEOUT_BOTONES_MS
    )
    expect(boton_guardar).to_be_enabled(
        timeout=TIMEOUT_BOTONES_MS
    )

    boton_guardar.scroll_into_view_if_needed()
    boton_guardar.click(timeout=TIMEOUT_BOTONES_MS)

    print("Representantes/contactos guardados.")
    page.wait_for_timeout(1_500)


def manejar_representantes_contactos_si_aparece(
    page: Page,
) -> None:
    with paso_rpa(
        "manejar representantes y contactos proveedor"
    ):
        panel = _panel_representantes_visible(page)

        if panel is None:
            print(
                "No apareció panel de representantes/contactos. "
                "Se continúa normal."
            )
            return

        print(
            "Panel REPRESENTANTES Y CONTACTOS detectado."
        )

        _seleccionar_todos_representantes(panel)
        _guardar_representantes(page, panel)

        # Si el panel se cierra solo, perfecto.
        # Si no, se cerrará en cerrar_modal_empresas().
        try:
            expect(panel).not_to_be_visible(timeout=5_000)

            print(
                "Panel representantes/contactos cerrado "
                "después de guardar."
            )

        except Exception:
            print(
                "Panel representantes/contactos sigue visible; "
                "se cerrará con CERRAR."
            )


def cerrar_modal_empresas(page: Page) -> None:
    with paso_rpa("cerrar modal empresas"):
        buscador = page.get_by_role(
            "textbox",
            name="Buscar por empresa, código",
        )

        # Puede haber 2 botones CERRAR:
        # 1. Panel representantes/contactos
        # 2. Modal empresas
        #
        # CBN también puede cerrar o reconstruir estos paneles
        # automáticamente mientras intentamos hacer clic.
        for intento in range(1, 5):

            # Antes de hacer nada, comprobamos si CBN
            # ya cerró el modal automáticamente.
            try:
                if not buscador.is_visible(timeout=1_000):
                    print(
                        "Modal empresas ya está cerrado."
                    )
                    return

            except Exception:
                print(
                    "Modal empresas ya está cerrado."
                )
                return

            # IMPORTANTE:
            # Se vuelve a localizar CERRAR en cada intento.
            botones_cerrar = page.get_by_role(
                "button",
                name="CERRAR",
                exact=True,
            )

            total_botones = botones_cerrar.count()

            if total_botones == 0:
                print(
                    "No hay botón CERRAR disponible. "
                    "Esperando actualización de CBN..."
                )

                page.wait_for_timeout(700)
                continue

            boton_cerrar = None

            # Buscamos desde el último botón hacia atrás.
            # Normalmente el panel que está encima es el último.
            for indice in range(
                total_botones - 1,
                -1,
                -1,
            ):
                candidato = botones_cerrar.nth(indice)

                try:
                    if (
                        candidato.is_visible(
                            timeout=1_000
                        )
                        and candidato.is_enabled(
                            timeout=1_000
                        )
                    ):
                        boton_cerrar = candidato
                        break

                except Exception:
                    continue

            if boton_cerrar is None:
                print(
                    "Los botones CERRAR están cambiando. "
                    "Esperando actualización de CBN..."
                )

                page.wait_for_timeout(700)
                continue

            print(
                "Cerrando ventana/panel de empresas. "
                f"Intento {intento}..."
            )

            try:
                # Timeout corto intencional.
                #
                # Si este botón fue eliminado por CBN,
                # no tiene sentido esperar 20 segundos
                # por ese mismo elemento.
                boton_cerrar.click(
                    timeout=5_000
                )

            except PlaywrightTimeoutError:
                print(
                    "CBN actualizó el panel mientras se "
                    "intentaba cerrar. "
                    "Se volverá a localizar el botón."
                )

                page.wait_for_timeout(700)

                # Es posible que CBN haya terminado de
                # cerrar el modal durante ese tiempo.
                try:
                    if not buscador.is_visible(
                        timeout=1_000
                    ):
                        print(
                            "Modal empresas cerrado "
                            "automáticamente por CBN."
                        )
                        return

                except Exception:
                    print(
                        "Modal empresas cerrado "
                        "automáticamente por CBN."
                    )
                    return

                continue

            # Dejamos que termine la animación de cierre.
            page.wait_for_timeout(700)

        # Comprobación final.
        try:
            if not buscador.is_visible(
                timeout=1_500
            ):
                print(
                    "Modal empresas cerrado correctamente."
                )
                return

        except Exception:
            print(
                "Modal empresas cerrado correctamente."
            )
            return

        raise RuntimeError(
            "No se pudo cerrar el modal de empresas "
            "después de varios intentos."
        )


def seleccionar_proveedor_por_codigo(
    page: Page,
    proveedor: str,
) -> None:
    abrir_modal_empresas(page)
    seleccionar_filtro_codigo_tributario(page)

    ultimo_error: Exception | None = None

    for intento in range(
        1,
        MAX_INTENTOS_BUSQUEDA + 1,
    ):
        try:
            print(
                f"Buscando proveedor {proveedor} en CBN. "
                f"Intento {intento}/"
                f"{MAX_INTENTOS_BUSQUEDA}."
            )

            buscar_empresa_por_codigo(
                page,
                proveedor,
            )

            seleccionar_empresa_resultado(
                page,
                proveedor,
            )

            ultimo_error = None
            break

        except Exception as error:
            ultimo_error = error

            if intento < MAX_INTENTOS_BUSQUEDA:
                print(
                    "CBN demoró en devolver el proveedor. "
                    "Se volverá a intentar la búsqueda."
                )

                page.wait_for_timeout(2_000)

    if ultimo_error is not None:
        raise RuntimeError(
            "CBN demoró demasiado en buscar el proveedor "
            f"{proveedor} después de "
            f"{MAX_INTENTOS_BUSQUEDA} intentos."
        ) from ultimo_error

    manejar_representantes_contactos_si_aparece(page)
    cerrar_modal_empresas(page)