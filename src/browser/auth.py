from __future__ import annotations

from playwright.sync_api import (
    Error as PlaywrightError,
    Page,
    TimeoutError as PlaywrightTimeoutError,
)


class NavegadorCerradoPorUsuarioError(RuntimeError):
    """El usuario cerró la ventana usada por el RPA."""


def _es_error_navegador_cerrado(error: Exception) -> bool:
    texto = str(error).lower()
    indicadores = (
        "target page, context or browser has been closed",
        "target closed",
        "page has been closed",
        "browser has been closed",
        "context has been closed",
    )
    return any(indicador in texto for indicador in indicadores)


def _asegurar_pagina_abierta(page: Page) -> None:
    try:
        cerrada = page.is_closed()
    except Exception as error:
        if _es_error_navegador_cerrado(error):
            raise NavegadorCerradoPorUsuarioError(
                "El navegador fue cerrado por el usuario."
            ) from error
        raise

    if cerrada:
        raise NavegadorCerradoPorUsuarioError(
            "El navegador fue cerrado por el usuario."
        )


def esta_en_login(page: Page) -> bool:
    """Detecta si la página actual está en la pantalla de login."""
    _asegurar_pagina_abierta(page)

    try:
        page.locator("#float-input-email").wait_for(timeout=2_000)
        return True
    except PlaywrightTimeoutError:
        return False
    except PlaywrightError as error:
        if _es_error_navegador_cerrado(error):
            raise NavegadorCerradoPorUsuarioError(
                "El navegador fue cerrado por el usuario."
            ) from error
        raise


def esta_dentro_del_sistema(page: Page) -> bool:
    """
    Detecta si el usuario ya está dentro del sistema.
    Primero revisa la URL, luego algunos elementos visibles.
    """
    _asegurar_pagina_abierta(page)

    try:
        url_actual = page.url.lower()
    except PlaywrightError as error:
        if _es_error_navegador_cerrado(error):
            raise NavegadorCerradoPorUsuarioError(
                "El navegador fue cerrado por el usuario."
            ) from error
        raise

    if "/home" in url_actual:
        return True

    posibles_elementos = [
        page.get_by_text("Panel de requerimientos", exact=False),
        page.get_by_text("SIDEBAR.MANAGEMENT_SHOPPING", exact=False),
        page.get_by_role("link", name="Panel de requerimientos"),
    ]

    for elemento in posibles_elementos:
        try:
            elemento.wait_for(timeout=2_000)
            return True
        except PlaywrightTimeoutError:
            pass
        except PlaywrightError as error:
            if _es_error_navegador_cerrado(error):
                raise NavegadorCerradoPorUsuarioError(
                    "El navegador fue cerrado por el usuario."
                ) from error
            raise

    return False


def esperar_login_manual(page: Page, timeout_minutos: int = 10) -> None:
    """
    Si el usuario no está logueado, espera a que haga login manualmente.
    Cuando detecta /home o algún elemento interno, continúa.
    """
    try:
        _asegurar_pagina_abierta(page)

        if esta_dentro_del_sistema(page):
            print("Sesión ya iniciada.")
            return

        if esta_en_login(page):
            print("Pantalla de login detectada.")
            print("Por favor, inicia sesión manualmente en el navegador abierto.")
        else:
            print("No se detectó claramente login ni sistema.")
            print("Esperando a que aparezca el sistema...")

        timeout_ms = timeout_minutos * 60 * 1000

        try:
            page.wait_for_url("**/home**", timeout=timeout_ms)
            print("Login detectado correctamente por URL /home. Continuando...")
            return
        except PlaywrightTimeoutError:
            pass

        _asegurar_pagina_abierta(page)

        try:
            page.get_by_text("Panel de requerimientos", exact=False).wait_for(
                timeout=10_000
            )
            print("Login detectado correctamente por elemento interno. Continuando...")
            return
        except PlaywrightTimeoutError:
            raise RuntimeError(
                f"No se detectó login exitoso después de {timeout_minutos} minutos."
            )

    except NavegadorCerradoPorUsuarioError:
        raise
    except PlaywrightError as error:
        if _es_error_navegador_cerrado(error):
            raise NavegadorCerradoPorUsuarioError(
                "El navegador fue cerrado por el usuario."
            ) from error
        raise
