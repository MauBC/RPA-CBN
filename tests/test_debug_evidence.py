from __future__ import annotations

from pathlib import Path

from src.utils.debug import (
    guardar_evidencia_error,
)


class PaginaPrueba:
    def __init__(self):
        self.screenshots = []
        self.movimientos = []
        self.esperas = []

    def evaluate(
        self,
        script,
        argumento=None,
    ):
        if argumento is None:
            return {
                "top": 1000,
                "viewport": 600,
                "max_top": 3000,
            }

        self.movimientos.append(
            float(
                argumento
            )
        )

        return None

    def screenshot(
        self,
        *,
        path,
        full_page=False,
    ):
        ruta = Path(
            path
        )

        ruta.write_bytes(
            b"png"
        )

        self.screenshots.append(
            (
                ruta.name,
                bool(
                    full_page
                ),
            )
        )

    def wait_for_timeout(
        self,
        tiempo,
    ):
        self.esperas.append(
            tiempo
        )

    def content(self):
        return (
            "<html>"
            "<body>prueba</body>"
            "</html>"
        )


def test_evidencia_guarda_contexto_arriba_actual_abajo(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setattr(
        "src.utils.debug._directorio_error",
        lambda contexto, timestamp: (
            tmp_path
            / "error"
        ),
    )

    pagina = PaginaPrueba()

    try:
        raise RuntimeError(
            "fallo de prueba"
        )
    except RuntimeError as error:
        rutas = guardar_evidencia_error(
            pagina,
            error,
            "orden_1001_prueba",
        )

    nombres = {
        nombre
        for nombre, _
        in pagina.screenshots
    }

    assert (
        "captura_actual.png"
        in nombres
    )

    assert (
        "captura_arriba.png"
        in nombres
    )

    assert (
        "captura_abajo.png"
        in nombres
    )

    assert (
        "captura.png"
        in nombres
    )

    assert (
        "screenshot"
        in rutas
    )

    assert (
        "screenshot_actual"
        in rutas
    )

    assert (
        "screenshot_arriba"
        in rutas
    )

    assert (
        "screenshot_abajo"
        in rutas
    )


def test_evidencia_restaura_scroll_original(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setattr(
        "src.utils.debug._directorio_error",
        lambda contexto, timestamp: (
            tmp_path
            / "error"
        ),
    )

    pagina = PaginaPrueba()

    try:
        raise RuntimeError(
            "fallo"
        )
    except RuntimeError as error:
        guardar_evidencia_error(
            pagina,
            error,
            "prueba",
        )

    # top=1000, viewport=600
    # desplazamiento=420
    assert pagina.movimientos[0] == 580
    assert pagina.movimientos[1] == 1420

    # Siempre restaura el punto exacto del error.
    assert pagina.movimientos[-1] == 1000


def test_captura_completa_sigue_existiendo_para_compatibilidad(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setattr(
        "src.utils.debug._directorio_error",
        lambda contexto, timestamp: (
            tmp_path
            / "error"
        ),
    )

    pagina = PaginaPrueba()

    try:
        raise RuntimeError(
            "fallo"
        )
    except RuntimeError as error:
        rutas = guardar_evidencia_error(
            pagina,
            error,
            "prueba",
        )

    completas = [
        item
        for item
        in pagina.screenshots
        if item == (
            "captura.png",
            True,
        )
    ]

    assert completas == [
        (
            "captura.png",
            True,
        )
    ]

    assert (
        rutas[
            "screenshot"
        ].name
        == "captura.png"
    )
