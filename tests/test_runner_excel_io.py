from pathlib import Path
from types import SimpleNamespace
import inspect

from src.core import runner


def test_indexar_filas_pendientes():
    pendientes = [
        SimpleNamespace(
            id_orden="1001",
            fila_excel=2,
        ),
        SimpleNamespace(
            id_orden="1002",
            fila_excel=5,
        ),
    ]

    resultado = (
        runner._indexar_filas_pendientes(
            pendientes
        )
    )

    assert resultado == {
        "1001": 2,
        "1002": 5,
    }


def test_resultado_ok_usa_fila_cacheada_sin_abrir_excel(
    monkeypatch,
):
    def prohibido(*args, **kwargs):
        raise AssertionError(
            "No debe abrir Excel para buscar la fila."
        )

    monkeypatch.setattr(
        runner,
        "obtener_fila_excel_por_id",
        prohibido,
    )

    orden = SimpleNamespace(
        id_orden="1001"
    )

    resultado = (
        runner.construir_resultado_ok(
            1,
            orden,
            Path("DATA.xlsx"),
            fila_excel=27,
        )
    )

    assert (
        resultado["FILA_EXCEL"]
        == 27
    )


def test_resultado_error_usa_fila_cacheada_sin_abrir_excel(
    monkeypatch,
):
    def prohibido(*args, **kwargs):
        raise AssertionError(
            "No debe abrir Excel para buscar la fila."
        )

    monkeypatch.setattr(
        runner,
        "obtener_fila_excel_por_id",
        prohibido,
    )

    monkeypatch.setattr(
        runner,
        "construir_mensaje_usuario",
        lambda error, orden: "Error prueba",
    )

    orden = SimpleNamespace(
        id_orden="1001"
    )

    resultado = (
        runner.construir_resultado_error(
            1,
            orden,
            Path("DATA.xlsx"),
            RuntimeError("fallo"),
            {},
            fila_excel=27,
        )
    )

    assert (
        resultado["FILA_EXCEL"]
        == 27
    )


def test_resultado_mantiene_fallback_si_no_recibe_fila(
    monkeypatch,
):
    llamadas = {
        "total": 0,
    }

    def buscar(*args, **kwargs):
        llamadas["total"] += 1
        return 44

    monkeypatch.setattr(
        runner,
        "obtener_fila_excel_por_id",
        buscar,
    )

    orden = SimpleNamespace(
        id_orden="1001"
    )

    resultado = (
        runner.construir_resultado_ok(
            1,
            orden,
            Path("DATA.xlsx"),
        )
    )

    assert resultado["FILA_EXCEL"] == 44
    assert llamadas["total"] == 1


def test_ejecutar_rpa_usa_cache_de_filas():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    assert (
        "filas_excel_por_id"
        in codigo
    )

    assert (
        "fila_excel=fila_excel_orden"
        in codigo
    )

    # La búsqueda histórica queda solo como fallback en
    # _resolver_fila_excel_resultado, no en el loop principal.
    assert (
        "obtener_fila_excel_por_id(ruta_excel, orden.id_orden)"
        not in codigo
    )
