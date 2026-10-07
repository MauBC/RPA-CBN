from __future__ import annotations

from pathlib import Path
import inspect

from src.core import runner
from src.gui.main_window import (
    VentanaPrincipal,
    _texto_accion_reintento,
)


def test_cp15b_acciones_por_estado():
    assert (
        _texto_accion_reintento(
            "Error"
        )
        == "Reintentar"
    )

    assert (
        _texto_accion_reintento(
            "Ignorada"
        )
        == "Reabrir"
    )

    assert (
        _texto_accion_reintento(
            "Pendiente"
        )
        == ""
    )


def test_cp15b_runner_expone_ignoradas(
    tmp_path,
    monkeypatch,
):
    excel = tmp_path / "DATA.xlsx"
    excel.write_bytes(
        b"x"
    )

    monkeypatch.setattr(
        runner,
        "validar_archivo_excel",
        lambda ruta: Path(ruta),
    )

    monkeypatch.setattr(
        runner,
        "inspeccionar_ids_ordenes_fallidas",
        lambda _ruta: [],
    )

    monkeypatch.setattr(
        runner,
        "inspeccionar_ids_ordenes_ignoradas",
        lambda _ruta: [
            "3001",
            "3002",
        ],
    )

    monkeypatch.setattr(
        runner,
        "obtener_estado_journal",
        lambda _ruta: {
            "pendientes": [],
            "fallidas": [],
            "inflight": [],
        },
    )

    monkeypatch.setattr(
        runner,
        "crear_excel_trabajo_pendientes",
        lambda *_args, **_kwargs: (
            None,
            [],
        ),
    )

    resultado = (
        runner.validar_excel_sin_ejecutar(
            excel
        )
    )

    assert resultado[
        "ignored_count"
    ] == 2

    assert [
        item["id_orden"]
        for item
        in resultado["ignored_orders"]
    ] == [
        "3001",
        "3002",
    ]

    assert resultado[
        "error_count"
    ] == 0

    assert (
        "Órdenes ignoradas: 2"
        in resultado["message"]
    )


def test_cp15b_treeview_multiseleccion():
    codigo = inspect.getsource(
        VentanaPrincipal._crear_interfaz
    )

    assert (
        'selectmode="extended"'
        in codigo
    )

    assert (
        'text="Ignorar seleccionadas"'
        in codigo
    )

    assert (
        'text="Reabrir seleccionadas"'
        in codigo
    )


def test_cp15b_cambio_manual_no_abre_cbn():
    codigo = inspect.getsource(
        VentanaPrincipal._cambiar_estado_manual
    )

    assert (
        "ignorar_ordenes_manual"
        in codigo
    )

    assert (
        "reabrir_ordenes_manual"
        in codigo
    )

    assert "Thread(" in codigo

    assert (
        "ejecutar_rpa("
        not in codigo
    )


def test_cp15b_reabrir_desde_accion():
    codigo = inspect.getsource(
        VentanaPrincipal._manejar_click_tabla
    )

    assert (
        'accion == "Reabrir"'
        in codigo
    )

    assert (
        "_cambiar_estado_manual("
        in codigo
    )


def test_cp15b_validacion_muestra_ignoradas():
    codigo = inspect.getsource(
        VentanaPrincipal._mostrar_validacion
    )

    assert (
        '"ignored_orders"'
        in codigo
    )

    assert (
        '"Ignorada"'
        in codigo
    )

    assert (
        '"Cerrada manualmente"'
        in codigo
    )
