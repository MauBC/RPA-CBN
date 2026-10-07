from __future__ import annotations

from pathlib import Path
import inspect

from openpyxl import Workbook

from src.core import runner
from src.excel.state_manager import (
    inspeccionar_ids_ordenes_fallidas,
)
from src.gui.main_window import (
    VentanaPrincipal,
)


def _crear_excel(
    ruta: Path,
) -> Path:
    wb = Workbook()

    try:
        ws = wb.active
        ws.title = "Ordenes"

        ws.append(
            [
                "ID_ORDEN",
                "ESTADO_RPA",
            ]
        )

        ws.append(
            [
                "1001",
                2,
            ]
        )

        ws.append(
            [
                "1001",
                2,
            ]
        )

        ws.append(
            [
                "1002",
                0,
            ]
        )

        ws.append(
            [
                "1003",
                1,
            ]
        )

        wb.save(
            ruta
        )

    finally:
        wb.close()

    return ruta


def test_cp14c_detecta_ids_estado_2_sin_modificar_excel(
    tmp_path,
):
    excel = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    ids = (
        inspeccionar_ids_ordenes_fallidas(
            excel
        )
    )

    assert ids == [
        "1001"
    ]

    assert (
        inspeccionar_ids_ordenes_fallidas(
            excel
        )
        == [
            "1001"
        ]
    )


def test_cp14c_validacion_expone_error_reintentable(
    tmp_path,
    monkeypatch,
):
    excel = tmp_path / "DATA.xlsx"
    excel.write_bytes(b"x")

    monkeypatch.setattr(
        runner,
        "validar_archivo_excel",
        lambda ruta: Path(ruta),
    )

    monkeypatch.setattr(
        runner,
        "inspeccionar_ids_ordenes_fallidas",
        lambda _ruta: [
            "1001"
        ],
    )

    monkeypatch.setattr(
        runner,
        "inspeccionar_ids_ordenes_ignoradas",
        lambda _ruta: [],
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
        "pending_count"
    ] == 0

    assert resultado[
        "error_count"
    ] == 1

    error = resultado[
        "error_orders"
    ][0]

    assert error[
        "id_orden"
    ] == "1001"

    assert error[
        "blocked"
    ] is False


def test_cp14c_journal_bloquea_reproceso_desde_preview(
    tmp_path,
    monkeypatch,
):
    excel = tmp_path / "DATA.xlsx"
    excel.write_bytes(b"x")

    monkeypatch.setattr(
        runner,
        "validar_archivo_excel",
        lambda ruta: Path(ruta),
    )

    monkeypatch.setattr(
        runner,
        "inspeccionar_ids_ordenes_fallidas",
        lambda _ruta: [
            "1001"
        ],
    )

    monkeypatch.setattr(
        runner,
        "inspeccionar_ids_ordenes_ignoradas",
        lambda _ruta: [],
    )

    monkeypatch.setattr(
        runner,
        "obtener_estado_journal",
        lambda _ruta: {
            "pendientes": [
                {
                    "id_orden": "1001",
                }
            ],
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

    error = resultado[
        "error_orders"
    ][0]

    assert error[
        "blocked"
    ] is True

    assert (
        error[
            "code"
        ]
        == "PENDIENTE_SINCRONIZACION"
    )

    assert (
        "Sincronizar Excel"
        in error[
            "reason"
        ]
    )


def test_cp14c_gui_distingue_error_de_revision():
    codigo = inspect.getsource(
        VentanaPrincipal._mostrar_validacion
    )

    assert (
        '"error_orders"'
        in codigo
    )

    assert (
        'estado_fila = "Error"'
        in codigo
    )

    assert (
        'estado_fila = "Revisión"'
        in codigo
    )

    assert (
        '"Disponible para reintento"'
        in codigo
    )
