from __future__ import annotations

from pathlib import Path
import inspect

import pandas as pd

from src.core import runner
from src.excel.state_manager import (
    crear_excel_trabajo_pendientes,
)
from src.gui.main_window import (
    VentanaPrincipal,
)


def _crear_excel_single(
    ruta: Path,
) -> Path:
    ordenes = pd.DataFrame(
        [
            {
                "ID_ORDEN": "1001",
                "ESTADO_RPA": 0,
                "TEXTO": "UNO",
            },
            {
                "ID_ORDEN": "1002",
                "ESTADO_RPA": 0,
                "TEXTO": "DOS",
            },
            {
                "ID_ORDEN": "1003",
                "ESTADO_RPA": 1,
                "TEXTO": "TRES",
            },
        ]
    )

    adjuntos = pd.DataFrame(
        [
            {
                "ID_ORDEN": "1001",
                "ARCHIVO": "uno.pdf",
            },
            {
                "ID_ORDEN": "1002",
                "ARCHIVO": "dos.pdf",
            },
        ]
    )

    with pd.ExcelWriter(
        ruta,
        engine="openpyxl",
    ) as writer:
        ordenes.to_excel(
            writer,
            sheet_name="Ordenes",
            index=False,
        )

        adjuntos.to_excel(
            writer,
            sheet_name="Adjuntos",
            index=False,
        )

    return ruta


def test_excel_trabajo_puede_filtrar_un_solo_id(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_WORKERS",
        "",
    )

    monkeypatch.setenv(
        "RPA_WORKER_ID",
        "",
    )

    excel = _crear_excel_single(
        tmp_path
        / "DATA.xlsx"
    )

    trabajo, pendientes = (
        crear_excel_trabajo_pendientes(
            excel,
            solo_lectura=True,
            ids_objetivo={
                "1002"
            },
        )
    )

    try:
        assert [
            p.id_orden
            for p in pendientes
        ] == [
            "1002"
        ]

        hojas = pd.read_excel(
            trabajo,
            sheet_name=None,
            dtype=object,
        )

        assert (
            hojas[
                "Ordenes"
            ][
                "ID_ORDEN"
            ].astype(
                str
            ).tolist()
            == [
                "1002"
            ]
        )

        assert (
            hojas[
                "Adjuntos"
            ][
                "ID_ORDEN"
            ].astype(
                str
            ).tolist()
            == [
                "1002"
            ]
        )

    finally:
        trabajo.unlink(
            missing_ok=True
        )


def test_runner_single_id_usa_filtro_sin_cambiar_modo_normal():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    assert (
        "solo_id_orden"
        in codigo
    )

    assert (
        "ids_objetivo="
        in codigo
    )

    assert (
        "if id_objetivo is None:"
        in codigo
    )


def test_gui_reintento_arranca_solo_id():
    codigo = inspect.getsource(
        VentanaPrincipal._iniciar_reintento_id
    )

    assert (
        "solo_id_orden=id_orden"
        in codigo
    )

    assert (
        "ejecutar_rpa("
        in codigo
    )


def test_gui_retry_ready_inicia_automaticamente_reintento():
    codigo = inspect.getsource(
        VentanaPrincipal._manejar_evento
    )

    assert (
        '"manual_retry_ready"'
        in codigo
    )

    assert (
        "self._iniciar_reintento_id("
        in codigo
    )
