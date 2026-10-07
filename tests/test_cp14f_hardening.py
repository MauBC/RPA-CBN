from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from packaging.requirements import Requirement

from src.excel import excel_transaction
from src.excel.excel_transaction import (
    ExcelAccesoError,
    exigir_excel_disponible,
)
from src.excel.file_access import (
    EstadoAccesoExcel,
)


ROOT = Path(
    __file__
).resolve().parents[1]


def test_cp14f_precheck_winerror5_no_disfraza_permiso_como_lock(
    tmp_path,
    monkeypatch,
):
    excel = (
        tmp_path
        / "DATA.xlsx"
    )

    excel.write_bytes(
        b"x"
    )

    diagnostico = SimpleNamespace(
        ruta=excel,
        estado=(
            EstadoAccesoExcel.ERROR_ACCESO
        ),
        puede_escribir=False,
        ocupado=False,
        winerror=5,
        detalle=(
            "Acceso denegado por permisos."
        ),
    )

    monkeypatch.setattr(
        excel_transaction,
        "diagnosticar_acceso_excel",
        lambda _: diagnostico,
    )

    with pytest.raises(
        ExcelAccesoError
    ):
        exigir_excel_disponible(
            excel
        )


def test_cp14f_fuentes_one_drive_sin_mojibake():
    for relativa in (
        "src/excel/excel_transaction.py",
        "src/excel/pending_sync.py",
    ):
        texto = (
            ROOT
            / relativa
        ).read_text(
            encoding="utf-8"
        )

        assert "Ã" not in texto
        assert "Â" not in texto


def test_cp14f_requirements_es_utf8_y_parseable():
    ruta = (
        ROOT
        / "requirements.txt"
    )

    datos = ruta.read_bytes()

    assert not datos.startswith(
        b"\xff\xfe"
    )

    texto = datos.decode(
        "utf-8"
    )

    lineas = [
        linea.strip()
        for linea
        in texto.splitlines()
        if (
            linea.strip()
            and not linea.strip().startswith("#")
        )
    ]

    assert len(lineas) == 40

    for linea in lineas:
        Requirement(
            linea
        )
