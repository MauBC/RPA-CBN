from decimal import Decimal

import pytest

from src.excel.reader import (
    _leer_filas_template,
    _template_es_porcentaje,
)

from src.excel.validators import ValidacionExcelError

from conftest import crear_template


def test_template_valido_se_lee_correctamente(tmp_path):
    ruta = crear_template(
        tmp_path / "template.xlsx",
        [
            ("100010", 25),
            ("100020", 75),
        ],
    )

    filas = _leer_filas_template(ruta)

    assert filas == [
        ("100010", Decimal("25.0000")),
        ("100020", Decimal("75.0000")),
    ]


def test_template_rechaza_cabeceras_invalidas(tmp_path):
    ruta = crear_template(
        tmp_path / "template.xlsx",
        [
            ("100010", 100),
        ],
        header_a="CODIGO_MALO",
    )

    with pytest.raises(
        ValidacionExcelError,
        match="no modifiques cabeceras",
    ):
        _leer_filas_template(ruta)


def test_template_rechaza_codigo_duplicado(tmp_path):
    ruta = crear_template(
        tmp_path / "template.xlsx",
        [
            ("100010", 50),
            ("100010", 50),
        ],
    )

    with pytest.raises(
        ValidacionExcelError,
        match="codigo duplicado",
    ):
        _leer_filas_template(ruta)


def test_template_porcentaje_se_detecta_por_nombre(tmp_path):
    ruta = crear_template(
        tmp_path / "plantilla_imputacion_porcent_google.xlsx",
        [
            ("100010", 100),
        ],
    )

    assert _template_es_porcentaje(ruta) is True


def test_template_normal_no_es_porcentaje(tmp_path):
    ruta = crear_template(
        tmp_path / "plantilla_imputacion_google.xlsx",
        [
            ("100010", 100),
        ],
    )

    assert _template_es_porcentaje(ruta) is False


def test_template_libera_archivo_despues_de_lectura(tmp_path):
    ruta = crear_template(
        tmp_path / "template.xlsx",
        [
            ("100010", 100),
        ],
    )

    filas = _leer_filas_template(ruta)

    assert filas

    # El archivo debe quedar liberado después de la lectura.
    ruta_renombrada = tmp_path / "template_renombrado.xlsx"

    ruta.rename(ruta_renombrada)

    assert ruta_renombrada.exists()
