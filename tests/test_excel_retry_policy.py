from __future__ import annotations

from pathlib import Path

from src.excel.excel_transaction import (
    ExcelArchivoOcupadoError,
)
from src.excel.retry_policy import (
    TipoErrorPersistencia,
    clasificar_error_persistencia,
    ejecutar_con_reintentos,
)


def test_clasifica_excel_ocupado_como_recuperable():
    error = ExcelArchivoOcupadoError(
        Path("DATA.xlsx"),
        "Excel abierto",
        32,
    )

    resultado = (
        clasificar_error_persistencia(
            error
        )
    )

    assert (
        resultado.tipo
        == TipoErrorPersistencia.RECUPERABLE
    )

    assert resultado.recuperable is True
    assert resultado.codigo == "EXCEL_OCUPADO"
    assert resultado.winerror == 32


def test_clasifica_value_error_como_no_recuperable():
    resultado = (
        clasificar_error_persistencia(
            ValueError(
                "ID_ORDEN inexistente"
            )
        )
    )

    assert (
        resultado.tipo
        == TipoErrorPersistencia.NO_RECUPERABLE
    )

    assert resultado.recuperable is False

    assert (
        resultado.codigo
        == "DATOS_EXCEL_INVALIDOS"
    )


def test_reintenta_error_temporal_y_luego_aplica(
    monkeypatch,
):
    monkeypatch.setattr(
        "src.excel.retry_policy.time.sleep",
        lambda _: None,
    )

    llamadas = {
        "total": 0,
    }

    def operacion():
        llamadas["total"] += 1

        if llamadas["total"] < 3:
            raise ExcelArchivoOcupadoError(
                "DATA.xlsx",
                "ocupado",
                32,
            )

        return "OK"

    resultado = ejecutar_con_reintentos(
        operacion
    )

    assert resultado.exito is True
    assert resultado.intentos == 3
    assert resultado.valor == "OK"
    assert llamadas["total"] == 3


def test_error_logico_no_se_reintenta(
    monkeypatch,
):
    sleeps = []

    monkeypatch.setattr(
        "src.excel.retry_policy.time.sleep",
        lambda segundos: sleeps.append(
            segundos
        ),
    )

    llamadas = {
        "total": 0,
    }

    def operacion():
        llamadas["total"] += 1

        raise ValueError(
            "No existe ID_ORDEN=999"
        )

    resultado = ejecutar_con_reintentos(
        operacion
    )

    assert resultado.exito is False
    assert resultado.intentos == 1
    assert llamadas["total"] == 1
    assert sleeps == []

    assert (
        resultado.clasificacion
        is not None
    )

    assert (
        resultado.clasificacion.recuperable
        is False
    )


def test_cambio_concurrente_excel_es_recuperable():
    from src.excel.excel_transaction import (
        ExcelArchivoCambioConcurrenteError,
    )

    error = ExcelArchivoCambioConcurrenteError(
        "DATA.xlsx",
        "sha-anterior",
        "sha-nuevo",
    )

    resultado = (
        clasificar_error_persistencia(
            error
        )
    )

    assert resultado.recuperable is True

    assert (
        resultado.codigo
        == "EXCEL_CAMBIO_CONCURRENTE"
    )
