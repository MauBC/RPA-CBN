from __future__ import annotations

from types import SimpleNamespace

from src.excel import sync_status
from src.excel.file_access import (
    EstadoAccesoExcel,
)
from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    TipoIncidenciaSincronizacion,
    obtener_estado_sincronizacion_excel,
)


def _preparar(
    tmp_path,
    monkeypatch,
    *,
    pid_journal,
):
    excel = tmp_path / "DATA.xlsx"

    excel.write_bytes(
        b"dummy"
    )

    diagnostico = SimpleNamespace(
        ruta=excel,
        estado=(
            EstadoAccesoExcel.DISPONIBLE
        ),
        puede_escribir=True,
        ocupado=False,
        winerror=None,
        detalle="OK",
    )

    monkeypatch.setattr(
        sync_status,
        "diagnosticar_acceso_excel",
        lambda _: diagnostico,
    )

    monkeypatch.setattr(
        sync_status,
        "obtener_estado_journal",
        lambda _: {
            "pendientes": [],
            "fallidas": [],
            "inflight": [
                {
                    "id_orden": "1001",
                    "pid": pid_journal,
                    "started_at": (
                        "2026-10-06T20:00:00+00:00"
                    ),
                }
            ],
        },
    )

    return excel


def test_inflight_mismo_pid_confirmado_es_activo(
    tmp_path,
    monkeypatch,
):
    excel = _preparar(
        tmp_path,
        monkeypatch,
        pid_journal=12345,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            excel,
            pid_ejecucion_activa=12345,
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.EN_EJECUCION
    )

    assert (
        resultado.requiere_atencion
        is False
    )

    assert resultado.inflight == 1

    incidencia = resultado.incidencias[0]

    assert (
        incidencia.tipo
        == TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO
    )

    assert (
        incidencia.codigo
        == "EJECUCION_ACTIVA"
    )


def test_inflight_sin_pid_confirmado_sigue_incierto(
    tmp_path,
    monkeypatch,
):
    excel = _preparar(
        tmp_path,
        monkeypatch,
        pid_journal=12345,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            excel
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.REQUIERE_REVISION
    )

    assert (
        resultado.requiere_atencion
        is True
    )

    assert (
        resultado.incidencias[0].tipo
        == TipoIncidenciaSincronizacion.INFLIGHT
    )


def test_inflight_pid_distinto_sigue_incierto(
    tmp_path,
    monkeypatch,
):
    excel = _preparar(
        tmp_path,
        monkeypatch,
        pid_journal=111,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            excel,
            pid_ejecucion_activa=222,
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.REQUIERE_REVISION
    )

    assert (
        resultado.incidencias[0].codigo
        == "EJECUCION_INCIERTA"
    )


def test_inflight_activo_no_se_confunde_con_revision(
    tmp_path,
    monkeypatch,
):
    excel = _preparar(
        tmp_path,
        monkeypatch,
        pid_journal=777,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            excel,
            pid_ejecucion_activa=777,
        )
    )

    assert resultado.fallidas == 0

    assert (
        "procesándose ahora"
        in resultado.mensaje
    )

    assert (
        "incierta"
        not in resultado.mensaje
    )
