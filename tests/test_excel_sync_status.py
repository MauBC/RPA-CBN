from __future__ import annotations

from src.excel.file_access import (
    DiagnosticoAccesoExcel,
    EstadoAccesoExcel,
)
from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    obtener_estado_sincronizacion_excel,
)


def _configurar(
    monkeypatch,
    tmp_path,
    *,
    estado_acceso,
    puede_escribir,
    pendientes=0,
    fallidas=0,
    winerror=None,
):
    ruta = tmp_path / "DATA.xlsx"

    if (
        estado_acceso
        != EstadoAccesoExcel.NO_EXISTE
    ):
        ruta.write_bytes(
            b"prueba"
        )

    diagnostico = (
        DiagnosticoAccesoExcel(
            ruta=ruta,
            estado=estado_acceso,
            puede_escribir=puede_escribir,
            lock_office_detectado=False,
            detalle="diagnóstico test",
            winerror=winerror,
        )
    )

    monkeypatch.setattr(
        "src.excel.sync_status.diagnosticar_acceso_excel",
        lambda _: diagnostico,
    )

    monkeypatch.setattr(
        "src.excel.sync_status.obtener_estado_journal",
        lambda _: {
            "pendientes": [
                {"id_orden": str(i)}
                for i in range(pendientes)
            ],
            "fallidas": [
                {"id_orden": str(i)}
                for i in range(fallidas)
            ],
        },
    )

    return ruta


def test_estado_sincronizado(
    tmp_path,
    monkeypatch,
):
    ruta = _configurar(
        monkeypatch,
        tmp_path,
        estado_acceso=(
            EstadoAccesoExcel.DISPONIBLE
        ),
        puede_escribir=True,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.SINCRONIZADO
    )

    assert resultado.sincronizado is True
    assert resultado.pendientes == 0
    assert resultado.fallidas == 0


def test_excel_ocupado_sin_pendientes(
    tmp_path,
    monkeypatch,
):
    ruta = _configurar(
        monkeypatch,
        tmp_path,
        estado_acceso=(
            EstadoAccesoExcel.OCUPADO
        ),
        puede_escribir=False,
        winerror=32,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.EXCEL_OCUPADO
    )

    assert resultado.ocupado is True
    assert resultado.winerror == 32

    assert (
        "Sin cambios pendientes"
        in resultado.mensaje
    )


def test_excel_ocupado_con_pendientes(
    tmp_path,
    monkeypatch,
):
    ruta = _configurar(
        monkeypatch,
        tmp_path,
        estado_acceso=(
            EstadoAccesoExcel.OCUPADO
        ),
        puede_escribir=False,
        pendientes=3,
        winerror=32,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.EXCEL_OCUPADO
    )

    assert resultado.pendientes == 3

    assert (
        "3 cambios pendientes"
        in resultado.mensaje
    )


def test_pendientes_con_excel_disponible(
    tmp_path,
    monkeypatch,
):
    ruta = _configurar(
        monkeypatch,
        tmp_path,
        estado_acceso=(
            EstadoAccesoExcel.DISPONIBLE
        ),
        puede_escribir=True,
        pendientes=2,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.PENDIENTE
    )

    assert resultado.pendientes == 2
    assert resultado.puede_escribir is True


def test_failed_updates_tienen_prioridad_visual(
    tmp_path,
    monkeypatch,
):
    ruta = _configurar(
        monkeypatch,
        tmp_path,
        estado_acceso=(
            EstadoAccesoExcel.OCUPADO
        ),
        puede_escribir=False,
        pendientes=2,
        fallidas=1,
        winerror=32,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.REQUIERE_REVISION
    )

    assert resultado.requiere_atencion is True
    assert resultado.fallidas == 1
    assert resultado.pendientes == 2

    assert (
        "1 actualización requiere revisión"
        in resultado.mensaje
    )


def test_error_leyendo_journal_no_se_oculta(
    tmp_path,
    monkeypatch,
):
    ruta = _configurar(
        monkeypatch,
        tmp_path,
        estado_acceso=(
            EstadoAccesoExcel.DISPONIBLE
        ),
        puede_escribir=True,
    )

    def fallar(_):
        raise RuntimeError(
            "journal corrupto"
        )

    monkeypatch.setattr(
        "src.excel.sync_status.obtener_estado_journal",
        fallar,
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert (
        resultado.estado
        == EstadoSincronizacionExcel.ERROR_ACCESO
    )

    assert resultado.requiere_atencion is True

    assert (
        "journal corrupto"
        in resultado.detalle
    )
