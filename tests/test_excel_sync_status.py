from __future__ import annotations

from src.excel.file_access import (
    DiagnosticoAccesoExcel,
    EstadoAccesoExcel,
)
from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    TipoIncidenciaSincronizacion,
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
    inflight=0,
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
                {
                    "id_orden": f"fallida-{i}"
                }
                for i in range(fallidas)
            ],
            "inflight": [
                {
                    "id_orden": f"inflight-{i}"
                }
                for i in range(inflight)
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
    assert resultado.inflight == 0
    assert resultado.incidencias == ()


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



def test_cp13_inflight_se_conserva_como_dato_estructurado(
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
        inflight=1,
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

    assert resultado.inflight == 1
    assert resultado.pendientes == 0
    assert resultado.fallidas == 0
    assert len(resultado.incidencias) == 1

    incidencia = resultado.incidencias[0]

    assert (
        incidencia.tipo
        == TipoIncidenciaSincronizacion.INFLIGHT
    )

    assert incidencia.id_orden == "inflight-0"
    assert incidencia.codigo == "EJECUCION_INCIERTA"
    assert incidencia.reintento_automatico is False

    assert (
        "No debe reprocesarse automáticamente"
        in incidencia.motivo
    )


def test_cp13_incidencias_conservan_contexto_operativo(
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

    monkeypatch.setattr(
        "src.excel.sync_status.obtener_estado_journal",
        lambda _: {
            "pendientes": [
                {
                    "id_orden": "1001",
                    "created_at": "2026-10-02T10:00:00",
                    "updated_at": "2026-10-02T10:05:00",
                    "intentos": 3,
                    "ultimo_codigo_error": "EXCEL_OCUPADO",
                    "ultimo_error": "Archivo abierto.",
                }
            ],
            "fallidas": [
                {
                    "id_orden": "1002",
                    "created_at": "2026-10-02T10:10:00",
                    "updated_at": "2026-10-02T10:12:00",
                    "sync_failed_at": "2026-10-02T10:13:00",
                    "intentos": 4,
                    "sync_error_code": "DATOS_EXCEL_INVALIDOS",
                    "sync_error_detail": "Cambio humano detectado.",
                }
            ],
            "inflight": [
                {
                    "id_orden": "1003",
                    "started_at": "2026-10-02T10:20:00",
                    "updated_at": "2026-10-02T10:21:00",
                }
            ],
        },
    )

    resultado = (
        obtener_estado_sincronizacion_excel(
            ruta
        )
    )

    assert resultado.pendientes == 1
    assert resultado.fallidas == 1
    assert resultado.inflight == 1
    assert len(resultado.incidencias) == 3

    por_tipo = {
        incidencia.tipo: incidencia
        for incidencia in resultado.incidencias
    }

    pendiente = por_tipo[
        TipoIncidenciaSincronizacion.PENDIENTE
    ]

    assert pendiente.id_orden == "1001"
    assert pendiente.fecha == "2026-10-02T10:05:00"
    assert pendiente.intentos == 3
    assert pendiente.codigo == "EXCEL_OCUPADO"
    assert pendiente.motivo == "Archivo abierto."
    assert pendiente.reintento_automatico is True

    fallida = por_tipo[
        TipoIncidenciaSincronizacion.FALLIDA
    ]

    assert fallida.id_orden == "1002"
    assert fallida.fecha == "2026-10-02T10:13:00"
    assert fallida.intentos == 4
    assert fallida.codigo == "DATOS_EXCEL_INVALIDOS"
    assert fallida.motivo == "Cambio humano detectado."
    assert fallida.reintento_automatico is False

    incierta = por_tipo[
        TipoIncidenciaSincronizacion.INFLIGHT
    ]

    assert incierta.id_orden == "1003"
    assert incierta.fecha == "2026-10-02T10:20:00"
    assert incierta.codigo == "EJECUCION_INCIERTA"
    assert incierta.reintento_automatico is False

    assert " ? " not in resultado.mensaje
    assert " · " in resultado.mensaje


def test_cp13_journal_sin_inflight_mantiene_compatibilidad(
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

    # Simula journal anterior a la incorporación de inflight.
    monkeypatch.setattr(
        "src.excel.sync_status.obtener_estado_journal",
        lambda _: {
            "pendientes": [
                {
                    "id_orden": "1001",
                }
            ],
            "fallidas": [],
        },
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

    assert resultado.pendientes == 1
    assert resultado.fallidas == 0
    assert resultado.inflight == 0
    assert len(resultado.incidencias) == 1

    assert (
        resultado.incidencias[0].tipo
        == TipoIncidenciaSincronizacion.PENDIENTE
    )
