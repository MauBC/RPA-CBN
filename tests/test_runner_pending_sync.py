from __future__ import annotations

import inspect

from src.core import runner


def test_runner_no_escribe_estado_excel_directamente():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    assert (
        "actualizar_resultado_orden("
        not in codigo
    )

    assert (
        "actualizar_estado_orden("
        not in codigo
    )

    assert (
        codigo.count(
            "_persistir_resultado_excel("
        )
        == 3
    )


def test_runner_sincroniza_antes_del_snapshot_y_overlay_antes_de_pendientes():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    pos_sync = codigo.index(
        "sincronizar_actualizaciones_pendientes("
    )

    pos_snapshot = codigo.index(
        "crear_snapshot_estable("
    )

    pos_overlay = codigo.index(
        "aplicar_actualizaciones_pendientes_a_snapshot("
    )

    pos_trabajo = codigo.index(
        "crear_excel_trabajo_pendientes("
    )

    assert (
        pos_sync
        < pos_snapshot
        < pos_overlay
        < pos_trabajo
    )


def test_resultado_ok_se_registra_antes_de_persistir_excel():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    pos_resultado = codigo.index(
        "resultados.append(resultado_ok)"
    )

    pos_persistencia = codigo.index(
        "_persistir_resultado_excel(",
        pos_resultado,
    )

    pos_evento_ok = codigo.index(
        '"order_ok"',
        pos_persistencia,
    )

    assert (
        pos_resultado
        < pos_persistencia
        < pos_evento_ok
    )


def test_helper_de_persistencia_usa_journal():
    codigo = inspect.getsource(
        runner._persistir_resultado_excel
    )

    assert (
        "persistir_o_encolar_resultado("
        in codigo
    )

    assert (
        "resultado.encolado"
        in codigo
    )

    assert (
        "raise "
        not in codigo
    )


def test_runner_bloquea_overlay_fallido_antes_de_abrir_navegador():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    pos_overlay = codigo.index(
        "aplicar_actualizaciones_pendientes_a_snapshot("
    )

    pos_bloqueo = codigo.index(
        'if overlay_pendientes["fallidas"]'
    )

    pos_browser = codigo.index(
        "abrir_contexto("
    )

    assert (
        pos_overlay
        < pos_bloqueo
        < pos_browser
    )

    assert (
        "_mensaje_bloqueo_sincronizacion("
        in codigo
    )


def test_mensaje_fatal_preserva_bloqueo_de_sincronizacion():
    mensaje = (
        "Se requiere revisión manual para evitar "
        "reprocesar órdenes ya ejecutadas en CBN."
    )

    error = (
        runner.SincronizacionExcelRequiereRevisionError(
            mensaje
        )
    )

    assert (
        runner._mensaje_error_fatal(
            error
        )
        == mensaje
    )



def test_cp13_mensaje_revision_incluye_ids_afectados():
    overlay = {
        "fallidas": 2,
        "detalles": [
            {
                "id_orden": "1001",
                "estado": (
                    "FAILED_UPDATE_REQUIERE_REVISION"
                ),
                "detalle": "Cambio humano.",
            },
            {
                "id_orden": "1002",
                "estado": (
                    "INFLIGHT_OVERLAY_ERROR"
                ),
                "detalle": "No se pudo aplicar overlay.",
            },
        ],
    }

    mensaje = (
        runner._mensaje_bloqueo_sincronizacion(
            overlay
        )
    )

    assert "revisión manual" in mensaje
    assert "1001" in mensaje
    assert "1002" in mensaje

    assert (
        "No reprocesar automáticamente"
        in mensaje
    )


def test_cp13_mensaje_revision_limita_ids():
    overlay = {
        "detalles": [
            {
                "id_orden": str(
                    1000 + indice
                ),
                "estado": (
                    "FAILED_UPDATE_REQUIERE_REVISION"
                ),
            }
            for indice in range(7)
        ],
    }

    mensaje = (
        runner._mensaje_bloqueo_sincronizacion(
            overlay
        )
    )

    assert "1000" in mensaje
    assert "1004" in mensaje
    assert "1005" not in mensaje
    assert "(+2 más)" in mensaje
