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
