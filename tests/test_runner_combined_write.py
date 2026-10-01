import inspect

from src.core import runner


def test_procesar_orden_no_persiste_excel():
    codigo = inspect.getsource(
        runner.procesar_orden
    )

    assert "actualizar_resumen_orden(" not in codigo
    assert "actualizar_estado_orden(" not in codigo
    assert "actualizar_resultado_orden(" not in codigo
    assert "_persistir_resultado_excel(" not in codigo


def test_runner_exito_usa_persistencia_desacoplada():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    # Desde CP7, ejecutar_rpa no escribe Excel directamente.
    assert "actualizar_resultado_orden(" not in codigo
    assert "actualizar_estado_orden(" not in codigo

    # El resultado exitoso pasa por la capa que puede
    # escribir inmediatamente o encolar.
    assert "_persistir_resultado_excel(" in codigo
    assert "estado_rpa=1" in codigo
    assert "resumen=resumen_orden" in codigo
