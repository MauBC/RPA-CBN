import inspect

from src.core import runner


def test_procesar_orden_no_persiste_excel():
    codigo = inspect.getsource(
        runner.procesar_orden
    )

    assert "actualizar_resumen_orden(" not in codigo
    assert "actualizar_estado_orden(" not in codigo
    assert "actualizar_resultado_orden(" not in codigo


def test_runner_exito_usa_actualizacion_combinada():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    assert "actualizar_resultado_orden(" in codigo
    assert "estado_rpa=1" in codigo
    assert "resumen=resumen_orden" in codigo
