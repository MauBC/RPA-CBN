import inspect

from src.core import runner


def test_ejecutar_rpa_integra_snapshots():
    codigo = inspect.getsource(
        runner.ejecutar_rpa
    )

    assert "crear_snapshot_estable(" in codigo

    assert (
        "crear_excel_trabajo_pendientes(entrada_original)"
        in codigo
    )

    assert "congelar_templates_ordenes(" in codigo

    assert "validar_ordenes_cargadas(" in codigo

    assert "guardar_manifest_snapshots(" in codigo


def test_validacion_manual_no_usa_snapshot_de_ejecucion():
    codigo = inspect.getsource(
        runner.validar_excel_sin_ejecutar
    )

    assert "solo_lectura=True" in codigo

    assert "crear_snapshot_estable(" not in codigo

    assert "entrada_original" not in codigo
