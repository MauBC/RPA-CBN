from __future__ import annotations

import inspect

from src.gui import main_window


def test_gui_tiene_indicador_sync_independiente():
    codigo = inspect.getsource(
        main_window.VentanaPrincipal._crear_interfaz
    )

    assert (
        "etiqueta_sync_excel"
        in codigo
    )

    assert (
        "etiqueta_estado"
        in codigo
    )


def test_consulta_sync_se_ejecuta_en_thread():
    codigo = inspect.getsource(
        main_window.VentanaPrincipal._consultar_estado_sincronizacion_excel
    )

    assert (
        "obtener_estado_sincronizacion_excel("
        in codigo
    )

    assert "Thread(" in codigo

    assert (
        '"excel_sync_status"'
        in codigo
    )


def test_gui_procesa_eventos_sync_por_cola():
    codigo = inspect.getsource(
        main_window.VentanaPrincipal._manejar_evento
    )

    assert (
        '"excel_sync_status"'
        in codigo
    )

    assert (
        '"excel_sync_status_error"'
        in codigo
    )

    assert (
        '"excel_sync_status_finished"'
        in codigo
    )


def test_ciclo_sync_es_periodico():
    codigo = inspect.getsource(
        main_window.VentanaPrincipal._ciclo_estado_sincronizacion_excel
    )

    assert (
        "_consultar_estado_sincronizacion_excel()"
        in codigo
    )

    assert "self.after(" in codigo

    assert (
        "self._intervalo_sync_ms"
        in codigo
    )
