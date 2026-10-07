from __future__ import annotations

import inspect

from src.gui.main_window import (
    VentanaPrincipal,
    _texto_accion_reintento,
)


def test_gui_muestra_reintentar_solo_para_error():
    assert (
        _texto_accion_reintento(
            "Error"
        )
        == "Reintentar"
    )

    assert (
        _texto_accion_reintento(
            "Correcto"
        )
        == ""
    )

    assert (
        _texto_accion_reintento(
            "Pendiente"
        )
        == ""
    )

    assert (
        _texto_accion_reintento(
            "Procesando"
        )
        == ""
    )


def test_gui_retry_usa_servicio_seguro_y_hilo():
    codigo = inspect.getsource(
        VentanaPrincipal._preparar_reintento_orden
    )

    assert (
        "preparar_reintento_manual("
        in codigo
    )

    assert (
        "messagebox.askyesno("
        in codigo
    )

    assert (
        "Thread("
        in codigo
    )


def test_gui_treeview_usa_columna_accion():
    codigo = inspect.getsource(
        VentanaPrincipal._crear_interfaz
    )

    assert (
        '"accion"'
        in codigo
    )

    assert (
        '"Acción"'
        in codigo
    )

    assert (
        '"<ButtonRelease-1>"'
        in codigo
    )
