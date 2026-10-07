from __future__ import annotations

from pathlib import Path
import inspect

from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    ResumenSincronizacionExcel,
)
from src.gui.main_window import (
    VentanaPrincipal,
    _puede_sincronizar_manualmente,
)


def _estado(
    *,
    pendientes: int,
    puede_escribir: bool,
):
    return ResumenSincronizacionExcel(
        ruta=Path("DATA.xlsx"),
        estado=(
            EstadoSincronizacionExcel.PENDIENTE
            if pendientes
            else EstadoSincronizacionExcel.SINCRONIZADO
        ),
        mensaje="prueba",
        pendientes=pendientes,
        fallidas=0,
        puede_escribir=puede_escribir,
        ocupado=not puede_escribir,
    )


def test_cp14b_boton_solo_si_hay_pendientes():
    assert (
        _puede_sincronizar_manualmente(
            _estado(
                pendientes=1,
                puede_escribir=True,
            )
        )
        is True
    )

    assert (
        _puede_sincronizar_manualmente(
            _estado(
                pendientes=0,
                puede_escribir=True,
            )
        )
        is False
    )


def test_cp14b_excel_no_escribible_bloquea_boton():
    assert (
        _puede_sincronizar_manualmente(
            _estado(
                pendientes=1,
                puede_escribir=False,
            )
        )
        is False
    )


def test_cp14b_rpa_o_sync_bloquean_boton():
    estado = _estado(
        pendientes=1,
        puede_escribir=True,
    )

    assert (
        _puede_sincronizar_manualmente(
            estado,
            ejecutando=True,
        )
        is False
    )

    assert (
        _puede_sincronizar_manualmente(
            estado,
            sincronizando=True,
        )
        is False
    )


def test_cp14b_gui_tiene_boton_separado():
    codigo = inspect.getsource(
        VentanaPrincipal._crear_interfaz
    )

    assert (
        'text="Sincronizar Excel"'
        in codigo
    )

    assert (
        "command=self._sincronizar_excel_manual"
        in codigo
    )


def test_cp14b_sincronizar_no_ejecuta_cbn():
    codigo = inspect.getsource(
        VentanaPrincipal._sincronizar_excel_manual
    )

    assert (
        "sincronizar_actualizaciones_pendientes("
        in codigo
    )

    assert "Thread(" in codigo

    assert "ejecutar_rpa(" not in codigo


def test_cp14b_eventos_sync():
    codigo = inspect.getsource(
        VentanaPrincipal._manejar_evento
    )

    assert (
        '"manual_sync_finished"'
        in codigo
    )

    assert (
        '"manual_sync_error"'
        in codigo
    )
