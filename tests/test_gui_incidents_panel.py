from __future__ import annotations

from pathlib import Path
import inspect

from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    IncidenciaSincronizacionExcel,
    ResumenSincronizacionExcel,
    TipoIncidenciaSincronizacion,
)
from src.gui.main_window import (
    VentanaPrincipal,
    _formatear_incidencias_gui,
    _puede_ver_incidencias,
)


def _resultado(
    incidencias=(),
):
    return ResumenSincronizacionExcel(
        ruta=Path("DATA.xlsx"),
        estado=(
            EstadoSincronizacionExcel.REQUIERE_REVISION
            if incidencias
            else EstadoSincronizacionExcel.SINCRONIZADO
        ),
        mensaje="prueba",
        pendientes=0,
        fallidas=0,
        puede_escribir=True,
        ocupado=False,
        incidencias=tuple(
            incidencias
        ),
    )


def test_cp14d_formatea_pendiente_con_accion_segura():
    resultado = _resultado(
        (
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.PENDIENTE
                ),
                id_orden="72",
                codigo="EXCEL_OCUPADO",
                motivo="OneDrive bloqueó el archivo.",
                intentos=3,
                reintento_automatico=True,
            ),
        )
    )

    texto = _formatear_incidencias_gui(
        resultado
    )

    assert "Orden: 72" in texto

    assert (
        "Pendiente de sincronización"
        in texto
    )

    assert (
        "Sincronizar Excel"
        in texto
    )

    assert (
        "No vuelva a ejecutar esta orden en CBN"
        in texto
    )


def test_cp14d_formatea_failed_update_como_revision():
    resultado = _resultado(
        (
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.FALLIDA
                ),
                id_orden="80",
                codigo="DATOS_EXCEL_INVALIDOS",
                motivo="Cambio humano detectado.",
                reintento_automatico=False,
            ),
        )
    )

    texto = _formatear_incidencias_gui(
        resultado
    )

    assert "Requiere revisión" in texto

    assert (
        "DATOS_EXCEL_INVALIDOS"
        in texto
    )

    assert (
        "No reprocesar automáticamente"
        in texto
    )


def test_cp14d_formatea_inflight_como_incierto():
    resultado = _resultado(
        (
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.INFLIGHT
                ),
                id_orden="90",
                codigo="EJECUCION_INCIERTA",
                motivo="Resultado no confirmado.",
                reintento_automatico=False,
            ),
        )
    )

    texto = _formatear_incidencias_gui(
        resultado
    )

    assert "Ejecución incierta" in texto

    assert (
        "Confirme primero"
        in texto
    )


def test_cp14d_boton_solo_si_hay_incidencias():
    vacio = _resultado()

    assert (
        _puede_ver_incidencias(
            vacio
        )
        is False
    )

    con_incidencia = _resultado(
        (
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.PENDIENTE
                ),
                id_orden="72",
            ),
        )
    )

    assert (
        _puede_ver_incidencias(
            con_incidencia
        )
        is True
    )


def test_cp14d_panel_no_se_abre_durante_ejecucion():
    resultado = _resultado(
        (
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.PENDIENTE
                ),
                id_orden="72",
            ),
        )
    )

    assert (
        _puede_ver_incidencias(
            resultado,
            ejecutando=True,
        )
        is False
    )


def test_cp14d_gui_tiene_panel_solo_lectura():
    codigo = inspect.getsource(
        VentanaPrincipal._mostrar_panel_incidencias
    )

    assert "CTkToplevel" in codigo
    assert "CTkTextbox" in codigo

    assert (
        "_formatear_incidencias_gui("
        in codigo
    )

    assert "ejecutar_rpa(" not in codigo
    assert "sincronizar_actualizaciones_pendientes(" not in codigo
