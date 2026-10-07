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
    _accion_incidencia_gui,
    _color_estado_sincronizacion,
    _formatear_incidencias_gui,
    _texto_estado_sincronizacion_gui,
    _texto_tipo_incidencia_gui,
)


def _resultado_activo():
    return ResumenSincronizacionExcel(
        ruta=Path("DATA.xlsx"),
        estado=(
            EstadoSincronizacionExcel.EN_EJECUCION
        ),
        mensaje="1 orden procesándose ahora",
        pendientes=0,
        fallidas=0,
        puede_escribir=True,
        ocupado=False,
        inflight=1,
        incidencias=(
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO
                ),
                id_orden="1001",
                codigo="EJECUCION_ACTIVA",
                motivo=(
                    "Esta orden está siendo procesada "
                    "actualmente por esta instancia del RPA."
                ),
            ),
        ),
    )


def test_cp14e2_gui_muestra_procesando_ahora():
    texto = (
        _texto_estado_sincronizacion_gui(
            _resultado_activo()
        )
    )

    assert (
        "Procesando ahora: 1001"
        in texto
    )

    assert (
        "No requiere intervención"
        in texto
    )

    assert (
        "Ejecución incierta"
        not in texto
    )


def test_cp14e2_panel_nombre_activo():
    assert (
        _texto_tipo_incidencia_gui(
            TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO
        )
        == "Procesándose ahora"
    )


def test_cp14e2_panel_no_pide_reprocesar():
    texto = _accion_incidencia_gui(
        TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO
    )

    assert (
        "No se requiere intervención"
        in texto
    )

    assert (
        "Espere a que termine"
        in texto
    )


def test_cp14e2_formato_panel_activo():
    texto = (
        _formatear_incidencias_gui(
            _resultado_activo()
        )
    )

    assert (
        "Estado: Procesándose ahora"
        in texto
    )

    assert (
        "Código: EJECUCION_ACTIVA"
        in texto
    )


def test_cp14e2_en_ejecucion_no_usa_color_error():
    assert (
        _color_estado_sincronizacion(
            EstadoSincronizacionExcel.EN_EJECUCION
        )
        ==
        _color_estado_sincronizacion(
            EstadoSincronizacionExcel.PENDIENTE
        )
    )


def test_cp14e2_consulta_pasa_pid_de_esta_instancia():
    codigo = inspect.getsource(
        VentanaPrincipal._consultar_estado_sincronizacion_excel
    )

    assert "os.getpid()" in codigo

    assert (
        "if self._ejecutando"
        in codigo
    )

    assert (
        "pid_ejecucion_activa="
        in codigo
    )
