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



# ============================================================
# CP13A.2 - observabilidad del estado Excel en GUI
# ============================================================

from pathlib import Path

from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    IncidenciaSincronizacionExcel,
    ResumenSincronizacionExcel,
    TipoIncidenciaSincronizacion,
)
from src.gui.main_window import (
    _texto_estado_sincronizacion_gui,
)


def _resultado_cp13_gui(
    *,
    estado,
    mensaje,
    pendientes=0,
    fallidas=0,
    inflight=0,
    incidencias=(),
):
    return ResumenSincronizacionExcel(
        ruta=Path("DATA.xlsx"),
        estado=estado,
        mensaje=mensaje,
        pendientes=pendientes,
        fallidas=fallidas,
        inflight=inflight,
        puede_escribir=True,
        ocupado=False,
        incidencias=tuple(
            incidencias
        ),
    )


def test_cp13_gui_sincronizado_permanece_compacto():
    resultado = _resultado_cp13_gui(
        estado=EstadoSincronizacionExcel.SINCRONIZADO,
        mensaje="Excel listo · Sin cambios pendientes",
    )

    texto = (
        _texto_estado_sincronizacion_gui(
            resultado
        )
    )

    assert (
        texto
        == "Excel: Excel listo · Sin cambios pendientes"
    )

    assert "\n" not in texto


def test_cp13_gui_muestra_inflight_y_bloqueo_reproceso():
    resultado = _resultado_cp13_gui(
        estado=(
            EstadoSincronizacionExcel.REQUIERE_REVISION
        ),
        mensaje="1 orden en ejecución incierta",
        inflight=1,
        incidencias=(
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.INFLIGHT
                ),
                id_orden="1003",
                codigo="EJECUCION_INCIERTA",
                motivo="Ejecución anterior incierta.",
                reintento_automatico=False,
            ),
        ),
    )

    texto = (
        _texto_estado_sincronizacion_gui(
            resultado
        )
    )

    assert "1003" in texto

    assert (
        "No reprocesar automáticamente"
        in texto
    )


def test_cp13_gui_distingue_revision_y_pendiente():
    resultado = _resultado_cp13_gui(
        estado=(
            EstadoSincronizacionExcel.REQUIERE_REVISION
        ),
        mensaje=(
            "1 actualización requiere revisión · "
            "1 cambio pendiente"
        ),
        pendientes=1,
        fallidas=1,
        incidencias=(
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.FALLIDA
                ),
                id_orden="2001",
                codigo="DATOS_EXCEL_INVALIDOS",
                motivo="Cambio humano.",
                reintento_automatico=False,
            ),
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.PENDIENTE
                ),
                id_orden="2002",
                codigo="EXCEL_OCUPADO",
                motivo="Excel abierto.",
                reintento_automatico=True,
            ),
        ),
    )

    texto = (
        _texto_estado_sincronizacion_gui(
            resultado
        )
    )

    assert (
        "Revisión manual: 2001."
        in texto
    )

    assert (
        "Pendientes de sincronizar: 2002."
        in texto
    )


def test_cp13_gui_limita_ids_para_no_saturar_interfaz():
    incidencias = tuple(
        IncidenciaSincronizacionExcel(
            tipo=(
                TipoIncidenciaSincronizacion.PENDIENTE
            ),
            id_orden=str(
                3000 + indice
            ),
            reintento_automatico=True,
        )
        for indice in range(7)
    )

    resultado = _resultado_cp13_gui(
        estado=(
            EstadoSincronizacionExcel.PENDIENTE
        ),
        mensaje="7 cambios pendientes",
        pendientes=7,
        incidencias=incidencias,
    )

    texto = (
        _texto_estado_sincronizacion_gui(
            resultado
        )
    )

    assert "3000" in texto
    assert "3001" in texto
    assert "3002" in texto
    assert "3003" in texto

    assert "3004" not in texto

    assert (
        "(+3 más)"
        in texto
    )



def test_cp13_gui_revision_advierte_no_reprocesar():
    resultado = _resultado_cp13_gui(
        estado=(
            EstadoSincronizacionExcel.REQUIERE_REVISION
        ),
        mensaje="1 actualización requiere revisión",
        fallidas=1,
        incidencias=(
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.FALLIDA
                ),
                id_orden="9001",
                codigo="DATOS_EXCEL_INVALIDOS",
                motivo="Cambio humano detectado.",
                reintento_automatico=False,
            ),
        ),
    )

    texto = (
        _texto_estado_sincronizacion_gui(
            resultado
        )
    )

    assert (
        "Revisión manual: 9001."
        in texto
    )

    assert (
        "No reprocesar automáticamente"
        in texto
    )
