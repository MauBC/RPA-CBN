from __future__ import annotations

from openpyxl import Workbook, load_workbook

from src.excel.pending_sync import (
    EstadoPersistenciaExcel,
    aplicar_actualizaciones_pendientes_a_snapshot,
    encolar_actualizacion_excel,
    obtener_actualizaciones_pendientes,
    persistir_o_encolar_resultado,
    ruta_journal_excel,
    sincronizar_actualizaciones_pendientes,
)


def _crear_excel(ruta):
    wb = Workbook()

    try:
        ws = wb.active
        ws.title = "Ordenes"

        ws.append(
            [
                "ID_ORDEN",
                "ESTADO_RPA",
                "RESUMEN",
            ]
        )

        ws.append(
            [
                "1001",
                0,
                "",
            ]
        )

        ws.append(
            [
                "1002",
                0,
                "",
            ]
        )

        wb.save(ruta)

    finally:
        wb.close()

    return ruta


def _leer_fila(ruta, id_orden):
    wb = load_workbook(
        ruta,
        data_only=True,
    )

    try:
        ws = wb["Ordenes"]

        for fila in range(
            2,
            ws.max_row + 1,
        ):
            if (
                str(
                    ws.cell(
                        fila,
                        1,
                    ).value
                )
                == str(id_orden)
            ):
                return {
                    "estado": ws.cell(
                        fila,
                        2,
                    ).value,
                    "resumen": ws.cell(
                        fila,
                        3,
                    ).value,
                }

    finally:
        wb.close()

    raise AssertionError(
        f"No se encontró ID {id_orden}"
    )


def test_journal_persiste_actualizacion(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_CBN_PENDING_DIR",
        str(
            tmp_path / "journal"
        ),
    )

    excel = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    journal = encolar_actualizacion_excel(
        excel,
        "1001",
        estado_rpa=1,
        resumen="OK CBN",
        detalle_error="Excel ocupado",
    )

    assert journal.exists()

    pendientes = (
        obtener_actualizaciones_pendientes(
            excel
        )
    )

    assert len(
        pendientes
    ) == 1

    assert (
        pendientes[0]["id_orden"]
        == "1001"
    )

    assert (
        pendientes[0]["estado_rpa"]
        == 1
    )

    assert (
        pendientes[0]["resumen"]
        == "OK CBN"
    )


def test_journal_deduplica_por_id_y_conserva_ultimo_valor(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_CBN_PENDING_DIR",
        str(
            tmp_path / "journal"
        ),
    )

    excel = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    encolar_actualizacion_excel(
        excel,
        "1001",
        estado_rpa=2,
        resumen=None,
    )

    encolar_actualizacion_excel(
        excel,
        "1001",
        estado_rpa=1,
        resumen="FINAL",
    )

    pendientes = (
        obtener_actualizaciones_pendientes(
            excel
        )
    )

    assert len(
        pendientes
    ) == 1

    assert (
        pendientes[0]["estado_rpa"]
        == 1
    )

    assert (
        pendientes[0]["resumen"]
        == "FINAL"
    )


def test_persistir_directamente_no_crea_journal(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_CBN_PENDING_DIR",
        str(
            tmp_path / "journal"
        ),
    )

    excel = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    resultado = (
        persistir_o_encolar_resultado(
            excel,
            "1001",
            estado_rpa=1,
            resumen="DIRECTO",
        )
    )

    assert (
        resultado.estado
        == EstadoPersistenciaExcel.APLICADO
    )

    assert not ruta_journal_excel(
        excel
    ).exists()

    fila = _leer_fila(
        excel,
        "1001",
    )

    assert fila["estado"] == 1
    assert fila["resumen"] == "DIRECTO"


def test_persistir_encola_si_excel_falla(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_CBN_PENDING_DIR",
        str(
            tmp_path / "journal"
        ),
    )

    excel = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    def fallar(*args, **kwargs):
        raise PermissionError(
            "archivo ocupado"
        )

    monkeypatch.setattr(
        "src.excel.state_manager.actualizar_resultado_orden",
        fallar,
    )

    resultado = (
        persistir_o_encolar_resultado(
            excel,
            "1001",
            estado_rpa=1,
            resumen="CBN OK",
        )
    )

    assert (
        resultado.estado
        == EstadoPersistenciaExcel.ENCOLADO
    )

    pendientes = (
        obtener_actualizaciones_pendientes(
            excel
        )
    )

    assert len(
        pendientes
    ) == 1

    assert (
        pendientes[0]["id_orden"]
        == "1001"
    )


def test_sincronizar_aplica_y_elimina_journal(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_CBN_PENDING_DIR",
        str(
            tmp_path / "journal"
        ),
    )

    excel = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    encolar_actualizacion_excel(
        excel,
        "1001",
        estado_rpa=1,
        resumen="SINCRONIZADO",
    )

    resultado = (
        sincronizar_actualizaciones_pendientes(
            excel
        )
    )

    assert resultado["total"] == 1
    assert resultado["aplicadas"] == 1
    assert resultado["restantes"] == 0

    fila = _leer_fila(
        excel,
        "1001",
    )

    assert fila["estado"] == 1

    assert (
        fila["resumen"]
        == "SINCRONIZADO"
    )

    assert not ruta_journal_excel(
        excel
    ).exists()


def test_overlay_snapshot_evitar_reprocesamiento(
    tmp_path,
    monkeypatch,
):
    monkeypatch.setenv(
        "RPA_CBN_PENDING_DIR",
        str(
            tmp_path / "journal"
        ),
    )

    original = _crear_excel(
        tmp_path / "DATA.xlsx"
    )

    snapshot = (
        tmp_path
        / "entrada_original.xlsx"
    )

    snapshot.write_bytes(
        original.read_bytes()
    )

    encolar_actualizacion_excel(
        original,
        "1001",
        estado_rpa=1,
        resumen="PORTAL YA OK",
    )

    resultado = (
        aplicar_actualizaciones_pendientes_a_snapshot(
            snapshot,
            original,
        )
    )

    assert resultado["aplicadas"] == 1

    fila_original = _leer_fila(
        original,
        "1001",
    )

    fila_snapshot = _leer_fila(
        snapshot,
        "1001",
    )

    # El DATA original continúa sin tocarse.
    assert fila_original["estado"] == 0

    # El snapshot sí refleja la realidad del portal.
    assert fila_snapshot["estado"] == 1

    assert (
        fila_snapshot["resumen"]
        == "PORTAL YA OK"
    )

    # El journal sigue existiendo porque todavía
    # no se sincronizó con el archivo original.
    assert ruta_journal_excel(
        original
    ).exists()
