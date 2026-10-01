from openpyxl import Workbook

from src.core import runner


def test_validar_excel_usa_modo_solo_lectura(
    tmp_path,
    monkeypatch,
):
    ruta = tmp_path / "DATA.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Ordenes"
    ws.append(["ID_ORDEN"])
    ws.append(["1001"])
    wb.save(ruta)
    wb.close()

    llamada = {}

    def fake_crear_excel_trabajo_pendientes(
        ruta_excel,
        *,
        solo_lectura=False,
    ):
        llamada["ruta"] = ruta_excel
        llamada["solo_lectura"] = solo_lectura

        return (
            tmp_path / "_temporal_inexistente.xlsx",
            [],
        )

    monkeypatch.setattr(
        runner,
        "crear_excel_trabajo_pendientes",
        fake_crear_excel_trabajo_pendientes,
    )

    resultado = runner.validar_excel_sin_ejecutar(ruta)

    assert resultado["ok"] is True
    assert resultado["pending_count"] == 0

    assert llamada["solo_lectura"] is True
