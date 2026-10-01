from pathlib import Path
import sys
from datetime import datetime

ROOT_DIR = Path(__file__).resolve().parents[1]

if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.browser.session import URL_SISTEMA, abrir_contexto
from src.browser.auth import esperar_login_manual
from src.excel.reader import leer_ordenes_excel
from src.excel.result_writer import guardar_resultados_ordenes
from src.excel.state_manager import (
    crear_excel_trabajo_pendientes,
    actualizar_estado_orden,
    obtener_fila_excel_por_id,
)
from src.excel.resumen_writer import actualizar_resumen_orden
from src.flows.asignar_comparativo import asignar_comparativo_desde_panel
from src.utils.debug import preparar_carpetas, guardar_evidencia_error
from src.utils.user_errors import construir_mensaje_usuario, construir_detalle_tecnico


RUTA_EXCEL = ROOT_DIR / "data" / "DATA.xlsx"


def construir_resultado_ok(indice: int, orden, info_asignacion) -> dict:
    return {
        "NRO": indice,
        "ID_ORDEN": orden.id_orden,
        "FILA_EXCEL": obtener_fila_excel_por_id(RUTA_EXCEL, orden.id_orden) or "",
        "ESTADO": "OK_ASIGNAR_COMPARATIVO",
        "N_COTIZACION": getattr(info_asignacion, "numero_cotizacion", ""),
        "ESTADO_CBN": getattr(info_asignacion, "estado_cbn", ""),
        "MENSAJE_USUARIO": getattr(info_asignacion, "mensaje", ""),
        "RESUMEN": getattr(info_asignacion, "resumen", ""),
        "PASO_ERROR": "",
        "DETALLE_TECNICO": "",
        "LOG": "",
        "SCREENSHOT": "",
        "HTML": "",
        "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def construir_resultado_error(indice: int, orden, error: Exception, rutas: dict) -> dict:
    return {
        "NRO": indice,
        "ID_ORDEN": getattr(orden, "id_orden", ""),
        "FILA_EXCEL": obtener_fila_excel_por_id(RUTA_EXCEL, getattr(orden, "id_orden", "")) or "",
        "ESTADO": "ERROR_ASIGNAR_COMPARATIVO",
        "N_COTIZACION": "",
        "ESTADO_CBN": "",
        "MENSAJE_USUARIO": construir_mensaje_usuario(error, orden),
        "PASO_ERROR": getattr(error, "paso", "asignar_comparativo"),
        "DETALLE_TECNICO": construir_detalle_tecnico(error),
        "LOG": str(rutas.get("log", "")),
        "SCREENSHOT": str(rutas.get("screenshot", "")),
        "HTML": str(rutas.get("html", "")),
        "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def main() -> None:
    preparar_carpetas()

    playwright = None
    contexto = None
    page = None
    resultados = []

    try:
        ruta_excel_trabajo, pendientes = crear_excel_trabajo_pendientes(RUTA_EXCEL)

        if not pendientes:
            print("No hay órdenes pendientes para asignar comparativo.")
            print("Este checkpoint procesa órdenes con ESTADO_RPA = 0.")
            return

        ordenes = leer_ordenes_excel(ruta_excel_trabajo)

        if not ordenes:
            raise RuntimeError("No hay órdenes válidas para asignar comparativo.")

        print("============================================================")
        print("CHECKPOINT: Asignar comparativo")
        print("Este script NO manda a cotizar y NO cierra.")
        print("Solo entra a COMPARATIVO y asigna proveedor.")
        print("Procesa órdenes con ESTADO_RPA = 0.")
        print("Si asigna bien, cambia ESTADO_RPA a 1.")
        print("Si falla, cambia ESTADO_RPA a 2.")
        print("============================================================")
        print(f"Órdenes pendientes de asignación: {len(ordenes)}")

        playwright, contexto, page = abrir_contexto()

        page.goto(
            URL_SISTEMA,
            wait_until="domcontentloaded",
            timeout=60_000,
        )

        esperar_login_manual(page)

        total = len(ordenes)

        for indice, orden in enumerate(ordenes, start=1):
            print("")
            print("------------------------------------------------------------")
            print(f"Asignando comparativo {indice}/{total}")
            print(f"ID_ORDEN: {orden.id_orden}")
            print(f"Fila Excel: {obtener_fila_excel_por_id(RUTA_EXCEL, orden.id_orden)}")
            print(f"Nombre de envío: {orden.texto}")
            print("------------------------------------------------------------")

            try:
                info_asignacion = asignar_comparativo_desde_panel(page, orden)

                print("Checkpoint asignar comparativo OK")
                print(f"N° de Cotización: {info_asignacion.numero_cotizacion}")
                print(f"URL COMPARATIVO: {info_asignacion.url_comparativo}")

                resultados.append(construir_resultado_ok(indice, orden, info_asignacion))

                actualizar_resumen_orden(
                    RUTA_EXCEL,
                    orden.id_orden,
                    getattr(info_asignacion, "resumen", ""),
                )
                print(f"RESUMEN actualizado para ID_ORDEN={orden.id_orden}")

                actualizar_estado_orden(RUTA_EXCEL, orden.id_orden, 1)
                print(f"ESTADO_RPA actualizado a 1 para ID_ORDEN={orden.id_orden}")

            except Exception as error:
                print("")
                print("Error en checkpoint asignar comparativo.")
                print(f"ID_ORDEN: {orden.id_orden}")
                print(f"Mensaje usuario: {construir_mensaje_usuario(error, orden)}")
                print(f"Detalle técnico: {error}")

                prefijo = f"asignar_comparativo_orden_{orden.id_orden}"
                rutas = guardar_evidencia_error(page, error, prefijo)

                if "log" in rutas:
                    print(f"Log guardado en: {rutas['log']}")
                if "screenshot" in rutas:
                    print(f"Screenshot guardado en: {rutas['screenshot']}")
                if "html" in rutas:
                    print(f"HTML guardado en: {rutas['html']}")

                resultados.append(construir_resultado_error(indice, orden, error, rutas))

                actualizar_resumen_orden(
                    RUTA_EXCEL,
                    orden.id_orden,
                    f"ERROR: {construir_detalle_tecnico(error)}",
                )
                print(f"RESUMEN actualizado con error para ID_ORDEN={orden.id_orden}")

                actualizar_estado_orden(RUTA_EXCEL, orden.id_orden, 2)
                print(f"ESTADO_RPA actualizado a 2 para ID_ORDEN={orden.id_orden}")

                print("Se continúa con la siguiente orden.")

        ruta_resultado = guardar_resultados_ordenes(RUTA_EXCEL, resultados)

        ok = sum(1 for item in resultados if item["ESTADO"] == "OK_ASIGNAR_COMPARATIVO")
        errores = len(resultados) - ok

        print("")
        print("============================================================")
        print("Checkpoint asignar comparativo terminado.")
        print(f"Órdenes asignadas OK: {ok}")
        print(f"Órdenes con error: {errores}")
        print(f"Resultado guardado en: {ruta_resultado}")
        print("============================================================")

    finally:
        if contexto is not None:
            try:
                contexto.close()
            except Exception:
                pass

        if playwright is not None:
            try:
                playwright.stop()
            except Exception:
                pass


if __name__ == "__main__":
    main()
