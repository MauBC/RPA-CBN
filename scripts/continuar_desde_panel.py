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
from src.flows.mandar_cotizar import mandar_a_cotizar_orden
from src.utils.debug import preparar_carpetas, guardar_evidencia_error
from src.utils.user_errors import construir_mensaje_usuario, construir_detalle_tecnico


RUTA_EXCEL = ROOT_DIR / "data" / "DATA.xlsx"


def construir_resultado_ok(indice: int, orden, info_mandar) -> dict:
    fila_excel = obtener_fila_excel_por_id(RUTA_EXCEL, orden.id_orden) or ""
    estado = "OK_MANDAR_COTIZAR"

    if getattr(info_mandar, "omitido", False):
        estado = "OMITIDO_BIEN"

    return {
        "NRO": indice,
        "ID_ORDEN": orden.id_orden,
        "FILA_EXCEL": fila_excel,
        "ESTADO": estado,
        "N_COTIZACION": getattr(info_mandar, "numero_cotizacion", ""),
        "ESTADO_CBN": getattr(info_mandar, "estado_cbn", ""),
        "MENSAJE_USUARIO": getattr(info_mandar, "mensaje", ""),
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
        "ESTADO": "ERROR_MANDAR_COTIZAR",
        "N_COTIZACION": "",
        "ESTADO_CBN": "",
        "MENSAJE_USUARIO": construir_mensaje_usuario(error, orden),
        "PASO_ERROR": getattr(error, "paso", "mandar_a_cotizar"),
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
            print("No hay órdenes pendientes con ESTADO_RPA = 0.")
            print("No se ejecutará mandar a cotizar.")
            return

        ordenes = leer_ordenes_excel(ruta_excel_trabajo)

        if not ordenes:
            raise RuntimeError("No hay órdenes pendientes válidas en DATA.xlsx.")

        print("============================================================")
        print("CHECKPOINT: Mandar a cotizar")
        print("Este script NO crea cotizaciones nuevas.")
        print("Busca cotizaciones ya creadas y ejecuta la fase mandar a cotizar.")
        print("============================================================")
        print(f"Órdenes pendientes a procesar: {len(ordenes)}")

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
            print(f"Procesando checkpoint {indice}/{total}")
            print(f"ID_ORDEN: {orden.id_orden}")
            print(f"Fila Excel: {obtener_fila_excel_por_id(RUTA_EXCEL, orden.id_orden)}")
            print(f"Nombre de envío: {orden.texto}")
            print("------------------------------------------------------------")

            try:
                info_mandar = mandar_a_cotizar_orden(page, orden)

                print("Checkpoint mandar a cotizar OK")
                print(f"N° de Cotización: {info_mandar.numero_cotizacion}")
                print(f"Estado CBN: {info_mandar.estado_cbn}")
                print(f"Posiciones cotizadas: {info_mandar.posiciones_cotizadas}")

                resultados.append(construir_resultado_ok(indice, orden, info_mandar))
                actualizar_estado_orden(RUTA_EXCEL, orden.id_orden, 1)
                print(f"ESTADO_RPA actualizado a 1 para ID_ORDEN={orden.id_orden}")

            except Exception as error:
                print("")
                print("Error en checkpoint mandar a cotizar.")
                print(f"ID_ORDEN: {orden.id_orden}")
                print(f"Mensaje usuario: {construir_mensaje_usuario(error, orden)}")
                print(f"Detalle técnico: {error}")

                prefijo = f"mandar_cotizar_orden_{orden.id_orden}"
                rutas = guardar_evidencia_error(page, error, prefijo)

                if "log" in rutas:
                    print(f"Log guardado en: {rutas['log']}")
                if "screenshot" in rutas:
                    print(f"Screenshot guardado en: {rutas['screenshot']}")
                if "html" in rutas:
                    print(f"HTML guardado en: {rutas['html']}")

                resultados.append(construir_resultado_error(indice, orden, error, rutas))
                actualizar_estado_orden(RUTA_EXCEL, orden.id_orden, 2)
                print(f"ESTADO_RPA actualizado a 2 para ID_ORDEN={orden.id_orden}")

                print("Se continúa con la siguiente orden.")

        ruta_resultado = guardar_resultados_ordenes(RUTA_EXCEL, resultados)

        ok = sum(1 for item in resultados if item["ESTADO"] == "OK_MANDAR_COTIZAR")
        omitidos = sum(1 for item in resultados if item["ESTADO"] == "OMITIDO_BIEN")
        errores = len(resultados) - ok - omitidos

        print("")
        print("============================================================")
        print("Checkpoint mandar a cotizar terminado.")
        print(f"Órdenes OK: {ok}")
        print(f"Órdenes omitidas por BIEN: {omitidos}")
        print(f"Órdenes con error: {errores}")
        print(f"Resultado guardado en: {ruta_resultado}")
        print("============================================================")

    except Exception as error:
        print("")
        print("Error general en checkpoint mandar a cotizar.")
        print(error)

        rutas = guardar_evidencia_error(page, error, "mandar_cotizar_error_general")

        resultado = [
            {
                "NRO": 1,
                "ID_ORDEN": "",
                "ESTADO": "ERROR_MANDAR_COTIZAR",
                "N_COTIZACION": "",
                "ESTADO_CBN": "",
                "MENSAJE_USUARIO": construir_mensaje_usuario(error, None),
                "PASO_ERROR": "mandar_a_cotizar",
                "DETALLE_TECNICO": str(error),
                "LOG": str(rutas.get("log", "")),
                "SCREENSHOT": str(rutas.get("screenshot", "")),
                "HTML": str(rutas.get("html", "")),
                "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ]

        ruta_resultado = guardar_resultados_ordenes(RUTA_EXCEL, resultado)
        print(f"Resultado guardado en: {ruta_resultado}")

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
