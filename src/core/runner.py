from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Any, Callable
import json
import shutil

from src.browser.session import URL_SISTEMA, abrir_contexto
from src.browser.auth import (
    NavegadorCerradoPorUsuarioError,
    esperar_login_manual,
)
from src.excel.reader import (
    leer_ordenes_excel,
    validar_ordenes_cargadas,
)
from src.excel.snapshot_manager import (
    congelar_templates_ordenes,
    crear_snapshot_estable,
    guardar_manifest_snapshots,
)
from src.excel.validators import ValidacionExcelError
from src.excel.result_writer import (
    guardar_resultado_error_validacion,
    guardar_resultados_ordenes,
)
from src.excel.resumen_writer import actualizar_resumen_orden
from src.excel.state_manager import (
    crear_excel_trabajo_pendientes,
    actualizar_estado_orden,
    obtener_fila_excel_por_id,
)
from src.flows.inicio_cotizacion import iniciar_cotizacion_con_texto
from src.flows.proveedor import seleccionar_proveedor_por_codigo
from src.flows.servicio import seleccionar_codigo_bien_servicio
from src.flows.posiciones import editar_posiciones_requerimiento
from src.flows.asignacion import asignar_imputacion_a_posiciones
from src.flows.imputacion_multiple import subir_templates_imputacion_multiple
from src.flows.fecha_adjunto import seleccionar_fecha_y_adjuntar
from src.flows.enviar import enviar_cotizacion
from src.flows.panel_cotizaciones import verificar_cotizacion_en_panel
from src.flows.mandar_cotizar import mandar_a_cotizar_orden
from src.flows.cerrar_cotizacion import cerrar_cotizacion_desde_panel
from src.flows.asignar_comparativo import asignar_comparativo_desde_panel
from src.utils.app_paths import (
    configurar_directorio_ejecucion,
    crear_directorio_ejecucion,
)
from src.utils.debug import preparar_carpetas, guardar_evidencia_error
from src.utils.execution_logger import capturar_salida_ejecucion
from src.utils.rpa_errors import PasoRPAError
from src.utils.user_errors import construir_mensaje_usuario, construir_detalle_tecnico


EventCallback = Callable[[dict[str, Any]], None]


@dataclass
class ResultadoEjecucion:
    estado: str
    directorio_ejecucion: str
    ruta_log: str
    ruta_resultado: str
    ruta_excel: str
    fecha_inicio: str
    fecha_fin: str
    total: int
    correctas: int
    errores: int
    mensaje: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _emitir(callback: EventCallback | None, tipo: str, **datos: Any) -> None:
    if callback is None:
        return

    evento = {"type": tipo, **datos}

    try:
        callback(evento)
    except Exception:
        pass


def validar_archivo_excel(ruta_excel: str | Path) -> Path:
    ruta_excel = Path(ruta_excel).expanduser().resolve()

    if not ruta_excel.exists():
        raise FileNotFoundError(f"No existe el archivo Excel: {ruta_excel}")

    if not ruta_excel.is_file():
        raise ValueError(f"La ruta no es un archivo: {ruta_excel}")

    extensiones_validas = {".xlsx", ".xlsm", ".xls"}

    if ruta_excel.suffix.lower() not in extensiones_validas:
        raise ValueError(
            f"Extensión no válida: {ruta_excel.suffix}. Use .xlsx, .xlsm o .xls."
        )

    return ruta_excel


def mostrar_orden_validada(orden) -> None:
    tipo_web = "BIEN" if orden.tipo_servicio == "B" else "SERVICIO"

    print("------------------------------------------------------------")
    print(f"ID orden: {orden.id_orden}")
    print(f"Proveedor: {orden.proveedor}")
    print(f"Texto: {orden.texto}")
    print(f"Valor: {orden.valor}")
    print(f"Moneda: {orden.moneda}")
    print("CECO/imputación: POR POSICIÓN - filas agrupadas de hoja Ordenes")
    print(f"Cuenta: {orden.cuenta}")
    print(f"Servicio/Código: {orden.servicio}")
    print(f"Tipo servicio: {orden.tipo_servicio} ({tipo_web})")
    print(f"Posiciones a agregar: {orden.posiciones}")
    print(f"Tipo imputación: {orden.tipo_imputacion}")
    print("TEMPLATE: definido por posición en cada fila de Ordenes")
    print(f"Cantidad de adjuntos generales: {len(orden.adjuntos)}")

    for indice, archivo in enumerate(orden.adjuntos, start=1):
        tamano_mb = archivo.stat().st_size / (1024 * 1024)
        print(f"Adjunto general {indice}: {archivo} ({tamano_mb:.2f} MB)")

    for indice, posicion in enumerate(orden.posiciones_detalle, start=1):
        print(
            f"Posición {indice}: "
            f"TEXTO_BREVE='{posicion.texto_breve}', "
            f"VALOR={posicion.valor}, "
            f"CECO={posicion.ceco}, "
            f"TIPO_IMPUTACION={orden.tipo_imputacion}, "
            f"TEMPLATE={posicion.template if posicion.template else 'VACÍO'}"
        )


def _emitir_etapa(
    callback: EventCallback | None,
    orden,
    etapa: str,
    mensaje: str,
) -> None:
    _emitir(
        callback,
        "order_stage",
        id_orden=str(orden.id_orden),
        stage=etapa,
        message=mensaje,
    )


def procesar_orden(
    page,
    orden,
    indice: int,
    total: int,
    ruta_excel: Path,
    callback: EventCallback | None = None,
):
    print("")
    print("============================================================")
    print(f"Procesando orden {indice}/{total}: ID_ORDEN={orden.id_orden}")
    print("============================================================")

    _emitir_etapa(callback, orden, "inicio", "Abriendo formulario de cotización")

    page.goto(
        URL_SISTEMA,
        wait_until="domcontentloaded",
        timeout=60_000,
    )

    _emitir_etapa(callback, orden, "cotización", "Ingresando nombre de cotización")
    print("Iniciando cotización...")
    iniciar_cotizacion_con_texto(page, orden)
    print("Campo TEXTO completado.")

    _emitir_etapa(callback, orden, "proveedor", "Seleccionando proveedor")
    print("Seleccionando proveedor...")
    seleccionar_proveedor_por_codigo(page, orden.proveedor)
    print("Proveedor seleccionado correctamente.")

    _emitir_etapa(callback, orden, "bien_servicio", "Seleccionando bien o servicio")
    print("Seleccionando código de bien/servicio...")
    seleccionar_codigo_bien_servicio(page, orden)
    print("Código de bien/servicio agregado correctamente.")

    _emitir_etapa(callback, orden, "posiciones", "Editando posiciones")
    print("Editando posiciones del requerimiento...")
    editar_posiciones_requerimiento(page, orden)
    print("Posiciones editadas correctamente.")

    _emitir_etapa(callback, orden, "imputación", "Asignando imputación")
    print("Asignando imputación a posiciones...")
    asignar_imputacion_a_posiciones(page, orden)
    print("Imputación asignada correctamente.")

    _emitir_etapa(callback, orden, "templates", "Procesando templates")
    print("Revisando templates de imputación múltiple...")
    subir_templates_imputacion_multiple(page, orden)
    print("Templates de imputación múltiple procesados correctamente.")

    _emitir_etapa(callback, orden, "adjuntos", "Colocando fecha y adjuntos")
    print("Colocando fecha y adjuntando archivos...")
    seleccionar_fecha_y_adjuntar(page, orden.adjuntos)
    print("Fecha y archivos completados correctamente.")

    _emitir_etapa(callback, orden, "enviar", "Enviando cotización")
    print("Enviando cotización...")
    enviar_cotizacion(page)
    print("Cotización enviada correctamente.")

    _emitir_etapa(callback, orden, "panel", "Verificando panel de cotizaciones")
    print("Verificando cotización en Panel de cotizaciones...")
    info_panel = verificar_cotizacion_en_panel(page, orden)

    _emitir_etapa(callback, orden, "mandar_cotizar", "Mandando a cotizar")
    print("Iniciando fase: mandar a cotizar...")
    info_mandar = mandar_a_cotizar_orden(page, orden)
    print(f"Mandar a cotizar OK: {info_mandar.mensaje}")

    _emitir_etapa(callback, orden, "cierre", "Cerrando cotización")
    print("Iniciando fase: cerrar cotización desde DURACIÓN...")
    info_cierre = cerrar_cotizacion_desde_panel(page, orden)
    print(f"Cierre OK: {info_cierre.mensaje}")

    _emitir_etapa(callback, orden, "comparativo", "Asignando comparativo")
    print("Iniciando fase: asignar comparativo y finalizar...")
    info_asignacion = asignar_comparativo_desde_panel(page, orden)
    print(f"Comparativo OK: {info_asignacion.mensaje}")

    actualizar_resumen_orden(
        ruta_excel,
        orden.id_orden,
        getattr(info_asignacion, "resumen", ""),
    )
    print(f"RESUMEN actualizado para ID_ORDEN={orden.id_orden}")
    print("Cotización verificada en Panel de cotizaciones.")

    return info_panel


def construir_resultado_ok(
    indice: int,
    orden,
    ruta_excel: Path,
    info_panel=None,
) -> dict[str, Any]:
    return {
        "NRO": indice,
        "ID_ORDEN": orden.id_orden,
        "FILA_EXCEL": obtener_fila_excel_por_id(ruta_excel, orden.id_orden) or "",
        "ESTADO": "OK",
        "N_COTIZACION": getattr(info_panel, "numero_cotizacion", ""),
        "ESTADO_CBN": getattr(info_panel, "estado_cbn", ""),
        "MENSAJE_USUARIO": "Orden procesada correctamente.",
        "PASO_ERROR": "",
        "DETALLE_TECNICO": "",
        "LOG": "",
        "SCREENSHOT": "",
        "HTML": "",
        "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def construir_resultado_error(
    indice: int,
    orden,
    ruta_excel: Path,
    error: Exception,
    rutas: dict,
) -> dict[str, Any]:
    return {
        "NRO": indice,
        "ID_ORDEN": getattr(orden, "id_orden", ""),
        "FILA_EXCEL": obtener_fila_excel_por_id(
            ruta_excel,
            getattr(orden, "id_orden", ""),
        ) or "",
        "ESTADO": "ERROR",
        "N_COTIZACION": "",
        "ESTADO_CBN": "",
        "MENSAJE_USUARIO": construir_mensaje_usuario(error, orden),
        "PASO_ERROR": getattr(error, "paso", "error_app"),
        "DETALLE_TECNICO": "Consulte el archivo indicado en la columna LOG.",
        "LOG": str(rutas.get("log", "")),
        "SCREENSHOT": str(rutas.get("screenshot", "")),
        "HTML": str(rutas.get("html", "")),
        "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def validar_excel_sin_ejecutar(ruta_excel: str | Path) -> dict[str, Any]:
    ruta_excel = validar_archivo_excel(ruta_excel)
    ruta_trabajo: Path | None = None

    try:
        ruta_trabajo, pendientes = crear_excel_trabajo_pendientes(
            ruta_excel,
            solo_lectura=True,
        )

        if not pendientes:
            return {
                "ok": True,
                "pending_count": 0,
                "orders": [],
                "message": "El Excel es válido, pero no tiene órdenes pendientes.",
            }

        ordenes = leer_ordenes_excel(ruta_trabajo)

        return {
            "ok": True,
            "pending_count": len(ordenes),
            "orders": [
                {
                    "id_orden": str(orden.id_orden),
                    "texto": str(orden.texto),
                    "posiciones": int(orden.posiciones),
                    "tipo_servicio": str(orden.tipo_servicio),
                    "valor": str(orden.valor),
                    "moneda": str(orden.moneda),
                }
                for orden in ordenes
            ],
            "message": f"Excel válido. Órdenes pendientes: {len(ordenes)}.",
        }
    finally:
        if ruta_trabajo is not None:
            try:
                ruta_trabajo.unlink()
            except Exception:
                pass


def _guardar_json(ruta: Path, datos: dict[str, Any]) -> None:
    ruta.write_text(
        json.dumps(datos, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )


def _copiar_si_existe(origen: Path, destino: Path) -> None:
    try:
        if origen.exists() and origen.is_file():
            shutil.copy2(origen, destino)
    except Exception:
        pass



def _error_raiz(error: Exception) -> Exception:
    return getattr(error, "error_original", error)


def _es_navegador_cerrado(error: Exception) -> bool:
    if isinstance(error, NavegadorCerradoPorUsuarioError):
        return True

    texto = str(_error_raiz(error)).lower()
    indicadores = (
        "target page, context or browser has been closed",
        "target closed",
        "page has been closed",
        "browser has been closed",
        "context has been closed",
    )
    return any(indicador in texto for indicador in indicadores)


def _mensaje_navegador_cerrado(durante_orden: bool = False) -> str:
    if durante_orden:
        return (
            "El navegador fue cerrado mientras se procesaba una orden. "
            "La ejecución se detuvo para evitar errores en las órdenes siguientes."
        )

    return (
        "La ejecución se detuvo porque se cerró el navegador antes de completar "
        "el inicio de sesión. No se procesó ninguna orden."
    )


def ejecutar_rpa(
    ruta_excel: str | Path,
    navegador: str = "msedge",
    callback: EventCallback | None = None,
    cancelar_evento: Event | None = None,
    pausar_al_final: bool = False,
) -> ResultadoEjecucion:
    ruta_excel = validar_archivo_excel(ruta_excel)
    cancelar_evento = cancelar_evento or Event()

    fecha_inicio_dt = datetime.now()
    fecha_inicio = fecha_inicio_dt.isoformat(timespec="seconds")
    directorio = crear_directorio_ejecucion(ruta_excel)
    configurar_directorio_ejecucion(directorio)
    preparar_carpetas()

    entrada_original = directorio / f"entrada_original{ruta_excel.suffix.lower()}"

    directorio_inputs = directorio / "inputs"
    directorio_templates = directorio_inputs / "templates"
    ruta_manifest = directorio_inputs / "snapshot_manifest.json"

    metadata_entrada: dict[str, Any] | None = None

    _emitir(
        callback,
        "run_directory",
        path=str(directorio),
    )

    playwright = None
    contexto = None
    page = None
    resultados: list[dict[str, Any]] = []
    ruta_resultado: Path | None = None
    ruta_trabajo: Path | None = None
    estado_final = "ERROR"
    mensaje_final = ""
    total = 0

    with capturar_salida_ejecucion(directorio, callback) as ruta_log:
        try:
            metadata_entrada = crear_snapshot_estable(
                ruta_excel,
                entrada_original,
            )

            metadata_entrada["tipo"] = "excel_principal"

            guardar_manifest_snapshots(
                ruta_manifest,
                metadata_entrada,
                [],
            )

            print(
                f"Snapshot de entrada creado: "
                f"{entrada_original}"
            )

            print(
                f"SHA256 entrada: "
                f"{metadata_entrada['sha256']}"
            )

            print("============================================================")
            print("RPA CBN - INICIO DE EJECUCIÓN")
            print(f"Archivo: {ruta_excel}")
            print(f"Directorio de ejecución: {directorio}")
            print(f"Navegador solicitado: {navegador}")
            print("============================================================")

            _emitir(callback, "validation_started", excel=str(ruta_excel))
            print("Leyendo y validando Excel...")

            ruta_trabajo, pendientes = crear_excel_trabajo_pendientes(entrada_original)

            if not pendientes:
                estado_final = "SIN_PENDIENTES"
                mensaje_final = "No hay órdenes pendientes con ESTADO_RPA = 0."
                print(mensaje_final)
                _emitir(callback, "no_pending", message=mensaje_final)
            else:
                ordenes = leer_ordenes_excel(ruta_trabajo)

                ordenes, metadata_templates = congelar_templates_ordenes(
                    ordenes,
                    directorio_templates,
                )

                # Validamos otra vez, pero ahora exactamente sobre
                # los templates congelados que utilizar? esta ejecuci?n.
                validar_ordenes_cargadas(
                    ordenes
                )

                if metadata_entrada is None:
                    raise RuntimeError(
                        "No existe metadata del snapshot principal."
                    )

                guardar_manifest_snapshots(
                    ruta_manifest,
                    metadata_entrada,
                    metadata_templates,
                )

                print(
                    f"Templates congelados correctamente: "
                    f"{len(metadata_templates)} archivo(s) ?nico(s)."
                )

                if not ordenes:
                    raise ValidacionExcelError(
                        "No se encontró ninguna orden pendiente válida en el Excel."
                    )

                total = len(ordenes)
                print(
                    f"Excel validado correctamente. "
                    f"Órdenes pendientes encontradas: {total}"
                )

                _emitir(
                    callback,
                    "validation_ok",
                    total=total,
                    orders=[
                        {
                            "id_orden": str(orden.id_orden),
                            "texto": str(orden.texto),
                            "posiciones": int(orden.posiciones),
                            "tipo_servicio": str(orden.tipo_servicio),
                            "valor": str(orden.valor),
                            "moneda": str(orden.moneda),
                        }
                        for orden in ordenes
                    ],
                )

                for orden in ordenes:
                    mostrar_orden_validada(orden)

                if cancelar_evento.is_set():
                    estado_final = "CANCELADA"
                    mensaje_final = "Ejecución cancelada antes de abrir el navegador."
                else:
                    _emitir(
                        callback,
                        "browser_opening",
                        browser=navegador,
                        message="Abriendo navegador...",
                    )

                    playwright, contexto, page = abrir_contexto(navegador=navegador)

                    page.goto(
                        URL_SISTEMA,
                        wait_until="domcontentloaded",
                        timeout=60_000,
                    )

                    _emitir(
                        callback,
                        "login_waiting",
                        message=(
                            "Verificando sesión. Si aparece el login, "
                            "complete el acceso en el navegador."
                        ),
                    )
                    esperar_login_manual(page)
                    print("Ya estamos dentro del sistema.")
                    _emitir(callback, "login_ok", message="Sesión de CBN lista.")

                    for indice, orden in enumerate(ordenes, start=1):
                        if cancelar_evento.is_set():
                            estado_final = "CANCELADA"
                            mensaje_final = (
                                "Cancelación solicitada. "
                                "No se iniciaron más órdenes."
                            )
                            print(mensaje_final)
                            break

                        _emitir(
                            callback,
                            "order_started",
                            id_orden=str(orden.id_orden),
                            index=indice,
                            total=total,
                            message=f"Procesando orden {indice} de {total}",
                        )

                        try:
                            print(
                                f"Fila Excel: "
                                f"{obtener_fila_excel_por_id(ruta_excel, orden.id_orden)}"
                            )

                            info_panel = procesar_orden(
                                page=page,
                                orden=orden,
                                indice=indice,
                                total=total,
                                ruta_excel=ruta_excel,
                                callback=callback,
                            )

                            resultado_ok = construir_resultado_ok(
                                indice,
                                orden,
                                ruta_excel,
                                info_panel,
                            )
                            resultados.append(resultado_ok)

                            actualizar_estado_orden(
                                ruta_excel,
                                orden.id_orden,
                                1,
                            )

                            print(
                                f"ESTADO_RPA actualizado a 1 "
                                f"para ID_ORDEN={orden.id_orden}"
                            )

                            _emitir(
                                callback,
                                "order_ok",
                                id_orden=str(orden.id_orden),
                                index=indice,
                                total=total,
                                quote_number=resultado_ok.get("N_COTIZACION", ""),
                                message="Orden procesada correctamente.",
                            )

                        except PasoRPAError as error:
                            navegador_cerrado = _es_navegador_cerrado(error)
                            mensaje_usuario = (
                                _mensaje_navegador_cerrado(durante_orden=True)
                                if navegador_cerrado
                                else construir_mensaje_usuario(error, orden)
                            )

                            print("")
                            print("No se pudo completar la orden.")
                            print(f"ID_ORDEN: {orden.id_orden}")
                            print(f"Mensaje: {mensaje_usuario}")
                            print(f"Paso: {error.paso}")

                            prefijo = (
                                f"orden_{orden.id_orden}_"
                                f"{error.paso.replace(' ', '_')}"
                            )
                            rutas = guardar_evidencia_error(
                                page,
                                error,
                                prefijo,
                            )

                            resultado_error = construir_resultado_error(
                                indice,
                                orden,
                                ruta_excel,
                                error,
                                rutas,
                            )
                            resultado_error["MENSAJE_USUARIO"] = mensaje_usuario

                            if navegador_cerrado:
                                resultado_error["PASO_ERROR"] = "navegador_cerrado"

                            resultados.append(resultado_error)
                            actualizar_estado_orden(
                                ruta_excel,
                                orden.id_orden,
                                2,
                            )

                            _emitir(
                                callback,
                                "order_error",
                                id_orden=str(orden.id_orden),
                                index=indice,
                                total=total,
                                stage=(
                                    "navegador_cerrado"
                                    if navegador_cerrado
                                    else error.paso
                                ),
                                message=mensaje_usuario,
                                evidence={
                                    clave: str(valor)
                                    for clave, valor in rutas.items()
                                },
                            )

                            print(f"Detalle técnico guardado en: {rutas.get('log', '')}")

                            if navegador_cerrado:
                                estado_final = "CANCELADA"
                                mensaje_final = mensaje_usuario
                                break

                            print("Se continúa con la siguiente orden.")

                        except Exception as error:
                            navegador_cerrado = _es_navegador_cerrado(error)
                            mensaje_usuario = (
                                _mensaje_navegador_cerrado(durante_orden=True)
                                if navegador_cerrado
                                else construir_mensaje_usuario(error, orden)
                            )

                            print("")
                            print("No se pudo completar la orden.")
                            print(f"ID_ORDEN: {orden.id_orden}")
                            print(f"Mensaje: {mensaje_usuario}")

                            prefijo = (
                                f"orden_{orden.id_orden}_navegador_cerrado"
                                if navegador_cerrado
                                else f"orden_{orden.id_orden}_error_app"
                            )
                            rutas = guardar_evidencia_error(
                                page,
                                error,
                                prefijo,
                            )

                            resultado_error = construir_resultado_error(
                                indice,
                                orden,
                                ruta_excel,
                                error,
                                rutas,
                            )
                            resultado_error["MENSAJE_USUARIO"] = mensaje_usuario

                            if navegador_cerrado:
                                resultado_error["PASO_ERROR"] = "navegador_cerrado"

                            resultados.append(resultado_error)
                            actualizar_estado_orden(
                                ruta_excel,
                                orden.id_orden,
                                2,
                            )

                            _emitir(
                                callback,
                                "order_error",
                                id_orden=str(orden.id_orden),
                                index=indice,
                                total=total,
                                stage=(
                                    "navegador_cerrado"
                                    if navegador_cerrado
                                    else "error_app"
                                ),
                                message=mensaje_usuario,
                                evidence={
                                    clave: str(valor)
                                    for clave, valor in rutas.items()
                                },
                            )

                            print(f"Detalle técnico guardado en: {rutas.get('log', '')}")

                            if navegador_cerrado:
                                estado_final = "CANCELADA"
                                mensaje_final = mensaje_usuario
                                break

                            print("Se continúa con la siguiente orden.")

                        completadas = len(resultados)
                        _emitir(
                            callback,
                            "progress",
                            completed=completadas,
                            total=total,
                            value=(completadas / total) if total else 0.0,
                        )

                    if estado_final != "CANCELADA":
                        estado_final = "COMPLETADA"
                        mensaje_final = "Ejecución terminada."

            if resultados:
                ruta_resultado = guardar_resultados_ordenes(
                    ruta_excel,
                    resultados,
                )

            correctas = sum(
                1 for item in resultados if item.get("ESTADO") == "OK"
            )
            errores = sum(
                1 for item in resultados if item.get("ESTADO") != "OK"
            )

            print("")
            print("============================================================")
            print(mensaje_final or "Proceso terminado.")
            print(f"Órdenes OK: {correctas}")
            print(f"Órdenes con error: {errores}")

            if ruta_resultado:
                print(f"Resultado final guardado en: {ruta_resultado}")

            print("============================================================")

            if pausar_al_final and page is not None:
                print("El navegador quedará pausado para revisión.")
                page.pause()

        except NavegadorCerradoPorUsuarioError as error:
            estado_final = "CANCELADA"
            mensaje_final = _mensaje_navegador_cerrado(durante_orden=False)

            rutas = guardar_evidencia_error(
                page,
                error,
                "navegador_cerrado_login",
            )

            print("")
            print(mensaje_final)
            print(f"Detalle técnico guardado en: {rutas.get('log', '')}")

            _emitir(
                callback,
                "run_cancelled",
                message=mensaje_final,
                stage="login",
                evidence={
                    clave: str(valor)
                    for clave, valor in rutas.items()
                },
            )

        except FileNotFoundError as error:
            estado_final = "ERROR"
            mensaje_final = str(error)
            print("Error de archivo.")
            print(error)

            rutas = guardar_evidencia_error(page, error, "archivo")
            resultado = [
                {
                    "NRO": 1,
                    "ID_ORDEN": "",
                    "ESTADO": "ERROR_ARCHIVO",
                    "N_COTIZACION": "",
                    "ESTADO_CBN": "",
                    "MENSAJE_USUARIO": str(error),
                    "PASO_ERROR": "archivo",
                    "DETALLE_TECNICO": str(error),
                    "LOG": str(rutas.get("log", "")),
                    "SCREENSHOT": str(rutas.get("screenshot", "")),
                    "HTML": str(rutas.get("html", "")),
                    "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            ]
            resultados.extend(resultado)
            ruta_resultado = guardar_resultados_ordenes(
                ruta_excel,
                resultado,
            )

            _emitir(
                callback,
                "fatal_error",
                message=mensaje_final,
                stage="archivo",
            )

        except ValidacionExcelError as error:
            estado_final = "ERROR_VALIDACION"
            mensaje_final = str(error)
            print("Error de validación del Excel.")
            print(error)

            ruta_resultado = guardar_resultado_error_validacion(
                ruta_excel,
                str(error),
            )
            guardar_evidencia_error(
                page,
                error,
                "validacion_excel",
            )

            _emitir(
                callback,
                "fatal_error",
                message=mensaje_final,
                stage="validacion_excel",
            )

        except Exception as error:
            estado_final = "ERROR"
            mensaje_final = construir_mensaje_usuario(error, None)

            rutas = guardar_evidencia_error(page, error, "error_app")

            print("La ejecución no pudo completarse.")
            print(mensaje_final)
            print(f"Detalle técnico guardado en: {rutas.get('log', '')}")
            resultado = [
                {
                    "NRO": 1,
                    "ID_ORDEN": "",
                    "ESTADO": "ERROR_APP",
                    "N_COTIZACION": "",
                    "ESTADO_CBN": "",
                    "MENSAJE_USUARIO": mensaje_final,
                    "PASO_ERROR": "error_app",
                    "DETALLE_TECNICO": "Consulte el archivo indicado en la columna LOG.",
                    "LOG": str(rutas.get("log", "")),
                    "SCREENSHOT": str(rutas.get("screenshot", "")),
                    "HTML": str(rutas.get("html", "")),
                    "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            ]
            resultados.extend(resultado)
            ruta_resultado = guardar_resultados_ordenes(
                ruta_excel,
                resultado,
            )

            _emitir(
                callback,
                "fatal_error",
                message=mensaje_final,
                stage="error_app",
            )

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

            if ruta_trabajo is not None:
                try:
                    ruta_trabajo.unlink()
                except Exception:
                    pass

            _copiar_si_existe(
                ruta_excel,
                directorio / f"entrada_actualizada{ruta_excel.suffix.lower()}",
            )

    fecha_fin = datetime.now().isoformat(timespec="seconds")
    correctas = sum(1 for item in resultados if item.get("ESTADO") == "OK")
    errores = sum(1 for item in resultados if item.get("ESTADO") != "OK")

    resultado_ejecucion = ResultadoEjecucion(
        estado=estado_final,
        directorio_ejecucion=str(directorio),
        ruta_log=str(ruta_log),
        ruta_resultado=str(ruta_resultado or ""),
        ruta_excel=str(ruta_excel),
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
        total=total,
        correctas=correctas,
        errores=errores,
        mensaje=mensaje_final,
    )

    _guardar_json(
        directorio / "resumen.json",
        resultado_ejecucion.to_dict(),
    )

    _emitir(
        callback,
        "run_finished",
        result=resultado_ejecucion.to_dict(),
    )

    configurar_directorio_ejecucion(None)

    return resultado_ejecucion
