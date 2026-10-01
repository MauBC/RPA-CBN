from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
from tkinter import filedialog, messagebox, ttk
from typing import Any
import os
import shutil
import subprocess
import sys

import customtkinter as ctk
from PIL import Image

from src.core.runner import ejecutar_rpa, validar_excel_sin_ejecutar
from src.utils.app_config import cargar_configuracion, guardar_configuracion
from src.excel.sync_status import (
    EstadoSincronizacionExcel,
    ResumenSincronizacionExcel,
    obtener_estado_sincronizacion_excel,
)
from src.utils.app_paths import (
    obtener_directorio_app,
    obtener_directorio_perfil,
)


NAVEGADORES_UI = {
    "Microsoft Edge": "msedge",
    "Google Chrome": "chrome",
}


EXTENSIONES_LOGO = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
NOMBRES_LOGO_PRIORITARIOS = (
    "RANSA.png",
    "RANSA.webp",
    "RANSA.jpg",
    "RANSA.jpeg",
    "ransa.png",
    "ransa.webp",
    "ransa.jpg",
    "ransa.jpeg",
    "logo_ransa.png",
    "Ransalogo.png",
)


def _buscar_logo_ransa(assets_dir: Path) -> Path | None:
    """Busca el logo corporativo sin confundirlo con el ícono del ejecutable."""
    for nombre in NOMBRES_LOGO_PRIORITARIOS:
        candidato = assets_dir / nombre
        if candidato.is_file():
            return candidato

    if not assets_dir.exists():
        return None

    nombres_validos = {"ransa", "logo_ransa", "ransalogo", "logo-ransa"}

    for candidato in sorted(assets_dir.iterdir()):
        if (
            candidato.is_file()
            and candidato.suffix.lower() in EXTENSIONES_LOGO
            and candidato.stem.lower() in nombres_validos
        ):
            return candidato

    return None


def _calcular_tamano_logo(imagen: Image.Image) -> tuple[int, int]:
    """Conserva la proporción del logo dentro del encabezado."""
    ancho_original, alto_original = imagen.size

    if ancho_original <= 0 or alto_original <= 0:
        return 130, 48

    ancho_maximo = 145
    alto_maximo = 55
    escala = min(ancho_maximo / ancho_original, alto_maximo / alto_original)

    return (
        max(1, round(ancho_original * escala)),
        max(1, round(alto_original * escala)),
    )


def _abrir_ruta(ruta: str | Path) -> None:
    ruta = Path(ruta)

    if not ruta.exists():
        raise FileNotFoundError(f"No existe la ruta: {ruta}")

    if os.name == "nt":
        os.startfile(str(ruta))  # type: ignore[attr-defined]
        return

    comando = ["open", str(ruta)] if sys.platform == "darwin" else ["xdg-open", str(ruta)]
    subprocess.Popen(comando)


def _formatear_importe(valor: Any) -> str:
    texto = str(valor or "").strip().replace(",", "")

    if not texto:
        return ""

    try:
        numero = Decimal(texto)
    except (InvalidOperation, ValueError):
        return texto

    return f"{numero:,.2f}"


def _resumen_importes(ordenes: list[dict[str, Any]]) -> str:
    totales: dict[str, Decimal] = {}

    for orden in ordenes:
        moneda = str(orden.get("moneda", "") or "").strip().upper()
        valor_texto = str(orden.get("valor", "0") or "0").replace(",", "")

        if not moneda:
            continue

        try:
            valor = Decimal(valor_texto)
        except (InvalidOperation, ValueError):
            continue

        totales[moneda] = totales.get(moneda, Decimal("0")) + valor

    partes = [f"{moneda} {_formatear_importe(valor)}" for moneda, valor in sorted(totales.items())]
    return " | ".join(partes)


def _color_estado_sincronizacion(
    estado: EstadoSincronizacionExcel,
):
    if (
        estado
        == EstadoSincronizacionExcel.SINCRONIZADO
    ):
        return (
            "#2E7D32",
            "#66BB6A",
        )

    if estado in {
        EstadoSincronizacionExcel.EXCEL_OCUPADO,
        EstadoSincronizacionExcel.PENDIENTE,
    }:
        return (
            "#C66A00",
            "#FFB74D",
        )

    return (
        "#C62828",
        "#EF5350",
    )


class VentanaPrincipal(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self._directorio_app = obtener_directorio_app()
        self._assets_dir = self._directorio_app / "assets"

        self.title("RPA CBN - Automatización de Cotizaciones")
        self.geometry("1180x790")
        self.minsize(1000, 700)

        # Ícono propio de la aplicación. Se cargará automáticamente cuando
        # exista assets/rpa_cbn.ico.
        self._ruta_icono_app = self._assets_dir / "rpa_cbn.ico"

        if self._ruta_icono_app.exists():
            try:
                self.iconbitmap(str(self._ruta_icono_app))
            except Exception:
                # La ausencia o incompatibilidad del ícono no debe impedir
                # que el RPA abra.
                pass

        # Logo de Ransa mostrado dentro del encabezado. Puede ser PNG, JPG,
        # JPEG, WebP o BMP. Es independiente de assets/rpa_cbn.ico.
        self.logo_ransa: ctk.CTkImage | None = None
        ruta_logo_ransa = _buscar_logo_ransa(self._assets_dir)

        if ruta_logo_ransa is not None:
            try:
                imagen_logo = Image.open(ruta_logo_ransa).convert("RGBA")
                tamano_logo = _calcular_tamano_logo(imagen_logo)
                self.logo_ransa = ctk.CTkImage(
                    light_image=imagen_logo,
                    dark_image=imagen_logo,
                    size=tamano_logo,
                )
            except Exception:
                # Un logo faltante o dañado no debe bloquear el RPA.
                self.logo_ransa = None

        self._cola: Queue[dict[str, Any]] = Queue()
        self._hilo: Thread | None = None
        self._cancelar = Event()
        self._ejecutando = False
        self._directorio_actual = ""
        self._resultado_actual = ""
        self._filas_tabla: dict[str, str] = {}
        self._datos_financieros: dict[str, tuple[str, str]] = {}

        # CP9 - Estado independiente de sincronización Excel.
        # La consulta se ejecuta fuera del hilo de Tkinter para que
        # un lock temporal del journal nunca congele la interfaz.
        self._consulta_sync_activa = False
        self._intervalo_sync_ms = 2500

        self._config = cargar_configuracion()

        self._crear_interfaz()
        self._cargar_config_en_interfaz()

        self.after(100, self._procesar_cola)

        # Primer diagnóstico poco después de mostrar la ventana.
        self.after(
            250,
            self._ciclo_estado_sincronizacion_excel,
        )

        self.protocol("WM_DELETE_WINDOW", self._cerrar_aplicacion)

    def _crear_interfaz(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        encabezado = ctk.CTkFrame(self, corner_radius=0)
        encabezado.grid(row=0, column=0, sticky="ew")
        encabezado.grid_columnconfigure(1, weight=1)

        if self.logo_ransa is not None:
            self.etiqueta_logo_ransa = ctk.CTkLabel(
                encabezado,
                text="",
                image=self.logo_ransa,
            )
            self.etiqueta_logo_ransa.grid(
                row=0,
                column=0,
                rowspan=2,
                padx=(24, 14),
                pady=12,
                sticky="w",
            )

        columna_texto = 1 if self.logo_ransa is not None else 0
        padding_izquierdo = (0, 24) if self.logo_ransa is not None else (24, 24)

        ctk.CTkLabel(
            encabezado,
            text="RPA CBN",
            font=ctk.CTkFont(size=26, weight="bold"),
        ).grid(
            row=0,
            column=columna_texto,
            padx=padding_izquierdo,
            pady=(18, 2),
            sticky="w",
        )

        ctk.CTkLabel(
            encabezado,
            text="Creación y procesamiento de cotizaciones",
            text_color=("gray35", "gray70"),
        ).grid(
            row=1,
            column=columna_texto,
            padx=padding_izquierdo,
            pady=(0, 16),
            sticky="w",
        )

        configuracion = ctk.CTkFrame(self)
        configuracion.grid(row=1, column=0, padx=20, pady=(16, 10), sticky="ew")
        configuracion.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(configuracion, text="Archivo Excel").grid(
            row=0, column=0, padx=(16, 8), pady=14, sticky="w"
        )

        self.entrada_excel = ctk.CTkEntry(
            configuracion,
            placeholder_text="Seleccione el archivo de órdenes...",
        )
        self.entrada_excel.grid(
            row=0, column=1, padx=8, pady=14, sticky="ew"
        )

        self.boton_examinar = ctk.CTkButton(
            configuracion,
            text="Examinar",
            width=105,
            command=self._seleccionar_excel,
        )
        self.boton_examinar.grid(row=0, column=2, padx=8, pady=14)

        ctk.CTkLabel(configuracion, text="Navegador").grid(
            row=0, column=3, padx=(18, 8), pady=14
        )

        self.selector_navegador = ctk.CTkOptionMenu(
            configuracion,
            values=list(NAVEGADORES_UI.keys()),
            width=155,
        )
        self.selector_navegador.grid(
            row=0, column=4, padx=(8, 16), pady=14
        )

        self.etiqueta_sync_excel = ctk.CTkLabel(
            configuracion,
            text="Excel: sin archivo seleccionado.",
            anchor="w",
            font=ctk.CTkFont(
                size=12,
                weight="bold",
            ),
            text_color=(
                "gray40",
                "gray65",
            ),
        )

        self.etiqueta_sync_excel.grid(
            row=1,
            column=1,
            columnspan=4,
            padx=(8, 16),
            pady=(0, 10),
            sticky="ew",
        )

        acciones = ctk.CTkFrame(self, fg_color="transparent")
        acciones.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        acciones.grid_columnconfigure(5, weight=1)

        self.boton_validar = ctk.CTkButton(
            acciones,
            text="Validar Excel",
            command=self._validar_excel,
        )
        self.boton_validar.grid(row=0, column=0, padx=(0, 8))

        self.boton_iniciar = ctk.CTkButton(
            acciones,
            text="Iniciar procesamiento",
            command=self._iniciar,
        )
        self.boton_iniciar.grid(row=0, column=1, padx=8)

        self.boton_detener = ctk.CTkButton(
            acciones,
            text="Detener después de la orden",
            fg_color="#C62828",
            hover_color="#8E0000",
            text_color="#FFFFFF",
            text_color_disabled="#FFFFFF",
            state="disabled",
            command=self._solicitar_detencion,
        )
        self.boton_detener.grid(row=0, column=2, padx=8)

        self.boton_sesion = ctk.CTkButton(
            acciones,
            text="Restablecer sesión",
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            command=self._restablecer_sesion,
        )
        self.boton_sesion.grid(row=0, column=3, padx=8)

        self.boton_carpeta = ctk.CTkButton(
            acciones,
            text="Abrir ejecución",
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            state="disabled",
            command=self._abrir_ejecucion,
        )
        self.boton_carpeta.grid(row=0, column=4, padx=8)

        cuerpo = ctk.CTkFrame(self)
        cuerpo.grid(row=3, column=0, padx=20, pady=(0, 14), sticky="nsew")
        cuerpo.grid_columnconfigure(0, weight=1)
        cuerpo.grid_rowconfigure(2, weight=1)
        cuerpo.grid_rowconfigure(4, weight=1)

        self.etiqueta_estado = ctk.CTkLabel(
            cuerpo,
            text="Seleccione un Excel para comenzar.",
            anchor="w",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.etiqueta_estado.grid(
            row=0, column=0, padx=16, pady=(14, 6), sticky="ew"
        )

        self.barra_progreso = ctk.CTkProgressBar(cuerpo)
        self.barra_progreso.grid(
            row=1, column=0, padx=16, pady=(0, 12), sticky="ew"
        )
        self.barra_progreso.set(0)

        tabla_frame = ctk.CTkFrame(cuerpo)
        tabla_frame.grid(
            row=2, column=0, padx=16, pady=(0, 12), sticky="nsew"
        )
        tabla_frame.grid_columnconfigure(0, weight=1)
        tabla_frame.grid_rowconfigure(0, weight=1)

        columnas = ("id", "importe", "moneda", "estado", "etapa", "detalle")
        self.tabla = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings",
            height=8,
        )
        self.tabla.heading("id", text="ID orden")
        self.tabla.heading("importe", text="Importe")
        self.tabla.heading("moneda", text="Moneda")
        self.tabla.heading("estado", text="Estado")
        self.tabla.heading("etapa", text="Etapa")
        self.tabla.heading("detalle", text="Detalle")
        self.tabla.column("id", width=85, anchor="center")
        self.tabla.column("importe", width=125, anchor="e")
        self.tabla.column("moneda", width=80, anchor="center")
        self.tabla.column("estado", width=110, anchor="center")
        self.tabla.column("etapa", width=150)
        self.tabla.column("detalle", width=430)

        scroll_tabla = ttk.Scrollbar(
            tabla_frame,
            orient="vertical",
            command=self.tabla.yview,
        )
        self.tabla.configure(yscrollcommand=scroll_tabla.set)

        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll_tabla.grid(row=0, column=1, sticky="ns")

        log_header = ctk.CTkFrame(cuerpo, fg_color="transparent")
        log_header.grid(row=3, column=0, padx=16, sticky="ew")
        log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            log_header,
            text="Registro de ejecución",
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        self.boton_resultado = ctk.CTkButton(
            log_header,
            text="Abrir resultado Excel",
            width=150,
            state="disabled",
            command=self._abrir_resultado,
        )
        self.boton_resultado.grid(row=0, column=1, pady=(0, 6))

        self.texto_log = ctk.CTkTextbox(cuerpo, wrap="word")
        self.texto_log.grid(
            row=4, column=0, padx=16, pady=(0, 16), sticky="nsew"
        )
        self.texto_log.configure(state="disabled")

    def _cargar_config_en_interfaz(self) -> None:
        ultimo_excel = str(self._config.get("ultimo_excel", "") or "")

        if ultimo_excel:
            self.entrada_excel.insert(0, ultimo_excel)

        navegador = str(self._config.get("navegador", "msedge"))

        for etiqueta, canal in NAVEGADORES_UI.items():
            if canal == navegador:
                self.selector_navegador.set(etiqueta)
                break

    def _guardar_config_actual(self) -> None:
        self._config["ultimo_excel"] = self.entrada_excel.get().strip()
        self._config["navegador"] = NAVEGADORES_UI.get(
            self.selector_navegador.get(),
            "msedge",
        )
        guardar_configuracion(self._config)

    def _seleccionar_excel(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de órdenes",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xlsm *.xls"),
                ("Todos los archivos", "*.*"),
            ],
        )

        if not ruta:
            return

        self.entrada_excel.delete(0, "end")
        self.entrada_excel.insert(0, ruta)
        self._guardar_config_actual()
        self.etiqueta_estado.configure(
            text="Excel seleccionado. Puede validarlo."
        )

        self._consultar_estado_sincronizacion_excel()

    def _obtener_excel(self) -> Path | None:
        texto = self.entrada_excel.get().strip()

        if not texto:
            messagebox.showwarning(
                "Archivo requerido",
                "Seleccione el archivo Excel que contiene las órdenes.",
            )
            return None

        return Path(texto)

    def _validar_excel(self) -> None:
        if self._ejecutando:
            return

        ruta = self._obtener_excel()

        if ruta is None:
            return

        self._guardar_config_actual()
        self._cambiar_controles(False)
        self.etiqueta_estado.configure(text="Validando Excel...")
        self._limpiar_tabla()

        def trabajo() -> None:
            try:
                resultado = validar_excel_sin_ejecutar(ruta)
                self._cola.put(
                    {
                        "type": "validation_preview",
                        "result": resultado,
                    }
                )
            except Exception as error:
                self._cola.put(
                    {
                        "type": "validation_preview_error",
                        "message": str(error),
                    }
                )

        self._hilo = Thread(target=trabajo, daemon=True)
        self._hilo.start()

    def _iniciar(self) -> None:
        if self._ejecutando:
            return

        ruta = self._obtener_excel()

        if ruta is None:
            return

        self._guardar_config_actual()
        self._ejecutando = True
        self._cancelar.clear()
        self._directorio_actual = ""
        self._resultado_actual = ""
        self._limpiar_tabla()
        self._limpiar_log()
        self.barra_progreso.set(0)
        self._cambiar_controles(False, ejecutando=True)

        navegador = NAVEGADORES_UI.get(
            self.selector_navegador.get(),
            "msedge",
        )

        def callback(evento: dict[str, Any]) -> None:
            self._cola.put(evento)

        def trabajo() -> None:
            try:
                ejecutar_rpa(
                    ruta_excel=ruta,
                    navegador=navegador,
                    callback=callback,
                    cancelar_evento=self._cancelar,
                )
            except Exception:
                self._cola.put(
                    {
                        "type": "fatal_error",
                        "message": (
                            "La aplicación no pudo completar la ejecución. "
                            "Revise la carpeta de evidencias para obtener el detalle técnico."
                        ),
                        "stage": "interfaz",
                    }
                )
                self._cola.put(
                    {
                        "type": "run_finished_unexpected",
                    }
                )

        self._hilo = Thread(target=trabajo, daemon=True)
        self._hilo.start()

    def _consultar_estado_sincronizacion_excel(
        self,
    ) -> None:
        """
        Ejecuta un diagnóstico de solo lectura en segundo plano.

        No sincroniza ni modifica DATA.xlsx.
        """
        if self._consulta_sync_activa:
            return

        ruta_texto = (
            self.entrada_excel.get().strip()
        )

        if not ruta_texto:
            self.etiqueta_sync_excel.configure(
                text=(
                    "Excel: sin archivo seleccionado."
                ),
                text_color=(
                    "gray40",
                    "gray65",
                ),
            )
            return

        self._consulta_sync_activa = True

        def trabajo() -> None:
            try:
                resultado = (
                    obtener_estado_sincronizacion_excel(
                        ruta_texto
                    )
                )

                self._cola.put(
                    {
                        "type": "excel_sync_status",
                        "excel": ruta_texto,
                        "result": resultado,
                    }
                )

            except Exception as error:
                self._cola.put(
                    {
                        "type": "excel_sync_status_error",
                        "excel": ruta_texto,
                        "message": str(error),
                    }
                )

            finally:
                self._cola.put(
                    {
                        "type": (
                            "excel_sync_status_finished"
                        ),
                    }
                )

        Thread(
            target=trabajo,
            daemon=True,
        ).start()

    def _ciclo_estado_sincronizacion_excel(
        self,
    ) -> None:
        """
        Refresco periódico no bloqueante del indicador Excel.
        """
        self._consultar_estado_sincronizacion_excel()

        self.after(
            self._intervalo_sync_ms,
            self._ciclo_estado_sincronizacion_excel,
        )

    def _mostrar_estado_sincronizacion_excel(
        self,
        resultado: ResumenSincronizacionExcel,
    ) -> None:
        color = _color_estado_sincronizacion(
            resultado.estado
        )

        self.etiqueta_sync_excel.configure(
            text=(
                f"Excel: {resultado.mensaje}"
            ),
            text_color=color,
        )

    def _solicitar_detencion(self) -> None:
        if not self._ejecutando:
            return

        self._cancelar.set()
        self.boton_detener.configure(state="disabled")
        self.etiqueta_estado.configure(
            text=(
                "Cancelación solicitada. "
                "El RPA no iniciará otra orden después de la actual."
            )
        )
        self._agregar_log(
            "Cancelación solicitada por el usuario. "
            "Se detendrá antes de la siguiente orden."
        )

    def _procesar_cola(self) -> None:
        try:
            while True:
                evento = self._cola.get_nowait()
                self._manejar_evento(evento)
        except Empty:
            pass

        self.after(100, self._procesar_cola)

    def _manejar_evento(self, evento: dict[str, Any]) -> None:
        tipo = evento.get("type")

        if tipo == "log":
            self._agregar_log(str(evento.get("message", "")))
            return

        if tipo == "excel_sync_status":
            ruta_evento = str(
                evento.get(
                    "excel",
                    "",
                )
            )

            ruta_actual = (
                self.entrada_excel.get().strip()
            )

            # Evita mostrar un resultado viejo si el usuario
            # seleccionó otro Excel mientras terminaba el hilo.
            if ruta_evento == ruta_actual:
                resultado = evento.get(
                    "result"
                )

                if isinstance(
                    resultado,
                    ResumenSincronizacionExcel,
                ):
                    self._mostrar_estado_sincronizacion_excel(
                        resultado
                    )

            return

        if tipo == "excel_sync_status_error":
            ruta_evento = str(
                evento.get(
                    "excel",
                    "",
                )
            )

            if (
                ruta_evento
                == self.entrada_excel.get().strip()
            ):
                self.etiqueta_sync_excel.configure(
                    text=(
                        "Excel: no se pudo consultar "
                        "el estado de sincronización."
                    ),
                    text_color=(
                        "#C62828",
                        "#EF5350",
                    ),
                )

            return

        if tipo == "excel_sync_status_finished":
            self._consulta_sync_activa = False
            return

        if tipo == "run_directory":
            self._directorio_actual = str(evento.get("path", ""))
            self.boton_carpeta.configure(state="normal")
            return

        if tipo == "validation_started":
            self.etiqueta_estado.configure(text="Validando el archivo Excel...")
            return

        if tipo == "validation_ok":
            total = int(evento.get("total", 0))
            ordenes_evento = list(evento.get("orders", []))
            resumen = _resumen_importes(ordenes_evento)
            texto_estado = f"Excel válido. Se procesarán {total} orden(es)."

            if resumen:
                texto_estado += f" Pendiente: {resumen}."

            self.etiqueta_estado.configure(text=texto_estado)

            for orden in ordenes_evento:
                self._actualizar_fila(
                    str(orden.get("id_orden", "")),
                    "Pendiente",
                    "",
                    str(orden.get("texto", "")),
                    valor=orden.get("valor", ""),
                    moneda=orden.get("moneda", ""),
                )
            return

        if tipo == "browser_opening":
            self.etiqueta_estado.configure(text="Abriendo navegador...")
            return

        if tipo == "login_waiting":
            self.etiqueta_estado.configure(text=str(evento.get("message", "")))
            return

        if tipo == "login_ok":
            self.etiqueta_estado.configure(text="Sesión iniciada. Procesando órdenes...")
            return

        if tipo == "order_started":
            id_orden = str(evento.get("id_orden", ""))
            self._actualizar_fila(
                id_orden,
                "Procesando",
                "Inicio",
                str(evento.get("message", "")),
            )
            self.etiqueta_estado.configure(
                text=str(evento.get("message", "Procesando orden..."))
            )
            return

        if tipo == "order_stage":
            id_orden = str(evento.get("id_orden", ""))
            self._actualizar_fila(
                id_orden,
                "Procesando",
                str(evento.get("stage", "")),
                str(evento.get("message", "")),
            )
            return

        if tipo == "order_ok":
            id_orden = str(evento.get("id_orden", ""))
            numero = str(evento.get("quote_number", "") or "")
            detalle = "Correcto"

            if numero:
                detalle = f"Cotización {numero}"

            self._actualizar_fila(
                id_orden,
                "Correcto",
                "Finalizado",
                detalle,
            )
            return

        if tipo == "order_error":
            id_orden = str(evento.get("id_orden", ""))
            self._actualizar_fila(
                id_orden,
                "Error",
                str(evento.get("stage", "")),
                str(evento.get("message", "")),
            )
            return

        if tipo == "progress":
            valor = float(evento.get("value", 0.0))
            self.barra_progreso.set(max(0.0, min(1.0, valor)))
            return

        if tipo == "no_pending":
            self.etiqueta_estado.configure(text=str(evento.get("message", "")))
            return

        if tipo == "run_cancelled":
            self.etiqueta_estado.configure(
                text=str(evento.get("message", "Ejecución detenida."))
            )
            return

        if tipo == "fatal_error":
            self.etiqueta_estado.configure(
                text=str(
                    evento.get(
                        "message",
                        "La ejecución terminó con error.",
                    )
                )
            )
            return

        if tipo == "run_finished":
            resultado = evento.get("result", {})
            self._finalizar_interfaz(resultado)
            return

        if tipo == "run_finished_unexpected":
            self._finalizar_interfaz({})
            return

        if tipo == "validation_preview":
            resultado = evento.get("result", {})
            self._mostrar_validacion(resultado)
            return

        if tipo == "validation_preview_error":
            mensaje = str(evento.get("message", "Error de validación"))
            self.etiqueta_estado.configure(text="El Excel contiene errores.")
            self._cambiar_controles(True)
            messagebox.showerror("Error de validación", mensaje)
            return

    def _mostrar_validacion(self, resultado: dict[str, Any]) -> None:
        self._limpiar_tabla()

        for orden in resultado.get("orders", []):
            self._actualizar_fila(
                str(orden.get("id_orden", "")),
                "Pendiente",
                "Validado",
                (
                    f"{orden.get('texto', '')} | "
                    f"{orden.get('posiciones', 0)} posición(es)"
                ),
                valor=orden.get("valor", ""),
                moneda=orden.get("moneda", ""),
            )

        ordenes_resultado = list(resultado.get("orders", []))
        resumen = _resumen_importes(ordenes_resultado)
        texto_estado = str(resultado.get("message", "Validación terminada."))

        if resumen:
            texto_estado += f" Total pendiente: {resumen}."

        self.etiqueta_estado.configure(text=texto_estado)
        self._cambiar_controles(True)

        messagebox.showinfo(
            "Validación terminada",
            str(resultado.get("message", "El archivo es válido.")),
        )

    def _finalizar_interfaz(self, resultado: dict[str, Any]) -> None:
        self._ejecutando = False
        self._cambiar_controles(True)
        self.boton_detener.configure(state="disabled")

        self._directorio_actual = str(
            resultado.get(
                "directorio_ejecucion",
                self._directorio_actual,
            )
            or self._directorio_actual
        )
        self._resultado_actual = str(resultado.get("ruta_resultado", "") or "")

        if self._directorio_actual:
            self.boton_carpeta.configure(state="normal")

        if self._resultado_actual:
            self.boton_resultado.configure(state="normal")

        estado = str(resultado.get("estado", "FINALIZADA"))
        correctas = int(resultado.get("correctas", 0) or 0)
        errores = int(resultado.get("errores", 0) or 0)

        self.etiqueta_estado.configure(
            text=(
                f"Ejecución {estado.lower()}. "
                f"Correctas: {correctas}. Errores: {errores}."
            )
        )

        if estado == "COMPLETADA":
            self.barra_progreso.set(1)

        # La ejecución pudo aplicar o generar nuevas entradas
        # del journal. Actualizamos el indicador cuanto antes.
        self._consultar_estado_sincronizacion_excel()

        if not resultado:
            messagebox.showerror(
                "No se pudo completar",
                (
                    "La aplicación no pudo completar la ejecución. "
                    "Revise la carpeta de evidencias para obtener el detalle técnico."
                ),
            )
            return

        mensaje = str(resultado.get("mensaje", "") or "").strip()
        detalle = (
            f"Estado: {estado}\n"
            f"Órdenes correctas: {correctas}\n"
            f"Órdenes con error: {errores}"
        )

        if mensaje:
            detalle += f"\n\n{mensaje}"

        detalle += "\n\nUse 'Abrir ejecución' para revisar logs y evidencias."

        if estado == "CANCELADA":
            messagebox.showwarning("Ejecución detenida", detalle)
        elif estado.startswith("ERROR"):
            messagebox.showerror("La ejecución terminó con error", detalle)
        else:
            messagebox.showinfo("Proceso terminado", detalle)

    def _actualizar_fila(
        self,
        id_orden: str,
        estado: str,
        etapa: str,
        detalle: str,
        valor: Any = "",
        moneda: Any = "",
    ) -> None:
        if not id_orden:
            return

        iid = self._filas_tabla.get(id_orden)

        valor_formateado = _formatear_importe(valor)
        moneda_texto = str(moneda or "").strip().upper()

        if valor_formateado or moneda_texto:
            valor_anterior, moneda_anterior = self._datos_financieros.get(
                id_orden,
                ("", ""),
            )
            self._datos_financieros[id_orden] = (
                valor_formateado or valor_anterior,
                moneda_texto or moneda_anterior,
            )

        valor_guardado, moneda_guardada = self._datos_financieros.get(
            id_orden,
            ("", ""),
        )
        valores = (
            id_orden,
            valor_guardado,
            moneda_guardada,
            estado,
            etapa,
            detalle,
        )

        if iid and self.tabla.exists(iid):
            self.tabla.item(iid, values=valores)
            return

        iid = self.tabla.insert("", "end", values=valores)
        self._filas_tabla[id_orden] = iid

    def _limpiar_tabla(self) -> None:
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        self._filas_tabla.clear()
        self._datos_financieros.clear()

    def _agregar_log(self, mensaje: str) -> None:
        if not mensaje:
            return

        self.texto_log.configure(state="normal")
        self.texto_log.insert("end", mensaje.rstrip() + "\n")
        self.texto_log.see("end")
        self.texto_log.configure(state="disabled")

    def _limpiar_log(self) -> None:
        self.texto_log.configure(state="normal")
        self.texto_log.delete("1.0", "end")
        self.texto_log.configure(state="disabled")

    def _cambiar_controles(
        self,
        habilitar: bool,
        ejecutando: bool = False,
    ) -> None:
        estado = "normal" if habilitar else "disabled"

        self.boton_examinar.configure(state=estado)
        self.boton_validar.configure(state=estado)
        self.boton_iniciar.configure(state=estado)
        self.selector_navegador.configure(state=estado)
        self.boton_sesion.configure(state=estado)

        if ejecutando:
            self.boton_detener.configure(state="normal")
        else:
            self.boton_detener.configure(state="disabled")

    def _abrir_ejecucion(self) -> None:
        try:
            _abrir_ruta(self._directorio_actual)
        except Exception as error:
            messagebox.showerror("No se pudo abrir", str(error))

    def _abrir_resultado(self) -> None:
        try:
            _abrir_ruta(self._resultado_actual)
        except Exception as error:
            messagebox.showerror("No se pudo abrir", str(error))

    def _restablecer_sesion(self) -> None:
        if self._ejecutando:
            return

        navegador = NAVEGADORES_UI.get(
            self.selector_navegador.get(),
            "msedge",
        )
        carpeta = obtener_directorio_perfil(navegador=navegador)

        confirmar = messagebox.askyesno(
            "Restablecer sesión",
            (
                f"Se eliminará la sesión guardada para {self.selector_navegador.get()}.\n\n"
                "En la próxima ejecución tendrá que iniciar sesión nuevamente.\n\n"
                "¿Desea continuar?"
            ),
        )

        if not confirmar:
            return

        try:
            shutil.rmtree(carpeta, ignore_errors=False)
            messagebox.showinfo(
                "Sesión restablecida",
                "La sesión se eliminó correctamente.",
            )
        except FileNotFoundError:
            messagebox.showinfo(
                "Sesión restablecida",
                "No había una sesión guardada.",
            )
        except Exception as error:
            messagebox.showerror(
                "No se pudo restablecer",
                (
                    "Cierre todas las ventanas de Edge/Chrome abiertas por el RPA "
                    f"e inténtelo nuevamente.\n\nDetalle: {error}"
                ),
            )

    def _cerrar_aplicacion(self) -> None:
        if self._ejecutando:
            messagebox.showwarning(
                "RPA en ejecución",
                (
                    "No cierre la aplicación mientras el RPA está procesando.\n\n"
                    "Use el botón 'Detener después de la orden'."
                ),
            )
            return

        self._guardar_config_actual()
        self.destroy()


def iniciar_interfaz() -> None:
    ventana = VentanaPrincipal()
    ventana.mainloop()
