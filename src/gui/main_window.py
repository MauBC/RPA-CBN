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
    TipoIncidenciaSincronizacion,
    obtener_estado_sincronizacion_excel,
)
from src.excel.pending_sync import (
    preparar_reintento_manual,
    ignorar_ordenes_manual,
    reabrir_ordenes_manual,
    sincronizar_actualizaciones_pendientes,
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



def _texto_accion_reintento(
    estado: Any,
) -> str:
    texto = str(
        estado
        or ""
    ).strip().casefold()

    if texto == "error":
        return "Reintentar"

    if texto == "ignorada":
        return "Reabrir"

    return ""

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


def _ids_incidencias_sincronizacion(
    resultado: ResumenSincronizacionExcel,
    tipo: TipoIncidenciaSincronizacion,
    *,
    limite: int = 4,
) -> str:
    """
    Devuelve una lista compacta de ID_ORDEN para la GUI.

    Evita convertir el indicador de sincronización en un
    visor de logs cuando existen muchas incidencias.
    """
    ids = [
        incidencia.id_orden
        for incidencia in resultado.incidencias
        if (
            incidencia.tipo == tipo
            and incidencia.id_orden
        )
    ]

    if not ids:
        return ""

    visibles = ids[
        :limite
    ]

    texto = ", ".join(
        visibles
    )

    restantes = (
        len(ids)
        - len(visibles)
    )

    if restantes > 0:
        texto += (
            f" (+{restantes} más)"
        )

    return texto


def _texto_estado_sincronizacion_gui(
    resultado: ResumenSincronizacionExcel,
) -> str:
    """
    Construye el texto orientado al operador.

    La primera línea conserva el resumen general existente.
    Las líneas adicionales aparecen solo cuando aportan una
    acción clara y usan el contrato estructurado de CP13A.1.
    """
    lineas = [
        f"Excel: {resultado.mensaje}"
    ]

    ids_activos = (
        _ids_incidencias_sincronizacion(
            resultado,
            TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO,
        )
    )

    ids_inflight = (
        _ids_incidencias_sincronizacion(
            resultado,
            TipoIncidenciaSincronizacion.INFLIGHT,
        )
    )

    ids_fallidas = (
        _ids_incidencias_sincronizacion(
            resultado,
            TipoIncidenciaSincronizacion.FALLIDA,
        )
    )

    ids_pendientes = (
        _ids_incidencias_sincronizacion(
            resultado,
            TipoIncidenciaSincronizacion.PENDIENTE,
        )
    )

    if ids_activos:
        lineas.append(
            "Procesando ahora: "
            f"{ids_activos}. "
            "No requiere intervención."
        )

    if ids_inflight:
        lineas.append(
            "Ejecución incierta: "
            f"{ids_inflight}. "
            "No reprocesar automáticamente."
        )

    if ids_fallidas:
        lineas.append(
            "Revisión manual: "
            f"{ids_fallidas}. "
            "No reprocesar automáticamente estas órdenes."
        )

    if (
        ids_pendientes
        and resultado.estado
        in {
            EstadoSincronizacionExcel.PENDIENTE,
            EstadoSincronizacionExcel.EXCEL_OCUPADO,
            EstadoSincronizacionExcel.EN_EJECUCION,
            EstadoSincronizacionExcel.REQUIERE_REVISION,
        }
    ):
        lineas.append(
            "Pendientes de sincronizar: "
            f"{ids_pendientes}."
        )

    return "\n".join(
        lineas
    )


def _texto_tipo_incidencia_gui(
    tipo: TipoIncidenciaSincronizacion,
) -> str:
    if tipo == TipoIncidenciaSincronizacion.PENDIENTE:
        return "Pendiente de sincronización"

    if tipo == TipoIncidenciaSincronizacion.FALLIDA:
        return "Requiere revisión"

    if tipo == TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO:
        return "Procesándose ahora"

    if tipo == TipoIncidenciaSincronizacion.INFLIGHT:
        return "Ejecución incierta"

    return str(tipo)


def _accion_incidencia_gui(
    tipo: TipoIncidenciaSincronizacion,
) -> str:
    if tipo == TipoIncidenciaSincronizacion.PENDIENTE:
        return (
            "Use 'Sincronizar Excel'. "
            "No vuelva a ejecutar esta orden en CBN."
        )

    if tipo == TipoIncidenciaSincronizacion.FALLIDA:
        return (
            "Revise el motivo antes de realizar cualquier acción. "
            "No reprocesar automáticamente."
        )

    if tipo == TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO:
        return (
            "No se requiere intervención. "
            "Espere a que termine la orden actual."
        )

    if tipo == TipoIncidenciaSincronizacion.INFLIGHT:
        return (
            "Confirme primero el resultado de la ejecución anterior. "
            "No reprocesar automáticamente."
        )

    return "Revise la incidencia antes de continuar."


def _formatear_incidencias_gui(
    resultado: ResumenSincronizacionExcel,
) -> str:
    if not resultado.incidencias:
        return (
            "No existen incidencias registradas "
            "para este archivo Excel."
        )

    bloques = []

    for incidencia in resultado.incidencias:
        lineas = [
            f"Orden: {incidencia.id_orden or '(sin ID)'}",
            (
                "Estado: "
                + _texto_tipo_incidencia_gui(
                    incidencia.tipo
                )
            ),
        ]

        if incidencia.codigo:
            lineas.append(
                f"Código: {incidencia.codigo}"
            )

        if incidencia.intentos:
            lineas.append(
                f"Intentos: {incidencia.intentos}"
            )

        if incidencia.fecha:
            lineas.append(
                f"Última actualización: {incidencia.fecha}"
            )

        if incidencia.motivo:
            lineas.append(
                f"Motivo: {incidencia.motivo}"
            )

        lineas.append(
            "Acción recomendada: "
            + _accion_incidencia_gui(
                incidencia.tipo
            )
        )

        bloques.append(
            "\n".join(
                lineas
            )
        )

    return (
        "\n\n"
        + ("-" * 58)
        + "\n\n"
    ).join(
        bloques
    )


def _puede_ver_incidencias(
    resultado: ResumenSincronizacionExcel,
    *,
    ejecutando: bool = False,
) -> bool:
    return (
        bool(resultado.incidencias)
        and not ejecutando
    )


def _puede_sincronizar_manualmente(
    resultado: ResumenSincronizacionExcel,
    *,
    ejecutando: bool = False,
    sincronizando: bool = False,
) -> bool:
    """
    Indica si la GUI puede intentar journal -> Excel.

    Esta acción nunca vuelve a ejecutar una orden en CBN.
    """
    return (
        resultado.pendientes > 0
        and resultado.puede_escribir
        and not ejecutando
        and not sincronizando
    )


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
        EstadoSincronizacionExcel.EN_EJECUCION,
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
        self._reintento_en_progreso = False
        self._directorio_actual = ""
        self._resultado_actual = ""
        self._filas_tabla: dict[str, str] = {}
        self._datos_financieros: dict[str, tuple[str, str]] = {}

        # CP9 - Estado independiente de sincronización Excel.
        # La consulta se ejecuta fuera del hilo de Tkinter para que
        # un lock temporal del journal nunca congele la interfaz.
        self._consulta_sync_activa = False
        self._intervalo_sync_ms = 2500
        self._sincronizacion_en_progreso = False
        self._ultimo_estado_sync: (
            ResumenSincronizacionExcel
            | None
        ) = None

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
            padx=(8, 8),
            pady=(0, 10),
            sticky="ew",
        )

        self.boton_ver_incidencias = ctk.CTkButton(
            configuracion,
            text="Ver incidencias",
            width=125,
            state="disabled",
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            command=self._mostrar_panel_incidencias,
        )

        self.boton_ver_incidencias.grid(
            row=1,
            column=5,
            padx=(0, 16),
            pady=(0, 10),
        )

        acciones = ctk.CTkFrame(self, fg_color="transparent")
        acciones.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        acciones.grid_columnconfigure(6, weight=1)

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

        self.boton_sincronizar_excel = ctk.CTkButton(
            acciones,
            text="Sincronizar Excel",
            state="disabled",
            command=self._sincronizar_excel_manual,
        )

        self.boton_sincronizar_excel.grid(
            row=0,
            column=2,
            padx=8,
        )

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
        self.boton_detener.grid(row=0, column=3, padx=8)

        self.boton_sesion = ctk.CTkButton(
            acciones,
            text="Restablecer sesión",
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            command=self._restablecer_sesion,
        )
        self.boton_sesion.grid(row=0, column=4, padx=8)

        self.boton_carpeta = ctk.CTkButton(
            acciones,
            text="Abrir ejecución",
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            state="disabled",
            command=self._abrir_ejecucion,
        )
        self.boton_carpeta.grid(row=0, column=5, padx=8)

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

        columnas = (
            "id",
            "importe",
            "moneda",
            "estado",
            "etapa",
            "detalle",
            "accion",
        )
        self.tabla = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings",
            selectmode="extended",
            height=8,
        )
        self.tabla.heading("id", text="ID orden")
        self.tabla.heading("importe", text="Importe")
        self.tabla.heading("moneda", text="Moneda")
        self.tabla.heading("estado", text="Estado")
        self.tabla.heading("etapa", text="Etapa")
        self.tabla.heading("detalle", text="Detalle")
        self.tabla.heading("accion", text="Acción")
        self.tabla.column("id", width=85, anchor="center")
        self.tabla.column("importe", width=125, anchor="e")
        self.tabla.column("moneda", width=80, anchor="center")
        self.tabla.column("estado", width=110, anchor="center")
        self.tabla.column("etapa", width=150)
        self.tabla.column("detalle", width=355)
        self.tabla.column(
            "accion",
            width=105,
            anchor="center",
            stretch=False,
        )

        scroll_tabla = ttk.Scrollbar(
            tabla_frame,
            orient="vertical",
            command=self.tabla.yview,
        )
        self.tabla.configure(yscrollcommand=scroll_tabla.set)

        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll_tabla.grid(row=0, column=1, sticky="ns")

        # ttk.Treeview no permite widgets reales dentro de cada
        # celda. La columna Acción funciona como un botón por fila:
        # cuando contiene "Reintentar", un clic ejecuta la acción.
        self.tabla.bind(
            "<ButtonRelease-1>",
            self._manejar_click_tabla,
            add="+",
        )

        acciones_tabla = ctk.CTkFrame(
            tabla_frame,
            fg_color="transparent",
        )
        acciones_tabla.grid(
            row=1,
            column=0,
            columnspan=2,
            pady=(8, 0),
            sticky="ew",
        )

        self.boton_ignorar_seleccion = ctk.CTkButton(
            acciones_tabla,
            text="Ignorar seleccionadas",
            width=170,
            state="disabled",
            command=self._ignorar_seleccionadas,
        )
        self.boton_ignorar_seleccion.grid(
            row=0,
            column=0,
            padx=(0, 8),
            sticky="w",
        )

        self.boton_reabrir_seleccion = ctk.CTkButton(
            acciones_tabla,
            text="Reabrir seleccionadas",
            width=170,
            state="disabled",
            command=self._reabrir_seleccionadas,
        )
        self.boton_reabrir_seleccion.grid(
            row=0,
            column=1,
            sticky="w",
        )

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

        self._ultimo_estado_sync = None

        self.boton_sincronizar_excel.configure(
            state="disabled"
        )

        self.boton_ver_incidencias.configure(
            state="disabled"
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
            self._ultimo_estado_sync = None

            self.boton_sincronizar_excel.configure(
                state="disabled"
            )

            self.boton_ver_incidencias.configure(
                state="disabled"
            )

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

        pid_ejecucion_activa = (
            os.getpid()
            if self._ejecutando
            else None
        )

        def trabajo() -> None:
            try:
                resultado = (
                    obtener_estado_sincronizacion_excel(
                        ruta_texto,
                        pid_ejecucion_activa=(
                            pid_ejecucion_activa
                        ),
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
        self._ultimo_estado_sync = resultado

        color = _color_estado_sincronizacion(
            resultado.estado
        )

        self.etiqueta_sync_excel.configure(
            text=(
                _texto_estado_sincronizacion_gui(
                    resultado
                )
            ),
            text_color=color,
        )

        puede_sincronizar = (
            _puede_sincronizar_manualmente(
                resultado,
                ejecutando=self._ejecutando,
                sincronizando=(
                    self._sincronizacion_en_progreso
                ),
            )
        )

        self.boton_sincronizar_excel.configure(
            state=(
                "normal"
                if puede_sincronizar
                else "disabled"
            )
        )

        puede_ver = (
            _puede_ver_incidencias(
                resultado,
                ejecutando=self._ejecutando,
            )
        )

        self.boton_ver_incidencias.configure(
            state=(
                "normal"
                if puede_ver
                else "disabled"
            )
        )

    def _mostrar_panel_incidencias(
        self,
    ) -> None:
        """
        Muestra información operacional del journal sin modificar
        el Excel ni ejecutar ninguna acción en CBN.
        """
        if self._ejecutando:
            return

        resultado = self._ultimo_estado_sync

        if (
            resultado is None
            or not resultado.incidencias
        ):
            messagebox.showinfo(
                "Incidencias Excel",
                (
                    "No existen incidencias registradas "
                    "para el archivo seleccionado."
                ),
            )
            return

        ventana = ctk.CTkToplevel(
            self
        )

        ventana.title(
            "Incidencias de sincronización Excel"
        )

        ventana.geometry(
            "760x520"
        )

        ventana.minsize(
            650,
            420,
        )

        ventana.transient(
            self
        )

        ventana.grid_columnconfigure(
            0,
            weight=1,
        )

        ventana.grid_rowconfigure(
            2,
            weight=1,
        )

        ctk.CTkLabel(
            ventana,
            text="Incidencias de sincronización",
            font=ctk.CTkFont(
                size=20,
                weight="bold",
            ),
            anchor="w",
        ).grid(
            row=0,
            column=0,
            padx=20,
            pady=(20, 4),
            sticky="ew",
        )

        ctk.CTkLabel(
            ventana,
            text=(
                "Estas incidencias protegen órdenes que no "
                "deben reprocesarse automáticamente."
            ),
            anchor="w",
            text_color=(
                "gray35",
                "gray70",
            ),
        ).grid(
            row=1,
            column=0,
            padx=20,
            pady=(0, 12),
            sticky="ew",
        )

        texto = ctk.CTkTextbox(
            ventana,
            wrap="word",
        )

        texto.grid(
            row=2,
            column=0,
            padx=20,
            pady=(0, 14),
            sticky="nsew",
        )

        texto.insert(
            "1.0",
            _formatear_incidencias_gui(
                resultado
            ),
        )

        texto.configure(
            state="disabled"
        )

        ctk.CTkButton(
            ventana,
            text="Cerrar",
            width=100,
            command=ventana.destroy,
        ).grid(
            row=3,
            column=0,
            padx=20,
            pady=(0, 20),
            sticky="e",
        )

        ventana.grab_set()

    def _sincronizar_excel_manual(
        self,
    ) -> None:
        """
        Reintenta únicamente journal -> Excel.

        No abre navegador y no vuelve a ejecutar CBN.
        """
        if self._ejecutando:
            return

        if self._sincronizacion_en_progreso:
            return

        ruta = self._obtener_excel()

        if ruta is None:
            return

        estado = self._ultimo_estado_sync

        if (
            estado is None
            or estado.pendientes <= 0
        ):
            messagebox.showinfo(
                "Sincronización Excel",
                (
                    "No existen resultados pendientes "
                    "de sincronizar."
                ),
            )
            return

        if not estado.puede_escribir:
            messagebox.showwarning(
                "Excel no disponible",
                (
                    "El archivo Excel no está disponible "
                    "para escritura.\n\n"
                    "Cierre Excel o espere a que OneDrive "
                    "libere el archivo y vuelva a intentarlo."
                ),
            )
            return

        self._sincronizacion_en_progreso = True

        self.boton_sincronizar_excel.configure(
            state="disabled"
        )

        self.etiqueta_estado.configure(
            text=(
                "Sincronizando resultados pendientes "
                "con el archivo Excel..."
            )
        )

        def trabajo() -> None:
            try:
                resultado = (
                    sincronizar_actualizaciones_pendientes(
                        ruta
                    )
                )

                self._cola.put(
                    {
                        "type": "manual_sync_finished",
                        "result": resultado,
                    }
                )

            except Exception as error:
                self._cola.put(
                    {
                        "type": "manual_sync_error",
                        "message": str(error),
                    }
                )

        Thread(
            target=trabajo,
            daemon=True,
        ).start()

    def _manejar_click_tabla(
        self,
        evento,
    ) -> None:
        """
        Ejecuta la accion mostrada en la columna Accion.
        """
        if self._ejecutando:
            return

        if self._reintento_en_progreso:
            return

        region = self.tabla.identify(
            "region",
            evento.x,
            evento.y,
        )

        if region != "cell":
            return

        columna = self.tabla.identify_column(
            evento.x
        )

        if columna != "#7":
            return

        iid = self.tabla.identify_row(
            evento.y
        )

        if not iid:
            return

        valores = self.tabla.item(
            iid,
            "values",
        )

        if (
            not valores
            or len(valores) < 7
        ):
            return

        accion = str(
            valores[6]
            or ""
        ).strip()

        id_orden = str(
            valores[0]
            or ""
        ).strip()

        if not id_orden:
            return

        self.tabla.selection_set(
            iid
        )

        self.tabla.focus(
            iid
        )

        if accion == "Reintentar":
            self._preparar_reintento_orden(
                id_orden,
                etapa=str(
                    valores[4]
                    or ""
                ),
                detalle=str(
                    valores[5]
                    or ""
                ),
            )

            return

        if accion == "Reabrir":
            self._cambiar_estado_manual(
                [
                    id_orden
                ],
                accion="reabrir",
            )

    def _ids_seleccionados_por_estado(
        self,
        estado_objetivo: str,
    ) -> list[str]:
        objetivo = str(
            estado_objetivo
            or ""
        ).strip().casefold()

        ids: list[str] = []

        for iid in self.tabla.selection():
            valores = self.tabla.item(
                iid,
                "values",
            )

            if (
                not valores
                or len(valores) < 4
            ):
                continue

            estado = str(
                valores[3]
                or ""
            ).strip().casefold()

            if estado != objetivo:
                continue

            id_orden = str(
                valores[0]
                or ""
            ).strip()

            if (
                id_orden
                and id_orden not in ids
            ):
                ids.append(
                    id_orden
                )

        return ids

    def _ignorar_seleccionadas(
        self,
    ) -> None:
        ids = self._ids_seleccionados_por_estado(
            "Error"
        )

        if not ids:
            messagebox.showwarning(
                "Sin errores seleccionados",
                (
                    "Seleccione una o más órdenes "
                    "con estado Error."
                ),
            )
            return

        self._cambiar_estado_manual(
            ids,
            accion="ignorar",
        )

    def _reabrir_seleccionadas(
        self,
    ) -> None:
        ids = self._ids_seleccionados_por_estado(
            "Ignorada"
        )

        if not ids:
            messagebox.showwarning(
                "Sin ignoradas seleccionadas",
                (
                    "Seleccione una o más órdenes "
                    "con estado Ignorada."
                ),
            )
            return

        self._cambiar_estado_manual(
            ids,
            accion="reabrir",
        )

    def _cambiar_estado_manual(
        self,
        ids_orden: list[str],
        *,
        accion: str,
    ) -> None:
        if self._ejecutando:
            messagebox.showwarning(
                "RPA en ejecución",
                (
                    "Espere a que termine la ejecución "
                    "actual antes de modificar órdenes."
                ),
            )
            return

        if self._reintento_en_progreso:
            return

        ruta = self._obtener_excel()

        if ruta is None:
            return

        ids = [
            str(
                item
                or ""
            ).strip()
            for item in ids_orden
            if str(
                item
                or ""
            ).strip()
        ]

        if not ids:
            return

        if accion == "ignorar":
            titulo = "Ignorar órdenes"

            descripcion = (
                f"Se marcarán {len(ids)} orden(es) "
                "como Ignorada.\n\n"
                "ESTADO_RPA cambiará de 2 a 3.\n"
                "Estas órdenes ya no se procesarán "
                "automáticamente.\n\n"
                "¿Desea continuar?"
            )

            operacion = ignorar_ordenes_manual

        elif accion == "reabrir":
            titulo = "Reabrir órdenes"

            descripcion = (
                f"Se reabrirán {len(ids)} orden(es).\n\n"
                "ESTADO_RPA cambiará de 3 a 2.\n"
                "Volverán a aparecer como errores "
                "disponibles para reintento.\n\n"
                "¿Desea continuar?"
            )

            operacion = reabrir_ordenes_manual

        else:
            raise ValueError(
                f"Acción no soportada: {accion}"
            )

        confirmar = messagebox.askyesno(
            titulo,
            descripcion,
        )

        if not confirmar:
            return

        self._reintento_en_progreso = True

        self._cambiar_controles(
            False
        )

        self.etiqueta_estado.configure(
            text=(
                f"Actualizando {len(ids)} "
                "orden(es)..."
            )
        )

        def trabajo() -> None:
            try:
                resultado = operacion(
                    ruta,
                    ids,
                )

                self._cola.put(
                    {
                        "type": (
                            "manual_state_change_ready"
                        ),
                        "action": accion,
                        "ids": ids,
                        "rows": resultado,
                    }
                )

            except Exception as error:
                self._cola.put(
                    {
                        "type": (
                            "manual_state_change_error"
                        ),
                        "action": accion,
                        "ids": ids,
                        "message": str(error),
                    }
                )

        Thread(
            target=trabajo,
            daemon=True,
        ).start()

    def _preparar_reintento_orden(
        self,
        id_orden: str,
        *,
        etapa: str = "",
        detalle: str = "",
    ) -> None:
        if self._ejecutando:
            messagebox.showwarning(
                "RPA en ejecución",
                (
                    "Espere a que termine la ejecución "
                    "actual antes de reintentar una orden."
                ),
            )
            return

        if self._reintento_en_progreso:
            return

        ruta = self._obtener_excel()

        if ruta is None:
            return

        texto_confirmacion = (
            f"Se preparará para reintento la orden "
            f"{id_orden}.\n\n"
            "Esto cambiará ESTADO_RPA de 2 a 0 en "
            "todas las posiciones de ese ID.\n\n"
        )

        if etapa:
            texto_confirmacion += (
                f"Etapa del error: {etapa}\n"
            )

        if detalle:
            texto_confirmacion += (
                f"Detalle: {detalle}\n\n"
            )

        texto_confirmacion += (
            "Importante: si el portal pudo haber completado "
            "parcialmente la operación, revise primero las "
            "evidencias de la ejecución.\n\n"
            "¿Desea continuar?"
        )

        confirmar = messagebox.askyesno(
            "Reintentar orden",
            texto_confirmacion,
        )

        if not confirmar:
            return

        self._reintento_en_progreso = True
        self._cambiar_controles(
            False
        )

        self.etiqueta_estado.configure(
            text=(
                f"Preparando reintento de "
                f"ID_ORDEN={id_orden}..."
            )
        )

        def trabajo() -> None:
            try:
                filas = (
                    preparar_reintento_manual(
                        ruta,
                        id_orden,
                    )
                )

                self._cola.put(
                    {
                        "type": (
                            "manual_retry_ready"
                        ),
                        "id_orden": id_orden,
                        "rows": filas,
                    }
                )

            except Exception as error:
                self._cola.put(
                    {
                        "type": (
                            "manual_retry_error"
                        ),
                        "id_orden": id_orden,
                        "message": str(error),
                    }
                )

        Thread(
            target=trabajo,
            daemon=True,
        ).start()

    def _iniciar_reintento_id(
        self,
        id_orden: str,
    ) -> None:
        """
        Ejecuta nuevamente únicamente el ID que acaba
        de ser restablecido de ESTADO_RPA=2 a 0.
        """
        if self._ejecutando:
            return

        ruta = self._obtener_excel()

        if ruta is None:
            return

        navegador = NAVEGADORES_UI.get(
            self.selector_navegador.get(),
            "msedge",
        )

        self._ejecutando = True
        self._reintento_en_progreso = False
        self._cancelar.clear()
        self._directorio_actual = ""
        self._resultado_actual = ""
        self.barra_progreso.set(
            0
        )

        self._cambiar_controles(
            False,
            ejecutando=True,
        )

        self._actualizar_fila(
            id_orden,
            "Pendiente",
            "Reintento",
            (
                "Reintento preparado. "
                "Iniciando únicamente esta orden..."
            ),
        )

        self.etiqueta_estado.configure(
            text=(
                f"Reintentando únicamente "
                f"ID_ORDEN={id_orden}..."
            )
        )

        self._agregar_log(
            ""
        )

        self._agregar_log(
            (
                "=============================="
                "=============================="
            )
        )

        self._agregar_log(
            (
                "REINTENTO MANUAL "
                f"ID_ORDEN={id_orden}"
            )
        )

        self._agregar_log(
            (
                "=============================="
                "=============================="
            )
        )

        def callback(
            evento: dict[str, Any]
        ) -> None:
            self._cola.put(
                evento
            )

        def trabajo() -> None:
            try:
                ejecutar_rpa(
                    ruta_excel=ruta,
                    navegador=navegador,
                    callback=callback,
                    cancelar_evento=(
                        self._cancelar
                    ),
                    solo_id_orden=id_orden,
                )

            except Exception:
                self._cola.put(
                    {
                        "type": "fatal_error",
                        "message": (
                            "La aplicación no pudo "
                            "completar el reintento. "
                            "Revise las evidencias."
                        ),
                        "stage": "interfaz",
                    }
                )

                self._cola.put(
                    {
                        "type": (
                            "run_finished_unexpected"
                        ),
                    }
                )

        self._hilo = Thread(
            target=trabajo,
            daemon=True,
        )

        self._hilo.start()

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

        if tipo == "manual_sync_finished":
            self._sincronizacion_en_progreso = False

            resultado = dict(
                evento.get(
                    "result",
                    {},
                )
                or {}
            )

            aplicadas = int(
                resultado.get(
                    "aplicadas",
                    0,
                )
                or 0
            )

            restantes = int(
                resultado.get(
                    "restantes",
                    0,
                )
                or 0
            )

            aisladas = int(
                resultado.get(
                    "aisladas",
                    0,
                )
                or 0
            )

            self._consultar_estado_sincronizacion_excel()

            if (
                restantes == 0
                and aisladas == 0
            ):
                self.etiqueta_estado.configure(
                    text=(
                        "Sincronización Excel completada."
                    )
                )

                messagebox.showinfo(
                    "Sincronización Excel",
                    (
                        "La sincronización terminó "
                        "correctamente.\n\n"
                        f"Resultados aplicados: {aplicadas}."
                    ),
                )

            else:
                self.etiqueta_estado.configure(
                    text=(
                        "La sincronización Excel "
                        "todavía requiere atención."
                    )
                )

                messagebox.showwarning(
                    "Sincronización incompleta",
                    (
                        f"Resultados aplicados: {aplicadas}\n"
                        f"Pendientes restantes: {restantes}\n"
                        f"Requieren revisión: {aisladas}\n\n"
                        "No vuelva a ejecutar las órdenes "
                        "afectadas en CBN."
                    ),
                )

            return

        if tipo == "manual_sync_error":
            self._sincronizacion_en_progreso = False

            mensaje = str(
                evento.get(
                    "message",
                    "No se pudo sincronizar el Excel.",
                )
            )

            self.etiqueta_estado.configure(
                text=(
                    "No se pudo completar "
                    "la sincronización Excel."
                )
            )

            self._consultar_estado_sincronizacion_excel()

            messagebox.showerror(
                "Error de sincronización",
                (
                    "No se volvió a ejecutar ninguna "
                    "orden en CBN.\n\n"
                    f"Detalle: {mensaje}"
                ),
            )

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

            iid = self._filas_tabla.get(
                id_orden
            )

            if (
                iid
                and self.tabla.exists(
                    iid
                )
            ):
                self.tabla.selection_set(
                    iid
                )

                self.tabla.focus(
                    iid
                )

                self.tabla.see(
                    iid
                )

            return

        if tipo == "manual_retry_ready":
            self._reintento_en_progreso = False

            id_orden = str(
                evento.get(
                    "id_orden",
                    "",
                )
            )

            filas = int(
                evento.get(
                    "rows",
                    0,
                )
                or 0
            )

            self._actualizar_fila(
                id_orden,
                "Pendiente",
                "Reintento preparado",
                (
                    "ESTADO_RPA restablecido "
                    f"a 0 en {filas} fila(s). "
                    "Iniciando reejecución..."
                ),
            )

            self._consultar_estado_sincronizacion_excel()

            self._iniciar_reintento_id(
                id_orden
            )

            return

        if tipo == "manual_retry_error":
            self._reintento_en_progreso = False

            self._cambiar_controles(
                True
            )

            mensaje = str(
                evento.get(
                    "message",
                    "No se pudo preparar el reintento.",
                )
            )

            self.etiqueta_estado.configure(
                text=(
                    "No se pudo preparar "
                    "el reintento."
                )
            )

            self._consultar_estado_sincronizacion_excel()

            messagebox.showerror(
                "Reintento no permitido",
                mensaje,
            )

            return

        if tipo == "manual_state_change_ready":
            self._reintento_en_progreso = False

            accion = str(
                evento.get(
                    "action",
                    "",
                )
            )

            ids = [
                str(item)
                for item in evento.get(
                    "ids",
                    [],
                )
            ]

            if accion == "ignorar":
                for id_orden in ids:
                    self._actualizar_fila(
                        id_orden,
                        "Ignorada",
                        "Cerrada manualmente",
                        (
                            "ESTADO_RPA=3. "
                            "No se procesará automáticamente."
                        ),
                    )

                mensaje = (
                    f"{len(ids)} orden(es) "
                    "marcada(s) como Ignorada."
                )

            else:
                for id_orden in ids:
                    self._actualizar_fila(
                        id_orden,
                        "Error",
                        "Disponible para reintento",
                        (
                            "ESTADO_RPA=2. "
                            "La orden puede reintentarse."
                        ),
                    )

                mensaje = (
                    f"{len(ids)} orden(es) "
                    "reabierta(s)."
                )

            self._cambiar_controles(
                True
            )

            self.etiqueta_estado.configure(
                text=mensaje
            )

            self._consultar_estado_sincronizacion_excel()

            messagebox.showinfo(
                "Órdenes actualizadas",
                mensaje,
            )

            return

        if tipo == "manual_state_change_error":
            self._reintento_en_progreso = False

            self._cambiar_controles(
                True
            )

            mensaje = str(
                evento.get(
                    "message",
                    "No se pudo actualizar el estado.",
                )
            )

            self.etiqueta_estado.configure(
                text=(
                    "No se pudo modificar "
                    "el estado de las órdenes."
                )
            )

            self._consultar_estado_sincronizacion_excel()

            messagebox.showerror(
                "Cambio no permitido",
                mensaje,
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

        for orden_error in resultado.get(
            "error_orders",
            [],
        ):
            id_orden = str(
                orden_error.get(
                    "id_orden",
                    "",
                )
                or ""
            ).strip()

            if not id_orden:
                continue

            bloqueada = bool(
                orden_error.get(
                    "blocked",
                    False,
                )
            )

            codigo = str(
                orden_error.get(
                    "code",
                    "",
                )
                or ""
            ).strip()

            motivo = str(
                orden_error.get(
                    "reason",
                    "",
                )
                or ""
            ).strip()

            if bloqueada:
                estado_fila = "Revisión"
                etapa_fila = (
                    codigo
                    or "Protección activa"
                )

            else:
                estado_fila = "Error"
                etapa_fila = (
                    "Disponible para reintento"
                )

            self._actualizar_fila(
                id_orden,
                estado_fila,
                etapa_fila,
                motivo,
            )

        for orden_ignorada in resultado.get(
            "ignored_orders",
            [],
        ):
            id_orden = str(
                orden_ignorada.get(
                    "id_orden",
                    "",
                )
                or ""
            ).strip()

            if not id_orden:
                continue

            motivo = str(
                orden_ignorada.get(
                    "reason",
                    (
                        "Orden ignorada manualmente. "
                        "ESTADO_RPA=3."
                    ),
                )
            )

            self._actualizar_fila(
                id_orden,
                "Ignorada",
                "Cerrada manualmente",
                motivo,
            )

        ordenes_resultado = list(
            resultado.get(
                "orders",
                [],
            )
        )

        resumen = _resumen_importes(
            ordenes_resultado
        )
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
            _texto_accion_reintento(
                estado
            ),
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

        self.boton_ignorar_seleccion.configure(
            state=estado
        )

        self.boton_reabrir_seleccion.configure(
            state=estado
        )

        if not habilitar:
            self.boton_sincronizar_excel.configure(
                state="disabled"
            )

        elif self._ultimo_estado_sync is not None:
            puede_sincronizar = (
                _puede_sincronizar_manualmente(
                    self._ultimo_estado_sync,
                    ejecutando=self._ejecutando,
                    sincronizando=(
                        self._sincronizacion_en_progreso
                    ),
                )
            )

            self.boton_sincronizar_excel.configure(
                state=(
                    "normal"
                    if puede_sincronizar
                    else "disabled"
                )
            )

        else:
            self.boton_sincronizar_excel.configure(
                state="disabled"
            )

        if (
            habilitar
            and self._ultimo_estado_sync is not None
            and _puede_ver_incidencias(
                self._ultimo_estado_sync,
                ejecutando=self._ejecutando,
            )
        ):
            self.boton_ver_incidencias.configure(
                state="normal"
            )
        else:
            self.boton_ver_incidencias.configure(
                state="disabled"
            )

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
