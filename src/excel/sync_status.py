from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from src.excel.file_access import (
    EstadoAccesoExcel,
    diagnosticar_acceso_excel,
)
from src.excel.pending_sync import (
    obtener_estado_journal,
)


class EstadoSincronizacionExcel(str, Enum):
    SINCRONIZADO = "SINCRONIZADO"
    EXCEL_OCUPADO = "EXCEL_OCUPADO"
    PENDIENTE = "PENDIENTE"
    EN_EJECUCION = "EN_EJECUCION"
    REQUIERE_REVISION = "REQUIERE_REVISION"
    ERROR_ACCESO = "ERROR_ACCESO"
    NO_EXISTE = "NO_EXISTE"


class TipoIncidenciaSincronizacion(str, Enum):
    PENDIENTE = "PENDIENTE"
    FALLIDA = "FALLIDA"
    INFLIGHT_ACTIVO = "INFLIGHT_ACTIVO"
    INFLIGHT = "INFLIGHT"


@dataclass(frozen=True)
class IncidenciaSincronizacionExcel:
    """
    Información operativa de una orden con estado de
    sincronización pendiente o incierto.

    Este contrato está pensado para presentación/diagnóstico.
    No reemplaza ni modifica el journal persistente.
    """

    tipo: TipoIncidenciaSincronizacion
    id_orden: str
    fecha: str = ""
    intentos: int = 0
    codigo: str = ""
    motivo: str = ""
    reintento_automatico: bool = False


@dataclass(frozen=True)
class ResumenSincronizacionExcel:
    ruta: Path
    estado: EstadoSincronizacionExcel
    mensaje: str
    pendientes: int
    fallidas: int
    puede_escribir: bool
    ocupado: bool
    inflight: int = 0
    winerror: int | None = None
    detalle: str = ""
    incidencias: tuple[
        IncidenciaSincronizacionExcel,
        ...
    ] = ()

    @property
    def sincronizado(self) -> bool:
        return (
            self.estado
            == EstadoSincronizacionExcel.SINCRONIZADO
        )

    @property
    def requiere_atencion(self) -> bool:
        return self.estado in {
            EstadoSincronizacionExcel.REQUIERE_REVISION,
            EstadoSincronizacionExcel.ERROR_ACCESO,
            EstadoSincronizacionExcel.NO_EXISTE,
        }


def _texto_cantidad(
    cantidad: int,
    singular: str,
    plural: str,
) -> str:
    palabra = (
        singular
        if cantidad == 1
        else plural
    )

    return f"{cantidad} {palabra}"


def _entero_seguro(
    valor: object,
) -> int:
    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def _primer_texto(
    item: dict,
    *claves: str,
) -> str:
    for clave in claves:
        valor = item.get(
            clave
        )

        if valor is None:
            continue

        texto = str(
            valor
        ).strip()

        if texto:
            return texto

    return ""


def _es_inflight_activo(
    item: dict,
    pid_ejecucion_activa: int | None,
) -> bool:
    """
    Solo considera activo un inflight cuando el caller confirma
    explícitamente el PID de la ejecución actual.

    Si no existe esa confirmación, se mantiene como incierto.
    """
    if pid_ejecucion_activa is None:
        return False

    pid_item = _entero_seguro(
        item.get(
            "pid",
            0,
        )
    )

    return (
        pid_item > 0
        and pid_item
        == int(
            pid_ejecucion_activa
        )
    )


def _construir_incidencias(
    estado_journal: dict,
    *,
    pid_ejecucion_activa: int | None = None,
) -> tuple[
    IncidenciaSincronizacionExcel,
    ...,
]:
    """
    Convierte la información técnica del journal a un contrato
    estable de observabilidad.

    No modifica ni normaliza el journal original.
    """
    resultado: list[
        IncidenciaSincronizacionExcel
    ] = []

    for item in estado_journal.get(
        "pendientes",
        [],
    ):
        if not isinstance(
            item,
            dict,
        ):
            continue

        motivo = _primer_texto(
            item,
            "ultimo_error",
        )

        if not motivo:
            motivo = (
                "Resultado pendiente de sincronizar "
                "con el archivo Excel."
            )

        resultado.append(
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.PENDIENTE
                ),
                id_orden=_primer_texto(
                    item,
                    "id_orden",
                ),
                fecha=_primer_texto(
                    item,
                    "updated_at",
                    "created_at",
                ),
                intentos=_entero_seguro(
                    item.get(
                        "intentos",
                        0,
                    )
                ),
                codigo=_primer_texto(
                    item,
                    "ultimo_codigo_error",
                ),
                motivo=motivo,
                reintento_automatico=True,
            )
        )

    for item in estado_journal.get(
        "fallidas",
        [],
    ):
        if not isinstance(
            item,
            dict,
        ):
            continue

        motivo = _primer_texto(
            item,
            "sync_error_detail",
            "ultimo_error",
        )

        if not motivo:
            motivo = (
                "La actualización requiere "
                "revisión manual."
            )

        resultado.append(
            IncidenciaSincronizacionExcel(
                tipo=(
                    TipoIncidenciaSincronizacion.FALLIDA
                ),
                id_orden=_primer_texto(
                    item,
                    "id_orden",
                ),
                fecha=_primer_texto(
                    item,
                    "sync_failed_at",
                    "updated_at",
                    "created_at",
                ),
                intentos=_entero_seguro(
                    item.get(
                        "intentos",
                        0,
                    )
                ),
                codigo=_primer_texto(
                    item,
                    "sync_error_code",
                    "ultimo_codigo_error",
                ),
                motivo=motivo,
                reintento_automatico=False,
            )
        )

    for item in estado_journal.get(
        "inflight",
        [],
    ):
        if not isinstance(
            item,
            dict,
        ):
            continue

        activo = _es_inflight_activo(
            item,
            pid_ejecucion_activa,
        )

        if activo:
            tipo = (
                TipoIncidenciaSincronizacion.INFLIGHT_ACTIVO
            )

            codigo = "EJECUCION_ACTIVA"

            motivo = (
                "Esta orden está siendo procesada "
                "actualmente por esta instancia del RPA."
            )

        else:
            tipo = (
                TipoIncidenciaSincronizacion.INFLIGHT
            )

            codigo = "EJECUCION_INCIERTA"

            motivo = (
                "La ejecución anterior no tiene "
                "un resultado final confirmado. "
                "No debe reprocesarse automáticamente."
            )

        resultado.append(
            IncidenciaSincronizacionExcel(
                tipo=tipo,
                id_orden=_primer_texto(
                    item,
                    "id_orden",
                ),
                fecha=_primer_texto(
                    item,
                    "started_at",
                    "updated_at",
                ),
                intentos=0,
                codigo=codigo,
                motivo=motivo,
                reintento_automatico=False,
            )
        )

    return tuple(
        resultado
    )


def obtener_estado_sincronizacion_excel(
    ruta_excel: str | Path,
    *,
    pid_ejecucion_activa: int | None = None,
) -> ResumenSincronizacionExcel:
    """
    Resume el estado técnico de Excel en un contrato simple
    pensado para la interfaz gráfica.

    No modifica el workbook ni intenta sincronizarlo.

    Prioridad visual:
    1. Archivo inexistente / error de acceso.
    2. Operaciones que requieren revisión manual.
    3. Operaciones pendientes.
    4. Excel abierto sin pendientes.
    5. Todo sincronizado.
    """
    ruta = Path(
        ruta_excel
    ).expanduser().resolve()

    try:
        diagnostico = (
            diagnosticar_acceso_excel(
                ruta
            )
        )

    except Exception as error:
        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.ERROR_ACCESO
            ),
            mensaje=(
                "No se pudo comprobar "
                "el estado del Excel."
            ),
            pendientes=0,
            fallidas=0,
            puede_escribir=False,
            ocupado=False,
            winerror=getattr(
                error,
                "winerror",
                None,
            ),
            detalle=str(error),
        )

    try:
        estado_journal = (
            obtener_estado_journal(
                ruta
            )
        )

        pendientes = len(
            estado_journal[
                "pendientes"
            ]
        )

        fallidas = len(
            estado_journal[
                "fallidas"
            ]
        )

        items_inflight = [
            item
            for item
            in estado_journal.get(
                "inflight",
                [],
            )
            if isinstance(
                item,
                dict,
            )
        ]

        inflight = len(
            items_inflight
        )

        inflight_activo = sum(
            1
            for item
            in items_inflight
            if _es_inflight_activo(
                item,
                pid_ejecucion_activa,
            )
        )

        inflight_incierto = (
            inflight
            - inflight_activo
        )

        incidencias = (
            _construir_incidencias(
                estado_journal,
                pid_ejecucion_activa=(
                    pid_ejecucion_activa
                ),
            )
        )

    except Exception as error:
        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.ERROR_ACCESO
            ),
            mensaje=(
                "No se pudo leer el estado "
                "de sincronización."
            ),
            pendientes=0,
            fallidas=0,
            puede_escribir=(
                diagnostico.puede_escribir
            ),
            ocupado=diagnostico.ocupado,
            winerror=diagnostico.winerror,
            detalle=str(error),
        )

    if (
        diagnostico.estado
        == EstadoAccesoExcel.NO_EXISTE
    ):
        mensaje = "El archivo Excel ya no existe."

        if pendientes:
            mensaje += (
                " "
                + _texto_cantidad(
                    pendientes,
                    "cambio pendiente",
                    "cambios pendientes",
                )
                + "."
            )

        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.NO_EXISTE
            ),
            mensaje=mensaje,
            pendientes=pendientes,
            fallidas=fallidas,
            puede_escribir=False,
            inflight=inflight,
            incidencias=incidencias,
            ocupado=False,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    if (
        diagnostico.estado
        == EstadoAccesoExcel.ERROR_ACCESO
    ):
        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.ERROR_ACCESO
            ),
            mensaje=(
                "Excel no disponible para escritura."
            ),
            pendientes=pendientes,
            fallidas=fallidas,
            puede_escribir=False,
            inflight=inflight,
            incidencias=incidencias,
            ocupado=False,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    # Una operación aislada requiere atención aunque además
    # existan cambios recuperables pendientes.
    if fallidas or inflight_incierto:
        partes = []

        if inflight_incierto:
            partes.append(
                _texto_cantidad(
                    inflight_incierto,
                    "orden en ejecuci\u00f3n incierta",
                    "órdenes en ejecuci\u00f3n incierta",
                )
            )

        if fallidas:
            partes.append(
                _texto_cantidad(
                    fallidas,
                    "actualizaci\u00f3n requiere revisi\u00f3n",
                    "actualizaciones requieren revisi\u00f3n",
                )
            )

        if pendientes:
            partes.append(
                _texto_cantidad(
                    pendientes,
                    "cambio pendiente",
                    "cambios pendientes",
                )
            )

        if diagnostico.ocupado:
            partes.append(
                "Excel abierto"
            )

        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.REQUIERE_REVISION
            ),
            mensaje=" · ".join(partes),
            pendientes=pendientes,
            fallidas=fallidas,
            puede_escribir=(
                diagnostico.puede_escribir
            ),
            inflight=inflight,
            incidencias=incidencias,
            ocupado=diagnostico.ocupado,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    if inflight_activo:
        partes = [
            _texto_cantidad(
                inflight_activo,
                "orden procesándose ahora",
                "órdenes procesándose ahora",
            )
        ]

        if pendientes:
            partes.append(
                _texto_cantidad(
                    pendientes,
                    "cambio pendiente",
                    "cambios pendientes",
                )
            )

        if diagnostico.ocupado:
            partes.append(
                "Excel abierto"
            )

        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.EN_EJECUCION
            ),
            mensaje=" · ".join(
                partes
            ),
            pendientes=pendientes,
            fallidas=0,
            puede_escribir=(
                diagnostico.puede_escribir
            ),
            inflight=inflight,
            incidencias=incidencias,
            ocupado=diagnostico.ocupado,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    if pendientes:
        cantidad = _texto_cantidad(
            pendientes,
            "cambio pendiente",
            "cambios pendientes",
        )

        if diagnostico.ocupado:
            mensaje = (
                f"Excel abierto · {cantidad}"
            )

            estado = (
                EstadoSincronizacionExcel.EXCEL_OCUPADO
            )

        else:
            mensaje = (
                f"{cantidad} · "
                "Excel disponible para sincronizar"
            )

            estado = (
                EstadoSincronizacionExcel.PENDIENTE
            )

        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=estado,
            mensaje=mensaje,
            pendientes=pendientes,
            fallidas=0,
            puede_escribir=(
                diagnostico.puede_escribir
            ),
            inflight=0,
            incidencias=incidencias,
            ocupado=diagnostico.ocupado,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    if diagnostico.ocupado:
        return ResumenSincronizacionExcel(
            ruta=ruta,
            estado=(
                EstadoSincronizacionExcel.EXCEL_OCUPADO
            ),
            mensaje=(
                "Excel abierto · "
                "Sin cambios pendientes"
            ),
            pendientes=0,
            fallidas=0,
            puede_escribir=False,
            ocupado=True,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    return ResumenSincronizacionExcel(
        ruta=ruta,
        estado=(
            EstadoSincronizacionExcel.SINCRONIZADO
        ),
        mensaje=(
            "Excel listo · "
            "Sin cambios pendientes"
        ),
        pendientes=0,
        fallidas=0,
        puede_escribir=True,
        ocupado=False,
        winerror=diagnostico.winerror,
        detalle=diagnostico.detalle,
    )
