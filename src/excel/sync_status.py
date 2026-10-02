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
    REQUIERE_REVISION = "REQUIERE_REVISION"
    ERROR_ACCESO = "ERROR_ACCESO"
    NO_EXISTE = "NO_EXISTE"


@dataclass(frozen=True)
class ResumenSincronizacionExcel:
    ruta: Path
    estado: EstadoSincronizacionExcel
    mensaje: str
    pendientes: int
    fallidas: int
    puede_escribir: bool
    ocupado: bool
    winerror: int | None = None
    detalle: str = ""

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


def obtener_estado_sincronizacion_excel(
    ruta_excel: str | Path,
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

        inflight = len(
            estado_journal.get(
                "inflight",
                [],
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
            ocupado=False,
            winerror=diagnostico.winerror,
            detalle=diagnostico.detalle,
        )

    # Una operación aislada requiere atención aunque además
    # existan cambios recuperables pendientes.
    if fallidas or inflight:
        partes = []

        if inflight:
            partes.append(
                _texto_cantidad(
                    inflight,
                    "orden en ejecuci\u00f3n incierta",
                    "ordenes en ejecuci\u00f3n incierta",
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
            mensaje=" ? ".join(partes),
            pendientes=pendientes,
            fallidas=fallidas,
            puede_escribir=(
                diagnostico.puede_escribir
            ),
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
