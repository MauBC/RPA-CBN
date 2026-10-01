from __future__ import annotations

from pathlib import Path
from typing import Any
import os
import uuid
import zipfile

from src.excel.file_access import (
    EstadoAccesoExcel,
    diagnosticar_acceso_excel,
)


class ExcelPersistenciaError(RuntimeError):
    """Error base al persistir cambios sobre un Excel."""


class ExcelArchivoOcupadoError(ExcelPersistenciaError):
    """
    El archivo está temporalmente ocupado por Excel u otro proceso.

    Este error es recuperable y más adelante podrá enviarse a una
    cola de sincronización.
    """

    def __init__(
        self,
        ruta: str | Path,
        detalle: str = "",
        winerror: int | None = None,
    ) -> None:
        self.ruta = Path(ruta)
        self.winerror = winerror

        mensaje = (
            f"El archivo Excel está ocupado: {self.ruta}"
        )

        if detalle:
            mensaje += f". {detalle}"

        super().__init__(mensaje)


class ExcelAccesoError(ExcelPersistenciaError):
    """Error no clasificado como bloqueo temporal."""


def exigir_excel_disponible(
    ruta_excel: str | Path,
) -> None:
    """
    Comprueba que Windows permita escribir el archivo.

    Es un pre-check. La operación final sigue validándose durante
    os.replace(), porque el archivo podría bloquearse entre ambos
    momentos.
    """
    diagnostico = diagnosticar_acceso_excel(
        ruta_excel
    )

    if (
        diagnostico.estado
        == EstadoAccesoExcel.DISPONIBLE
    ):
        return

    if (
        diagnostico.estado
        == EstadoAccesoExcel.OCUPADO
    ):
        raise ExcelArchivoOcupadoError(
            diagnostico.ruta,
            diagnostico.detalle,
            diagnostico.winerror,
        )

    raise ExcelAccesoError(
        f"No se puede escribir el archivo "
        f"'{diagnostico.ruta}'. "
        f"Estado={diagnostico.estado.value}. "
        f"{diagnostico.detalle}"
    )


def _ruta_temporal_excel(
    ruta_excel: Path,
) -> Path:
    """
    El temporal se crea en el MISMO directorio que el original.

    Esto permite que os.replace() ocurra dentro del mismo volumen.
    """
    return ruta_excel.with_name(
        f".{ruta_excel.stem}."
        f"rpa-{os.getpid()}-"
        f"{uuid.uuid4().hex}"
        f"{ruta_excel.suffix}"
    )


def _validar_paquete_excel(
    ruta: Path,
) -> None:
    """
    XLSX/XLSM son paquetes ZIP.

    Verificamos que el archivo temporal generado sea estructuralmente
    legible antes de sustituir el original.
    """
    suffix = ruta.suffix.lower()

    if suffix not in {
        ".xlsx",
        ".xlsm",
    }:
        return

    try:
        with zipfile.ZipFile(
            ruta,
            "r",
        ) as archivo_zip:
            archivo_corrupto = (
                archivo_zip.testzip()
            )

            if archivo_corrupto is not None:
                raise ExcelPersistenciaError(
                    f"El Excel temporal contiene "
                    f"un elemento corrupto: "
                    f"{archivo_corrupto}"
                )

    except zipfile.BadZipFile as exc:
        raise ExcelPersistenciaError(
            f"El archivo temporal no es "
            f"un Excel válido: {ruta}"
        ) from exc


def guardar_workbook_atomico(
    workbook: Any,
    ruta_excel: str | Path,
) -> None:
    """
    Persiste un Workbook sin escribir directamente sobre el original.

    Secuencia:
    1. Comprueba disponibilidad.
    2. Guarda a un temporal hermano.
    3. Valida el paquete temporal.
    4. Fuerza flush al disco.
    5. Reemplaza el original con os.replace().
    6. Limpia cualquier temporal restante.

    El Workbook debe cerrarse por el caller.
    """
    ruta_excel = Path(
        ruta_excel
    ).resolve()

    exigir_excel_disponible(
        ruta_excel
    )

    temporal = _ruta_temporal_excel(
        ruta_excel
    )

    try:
        workbook.save(
            temporal
        )

        _validar_paquete_excel(
            temporal
        )

        # Pedimos al SO que vacíe el contenido del temporal
        # antes de reemplazar el original.
        with temporal.open("r+b") as archivo:
            os.fsync(
                archivo.fileno()
            )

        try:
            os.replace(
                temporal,
                ruta_excel,
            )

        except OSError as exc:
            winerror = getattr(
                exc,
                "winerror",
                None,
            )

            if winerror in (
                32,
                33,
            ):
                raise ExcelArchivoOcupadoError(
                    ruta_excel,
                    "El archivo fue bloqueado antes "
                    "del reemplazo final.",
                    winerror,
                ) from exc

            raise ExcelAccesoError(
                f"No se pudo reemplazar "
                f"'{ruta_excel}'. "
                f"Detalle: {exc}"
            ) from exc

    finally:
        try:
            temporal.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass
