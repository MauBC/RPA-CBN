from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any
import hashlib
import json
import os
import shutil
import time
import uuid

from src.excel.validators import OrdenCotizacion


class SnapshotError(RuntimeError):
    """Error al crear una copia consistente de un input del RPA."""


def sha256_archivo(ruta: str | Path) -> str:
    ruta = Path(ruta)

    digest = hashlib.sha256()

    with ruta.open("rb") as archivo:
        while True:
            bloque = archivo.read(1024 * 1024)

            if not bloque:
                break

            digest.update(bloque)

    return digest.hexdigest()


def crear_snapshot_estable(
    origen: str | Path,
    destino: str | Path,
    *,
    intentos: int = 3,
    espera_segundos: float = 0.15,
) -> dict[str, Any]:
    """
    Copia un archivo asegurando que el origen no cambió durante la copia.

    Para cada intento:
    1. SHA256 origen antes.
    2. Copia a temporal.
    3. SHA256 copia.
    4. SHA256 origen después.
    5. Solo acepta si los tres hashes coinciden.
    """
    origen = Path(origen).resolve()
    destino = Path(destino).resolve()

    if not origen.exists():
        raise SnapshotError(
            f"No existe el archivo a capturar: {origen}"
        )

    if not origen.is_file():
        raise SnapshotError(
            f"La ruta no es un archivo: {origen}"
        )

    if origen == destino:
        raise SnapshotError(
            "El snapshot no puede sobrescribir el archivo de origen."
        )

    destino.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    ultimo_error: Exception | None = None

    for intento in range(1, intentos + 1):
        temporal = destino.with_name(
            f".{destino.name}."
            f"{os.getpid()}."
            f"{uuid.uuid4().hex}.tmp"
        )

        try:
            hash_antes = sha256_archivo(origen)

            shutil.copy2(
                origen,
                temporal,
            )

            hash_copia = sha256_archivo(temporal)
            hash_despues = sha256_archivo(origen)

            if (
                hash_antes
                == hash_copia
                == hash_despues
            ):
                os.replace(
                    temporal,
                    destino,
                )

                stat = origen.stat()

                return {
                    "origen": str(origen),
                    "snapshot": str(destino),
                    "sha256": hash_copia,
                    "bytes": destino.stat().st_size,
                    "origen_mtime_ns": stat.st_mtime_ns,
                    "capturado_en": datetime.now().isoformat(
                        timespec="seconds"
                    ),
                    "intento": intento,
                }

            ultimo_error = SnapshotError(
                f"El archivo cambió durante la copia "
                f"(intento {intento}/{intentos})."
            )

        except (
            PermissionError,
            OSError,
        ) as exc:
            ultimo_error = exc

        finally:
            try:
                temporal.unlink()
            except FileNotFoundError:
                pass
            except Exception:
                pass

        if intento < intentos:
            time.sleep(espera_segundos)

    detalle = (
        f" Detalle: {ultimo_error}"
        if ultimo_error is not None
        else ""
    )

    raise SnapshotError(
        f"No se pudo crear un snapshot consistente de "
        f"'{origen}' después de {intentos} intentos."
        f"{detalle}"
    )


def _clave_ruta(ruta: Path) -> str:
    return os.path.normcase(
        str(ruta.resolve())
    )


def _nombre_snapshot_template(
    ruta: Path,
) -> str:
    clave = _clave_ruta(ruta)

    prefijo = hashlib.sha256(
        clave.encode("utf-8")
    ).hexdigest()[:10]

    # Conservamos el nombre original.
    # Esto también conserva reglas como "_porcent_".
    return f"{prefijo}_{ruta.name}"


def congelar_templates_ordenes(
    ordenes: list[OrdenCotizacion],
    directorio_templates: str | Path,
) -> tuple[list[OrdenCotizacion], list[dict[str, Any]]]:
    """
    Copia los TEMPLATE utilizados por las órdenes y devuelve nuevas
    dataclasses apuntando exclusivamente a los snapshots.

    Si varias posiciones usan el mismo template, solo se copia una vez.
    """
    directorio_templates = Path(
        directorio_templates
    )

    directorio_templates.mkdir(
        parents=True,
        exist_ok=True,
    )

    snapshots_por_origen: dict[str, Path] = {}
    metadatos: list[dict[str, Any]] = []

    def obtener_snapshot(
        ruta: Path | None,
    ) -> Path | None:
        if ruta is None:
            return None

        ruta = Path(ruta).resolve()
        clave = _clave_ruta(ruta)

        existente = snapshots_por_origen.get(
            clave
        )

        if existente is not None:
            return existente

        destino = (
            directorio_templates
            / _nombre_snapshot_template(ruta)
        )

        metadata = crear_snapshot_estable(
            ruta,
            destino,
        )

        metadata["tipo"] = "template"

        snapshots_por_origen[clave] = destino
        metadatos.append(metadata)

        return destino

    ordenes_congeladas: list[
        OrdenCotizacion
    ] = []

    for orden in ordenes:
        posiciones_nuevas = []

        for posicion in orden.posiciones_detalle:
            posiciones_nuevas.append(
                replace(
                    posicion,
                    template=obtener_snapshot(
                        posicion.template
                    ),
                )
            )

        ordenes_congeladas.append(
            replace(
                orden,
                template=obtener_snapshot(
                    orden.template
                ),
                posiciones_detalle=posiciones_nuevas,
            )
        )

    return ordenes_congeladas, metadatos


def guardar_manifest_snapshots(
    ruta_manifest: str | Path,
    entrada: dict[str, Any],
    templates: list[dict[str, Any]],
) -> Path:
    ruta_manifest = Path(ruta_manifest)

    ruta_manifest.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    datos = {
        "version": 1,
        "generado_en": datetime.now().isoformat(
            timespec="seconds"
        ),
        "entrada": entrada,
        "templates": templates,
    }

    temporal = ruta_manifest.with_name(
        f".{ruta_manifest.name}."
        f"{os.getpid()}."
        f"{uuid.uuid4().hex}.tmp"
    )

    try:
        temporal.write_text(
            json.dumps(
                datos,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        os.replace(
            temporal,
            ruta_manifest,
        )

    finally:
        try:
            temporal.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass

    return ruta_manifest
