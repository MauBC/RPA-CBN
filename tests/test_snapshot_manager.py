from decimal import Decimal
from pathlib import Path

from src.excel.snapshot_manager import (
    congelar_templates_ordenes,
    crear_snapshot_estable,
    guardar_manifest_snapshots,
    sha256_archivo,
)

from src.excel.validators import (
    OrdenCotizacion,
    PosicionCotizacion,
)


def _crear_archivo(ruta: Path, contenido: bytes) -> Path:
    ruta.write_bytes(contenido)
    return ruta


def test_snapshot_estable_es_identico_al_origen(tmp_path):
    origen = _crear_archivo(
        tmp_path / "origen.xlsx",
        b"contenido-original",
    )

    destino = tmp_path / "snapshots" / "entrada.xlsx"

    metadata = crear_snapshot_estable(
        origen,
        destino,
    )

    assert destino.exists()
    assert destino.read_bytes() == origen.read_bytes()

    assert (
        sha256_archivo(destino)
        == sha256_archivo(origen)
        == metadata["sha256"]
    )


def test_snapshot_no_cambia_si_original_cambia_despues(tmp_path):
    origen = _crear_archivo(
        tmp_path / "DATA.xlsx",
        b"VERSION-1",
    )

    destino = tmp_path / "entrada_original.xlsx"

    crear_snapshot_estable(
        origen,
        destino,
    )

    origen.write_bytes(
        b"VERSION-2-MODIFICADA"
    )

    assert destino.read_bytes() == b"VERSION-1"
    assert origen.read_bytes() != destino.read_bytes()


def test_templates_repetidos_se_copian_una_sola_vez(tmp_path):
    template = _crear_archivo(
        tmp_path / "plantilla.xlsx",
        b"TEMPLATE",
    )

    posicion_1 = PosicionCotizacion(
        id_orden="1001",
        texto_breve="Posicion 1",
        valor=Decimal("50"),
        ceco="100010",
        template=template,
    )

    posicion_2 = PosicionCotizacion(
        id_orden="1001",
        texto_breve="Posicion 2",
        valor=Decimal("50"),
        ceco="100010",
        template=template,
    )

    orden = OrdenCotizacion(
        id_orden="1001",
        proveedor="PROV",
        texto="Prueba",
        valor=Decimal("100"),
        moneda="PEN",
        ceco=None,
        cuenta="CUENTA",
        servicio="SERVICIO",
        tipo_servicio="SERVICIO",
        posiciones=2,
        tipo_imputacion="K",
        template=None,
        adjuntos=[],
        posiciones_detalle=[
            posicion_1,
            posicion_2,
        ],
    )

    congeladas, metadata = congelar_templates_ordenes(
        [orden],
        tmp_path / "templates",
    )

    assert len(metadata) == 1

    ruta_1 = congeladas[0].posiciones_detalle[0].template
    ruta_2 = congeladas[0].posiciones_detalle[1].template

    assert ruta_1 is not None
    assert ruta_1 == ruta_2
    assert ruta_1 != template.resolve()
    assert ruta_1.exists()


def test_snapshot_template_conserva_nombre_porcent(tmp_path):
    template = _crear_archivo(
        tmp_path / "plantilla_imputacion_porcent_google.xlsx",
        b"TEMPLATE",
    )

    posicion = PosicionCotizacion(
        id_orden="1001",
        texto_breve="Posicion",
        valor=Decimal("100"),
        ceco="100010",
        template=template,
    )

    orden = OrdenCotizacion(
        id_orden="1001",
        proveedor="PROV",
        texto="Prueba",
        valor=Decimal("100"),
        moneda="PEN",
        ceco=None,
        cuenta="CUENTA",
        servicio="SERVICIO",
        tipo_servicio="SERVICIO",
        posiciones=1,
        tipo_imputacion="K",
        template=None,
        adjuntos=[],
        posiciones_detalle=[posicion],
    )

    congeladas, _ = congelar_templates_ordenes(
        [orden],
        tmp_path / "templates",
    )

    ruta_snapshot = (
        congeladas[0]
        .posiciones_detalle[0]
        .template
    )

    assert ruta_snapshot is not None
    assert "_porcent_" in ruta_snapshot.name.lower()


def test_manifest_se_genera(tmp_path):
    origen = _crear_archivo(
        tmp_path / "DATA.xlsx",
        b"DATA",
    )

    snapshot = tmp_path / "entrada_original.xlsx"

    metadata = crear_snapshot_estable(
        origen,
        snapshot,
    )

    metadata["tipo"] = "excel_principal"

    manifest = guardar_manifest_snapshots(
        tmp_path / "inputs" / "snapshot_manifest.json",
        metadata,
        [],
    )

    assert manifest.exists()

    contenido = manifest.read_text(
        encoding="utf-8"
    )

    assert "excel_principal" in contenido
    assert metadata["sha256"] in contenido
