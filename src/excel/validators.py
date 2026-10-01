from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
import re

import pandas as pd


class ValidacionExcelError(ValueError):
    pass


MAX_ARCHIVO_MB = 50
MAX_ARCHIVO_BYTES = MAX_ARCHIVO_MB * 1024 * 1024

MAX_POSICIONES = 12
MIN_POSICIONES = 1

MAX_TEXTO_POSICION = 40
MAX_TEXTO_COTIZACION = 80

EXTENSIONES_ADJUNTO_VALIDAS = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
}

EXTENSIONES_TEMPLATE_VALIDAS = {
    ".xlsx",
}

TIPOS_IMPUTACION_VALIDOS = {"K", "H", "F", "I"}


@dataclass(frozen=True)
class PosicionCotizacion:
    id_orden: str
    texto_breve: str
    valor: Decimal
    ceco: str
    adjunto: Path | None = None
    template: Path | None = None


@dataclass(frozen=True)
class OrdenCotizacion:
    id_orden: str
    proveedor: str
    texto: str
    valor: Decimal
    moneda: str
    ceco: str | None
    cuenta: str
    servicio: str
    tipo_servicio: str
    posiciones: int
    tipo_imputacion: str
    template: Path | None
    adjuntos: list[Path]
    posiciones_detalle: list[PosicionCotizacion]


def _es_vacio(valor: Any) -> bool:
    if valor is None:
        return True

    try:
        if pd.isna(valor):
            return True
    except Exception:
        pass

    if isinstance(valor, str) and not valor.strip():
        return True

    return False


def normalizar_id(valor: Any, nombre_columna: str = "ID_ORDEN") -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    texto = str(valor).strip()

    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]

    if not texto:
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    return texto


def validar_proveedor(valor: Any) -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError("PROVEEDOR esta vacio.")

    proveedor = str(valor).strip()

    if proveedor.endswith(".0") and proveedor[:-2].isdigit():
        proveedor = proveedor[:-2]

    if not proveedor:
        raise ValidacionExcelError("PROVEEDOR esta vacio.")

    if not re.fullmatch(r"\d+(?:-\d+)*", proveedor):
        raise ValidacionExcelError(
            f"PROVEEDOR invalido: '{proveedor}'. "
            "Debe contener numeros y puede incluir guiones internos."
        )

    # Solo el RUC peruano tradicional tiene exactamente 11 dígitos.
    if not re.fullmatch(r"\d{11}", proveedor):
        print(
            f"WARNING: PROVEEDOR '{proveedor}' no tiene el formato tradicional "
            "de RUC peruano de 11 digitos. Se continuara porque puede ser "
            "un proveedor extranjero."
        )

    return proveedor

def validar_texto(
    valor: Any,
    nombre_columna: str = "TEXTO",
    max_caracteres: int = MAX_TEXTO_POSICION,
) -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    texto = str(valor).strip()

    if not texto:
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    if len(texto) > max_caracteres:
        raise ValidacionExcelError(
            f"{nombre_columna} invalido: tiene {len(texto)} caracteres. "
            f"Maximo permitido: {max_caracteres}."
        )

    return texto


def validar_texto_cotizacion(valor: Any) -> str:
    """
    Texto principal de la cotización.
    CBN permite más caracteres aquí que en la especificación de posición.
    """
    return validar_texto(
        valor=valor,
        nombre_columna="TEXTO",
        max_caracteres=MAX_TEXTO_COTIZACION,
    )


def validar_valor(valor: Any, nombre_columna: str = "Valor") -> Decimal:
    if _es_vacio(valor):
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    texto = str(valor).strip().replace(",", "")

    try:
        numero = Decimal(texto)
    except InvalidOperation as exc:
        raise ValidacionExcelError(
            f"{nombre_columna} invalido: '{valor}'. Debe ser numerico."
        ) from exc

    if numero < 0:
        raise ValidacionExcelError(
            f"{nombre_columna} invalido: '{valor}'. No debe ser negativo."
        )

    return numero.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)


def validar_moneda(valor: Any) -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError("Moneda esta vacia.")

    moneda = str(valor).strip().upper()

    if moneda not in {"PEN", "USD"}:
        raise ValidacionExcelError(
            f"Moneda invalida: '{moneda}'. Valores permitidos: PEN o USD."
        )

    return moneda


def validar_tipo_servicio(valor: Any) -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError("TIPO_SERVICIO esta vacio. Use B para BIEN o S para SERVICIO.")

    tipo = str(valor).strip().upper()

    if tipo not in {"B", "S"}:
        raise ValidacionExcelError(
            f"TIPO_SERVICIO invalido: '{tipo}'. Use B para BIEN o S para SERVICIO."
        )

    return tipo


def validar_tipo_imputacion(valor: Any) -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError(
            "TIPO_IMPUTACION esta vacio. Use K, H, F o I."
        )

    tipo = str(valor).strip().upper()

    if tipo not in TIPOS_IMPUTACION_VALIDOS:
        raise ValidacionExcelError(
            f"TIPO_IMPUTACION invalido: '{tipo}'. Valores permitidos: K, H, F o I."
        )

    return tipo


def validar_posiciones(valor: Any) -> int:
    if _es_vacio(valor):
        raise ValidacionExcelError("POSICIONES esta vacio.")

    try:
        posiciones = int(float(str(valor).strip()))
    except ValueError as exc:
        raise ValidacionExcelError(
            f"POSICIONES invalido: '{valor}'. Debe ser un numero entero."
        ) from exc

    if posiciones < MIN_POSICIONES or posiciones > MAX_POSICIONES:
        raise ValidacionExcelError(
            f"POSICIONES invalido: {posiciones}. Debe estar entre {MIN_POSICIONES} y {MAX_POSICIONES}."
        )

    return posiciones


def validar_codigo_libre(nombre_columna: str, valor: Any) -> str:
    if _es_vacio(valor):
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    codigo = str(valor).strip()

    if codigo.endswith(".0") and codigo[:-2].isdigit():
        codigo = codigo[:-2]

    if not codigo:
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    return codigo


def validar_codigo_opcional(nombre_columna: str, valor: Any) -> str | None:
    if _es_vacio(valor):
        return None

    codigo = str(valor).strip()

    if codigo.endswith(".0") and codigo[:-2].isdigit():
        codigo = codigo[:-2]

    if not codigo:
        return None

    return codigo


def _validar_archivo_base(valor: Any, extensiones_validas: set[str], nombre_columna: str) -> Path:
    if _es_vacio(valor):
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    texto_ruta = str(valor).strip().strip('"')

    if not texto_ruta:
        raise ValidacionExcelError(f"{nombre_columna} esta vacio.")

    ruta = Path(texto_ruta).expanduser().resolve()

    if ruta.name.startswith("~$"):
        raise ValidacionExcelError(f"Archivo temporal de Excel no permitido: {ruta}")

    if not ruta.exists:
        raise ValidacionExcelError(f"No existe el archivo: {ruta}")

    if not ruta.exists():
        raise ValidacionExcelError(f"No existe el archivo: {ruta}")

    if not ruta.is_file():
        raise ValidacionExcelError(f"La ruta no es un archivo: {ruta}")

    if ruta.suffix.lower() not in extensiones_validas:
        raise ValidacionExcelError(
            f"Extension no permitida para {nombre_columna}: {ruta.suffix}. Archivo: {ruta}"
        )

    tamano_bytes = ruta.stat().st_size

    if tamano_bytes > MAX_ARCHIVO_BYTES:
        tamano_mb = tamano_bytes / (1024 * 1024)
        raise ValidacionExcelError(
            f"El archivo supera {MAX_ARCHIVO_MB} MB: {ruta} ({tamano_mb:.2f} MB)"
        )

    return ruta


def validar_adjunto(valor: Any) -> Path:
    return _validar_archivo_base(valor, EXTENSIONES_ADJUNTO_VALIDAS, "Archivo")


def validar_adjunto_opcional(valor: Any) -> Path | None:
    if _es_vacio(valor):
        return None

    return validar_adjunto(valor)


def validar_template_opcional(valor: Any) -> Path | None:
    if _es_vacio(valor):
        return None

    return _validar_archivo_base(valor, EXTENSIONES_TEMPLATE_VALIDAS, "TEMPLATE")



def construir_posicion_desde_fila(fila: pd.Series) -> PosicionCotizacion:
    """
    Construye una posición desde una fila.

    Compatible con:
    - hoja antigua Posiciones: VALOR
    - nueva hoja Ordenes agrupada: Valor
    """
    valor_raw = fila.get("VALOR")

    if _es_vacio(valor_raw):
        valor_raw = fila.get("Valor")

    texto_breve_raw = fila.get("TEXTO_BREVE")

    if _es_vacio(texto_breve_raw):
        # Para textos cortos se permite reutilizar el nombre de la cotización.
        # Si TEXTO supera 40 caracteres, validar_texto mostrará el error para
        # que el usuario complete TEXTO_BREVE.
        texto_breve_raw = fila.get("TEXTO")

    return PosicionCotizacion(
        id_orden=normalizar_id(fila["ID_ORDEN"]),
        texto_breve=validar_texto(texto_breve_raw, "TEXTO_BREVE"),
        valor=validar_valor(valor_raw, "Valor de posicion"),
        ceco=validar_codigo_libre("CECO", fila["CECO"]),
        adjunto=validar_adjunto_opcional(fila.get("ADJUNTO")),
        template=validar_template_opcional(fila.get("TEMPLATE")),
    )


def construir_orden_desde_fila(
    fila: pd.Series,
    adjuntos: list[Path],
    posiciones_detalle: list[PosicionCotizacion],
) -> OrdenCotizacion:
    """
    Compatibilidad con lectores/checkpoints antiguos.

    En el modelo nuevo se recomienda construir la orden con
    construir_orden_desde_grupo().
    """
    posiciones = len(posiciones_detalle)

    if posiciones == 0:
        posiciones_raw = fila.get("POSICIONES")

        if not _es_vacio(posiciones_raw):
            posiciones = validar_posiciones(posiciones_raw)
        else:
            posiciones = 1

    valor_total = validar_valor(fila["Valor"])

    if posiciones_detalle:
        valor_total = sum(
            (posicion.valor for posicion in posiciones_detalle),
            Decimal("0"),
        ).quantize(Decimal("0.0001"))

    return OrdenCotizacion(
        id_orden=normalizar_id(fila["ID_ORDEN"]),
        proveedor=validar_proveedor(fila["PROVEEDOR"]),
        texto=validar_texto_cotizacion(fila["TEXTO"]),
        valor=valor_total,
        moneda=validar_moneda(fila["Moneda"]),
        ceco=None if posiciones_detalle else validar_codigo_opcional("CECO", fila.get("CECO")),
        cuenta=validar_codigo_libre("Cuenta", fila["Cuenta"]),
        servicio=validar_codigo_libre("Servicio", fila["Servicio"]),
        tipo_servicio=validar_tipo_servicio(fila["TIPO_SERVICIO"]),
        posiciones=posiciones,
        tipo_imputacion=validar_tipo_imputacion(fila["TIPO_IMPUTACION"]),
        template=None if posiciones_detalle else validar_template_opcional(fila.get("TEMPLATE")),
        adjuntos=adjuntos,
        posiciones_detalle=posiciones_detalle,
    )


def _valor_general_grupo(
    fila: pd.Series,
    primera_fila: pd.Series,
    columna: str,
) -> Any:
    """
    En filas adicionales del mismo ID_ORDEN los datos generales pueden:
    - repetirse, o
    - quedar vacíos para heredar el valor de la primera fila.
    """
    valor = fila.get(columna)

    if _es_vacio(valor):
        return primera_fila.get(columna)

    return valor


def construir_orden_desde_grupo(
    grupo: pd.DataFrame,
    adjuntos: list[Path],
) -> OrdenCotizacion:
    """
    Nueva estructura:
    - Cada ID_ORDEN representa una sola cotización.
    - Cada fila de ese ID_ORDEN representa una posición.
    - Valor es el importe de esa posición.
    - TEXTO_BREVE, CECO y TEMPLATE pertenecen a esa posición.
    - POSICIONES ya no es necesario; se calcula contando filas.
    """
    if grupo.empty:
        raise ValidacionExcelError("No se puede construir una orden desde un grupo vacío.")

    primera_fila = grupo.iloc[0]

    id_orden = normalizar_id(primera_fila["ID_ORDEN"])
    proveedor = validar_proveedor(primera_fila["PROVEEDOR"])
    texto = validar_texto_cotizacion(primera_fila["TEXTO"])
    moneda = validar_moneda(primera_fila["Moneda"])
    cuenta = validar_codigo_libre("Cuenta", primera_fila["Cuenta"])
    servicio = validar_codigo_libre("Servicio", primera_fila["Servicio"])
    tipo_servicio = validar_tipo_servicio(primera_fila["TIPO_SERVICIO"])
    tipo_imputacion = validar_tipo_imputacion(primera_fila["TIPO_IMPUTACION"])

    posiciones_detalle: list[PosicionCotizacion] = []

    for numero_posicion, (_, fila_original) in enumerate(grupo.iterrows(), start=1):
        fila = fila_original.copy()

        # Datos generales: pueden repetirse o quedar vacíos después de la primera fila.
        for columna in (
            "PROVEEDOR",
            "TEXTO",
            "Moneda",
            "Cuenta",
            "Servicio",
            "TIPO_SERVICIO",
            "TIPO_IMPUTACION",
        ):
            fila[columna] = _valor_general_grupo(fila, primera_fila, columna)

        id_fila = normalizar_id(fila["ID_ORDEN"])

        if id_fila != id_orden:
            raise ValidacionExcelError(
                f"El grupo contiene ID_ORDEN diferentes: '{id_orden}' y '{id_fila}'."
            )

        proveedor_fila = validar_proveedor(fila["PROVEEDOR"])
        texto_fila = validar_texto_cotizacion(fila["TEXTO"])
        moneda_fila = validar_moneda(fila["Moneda"])
        cuenta_fila = validar_codigo_libre("Cuenta", fila["Cuenta"])
        servicio_fila = validar_codigo_libre("Servicio", fila["Servicio"])
        tipo_servicio_fila = validar_tipo_servicio(fila["TIPO_SERVICIO"])
        tipo_imputacion_fila = validar_tipo_imputacion(fila["TIPO_IMPUTACION"])

        comparaciones = [
            ("PROVEEDOR", proveedor_fila, proveedor),
            ("TEXTO", texto_fila, texto),
            ("Moneda", moneda_fila, moneda),
            ("Cuenta", cuenta_fila, cuenta),
            ("Servicio", servicio_fila, servicio),
            ("TIPO_SERVICIO", tipo_servicio_fila, tipo_servicio),
            ("TIPO_IMPUTACION", tipo_imputacion_fila, tipo_imputacion),
        ]

        for columna, actual, esperado in comparaciones:
            if actual != esperado:
                raise ValidacionExcelError(
                    f"Orden {id_orden}, posicion {numero_posicion}: "
                    f"el campo general {columna}='{actual}' no coincide con "
                    f"la primera fila del grupo ('{esperado}')."
                )

        try:
            posicion = construir_posicion_desde_fila(fila)
        except ValidacionExcelError as exc:
            raise ValidacionExcelError(
                f"Orden {id_orden}, posicion {numero_posicion}: {exc}"
            ) from exc

        posiciones_detalle.append(posicion)

    if len(posiciones_detalle) > MAX_POSICIONES:
        raise ValidacionExcelError(
            f"La orden '{id_orden}' tiene {len(posiciones_detalle)} posiciones. "
            f"Máximo permitido: {MAX_POSICIONES}."
        )

    # Regla temporal para BIEN:
    # Hasta contar con un caso real de varios bienes/posiciones,
    # una orden de tipo B debe tener exactamente una sola posición.
    if tipo_servicio == "B" and len(posiciones_detalle) != 1:
        raise ValidacionExcelError(
            f"La orden '{id_orden}' es de tipo BIEN y tiene "
            f"{len(posiciones_detalle)} posiciones. "
            "Por ahora, las órdenes de BIEN deben tener exactamente 1 posición."
        )

    valor_total = sum(
        (posicion.valor for posicion in posiciones_detalle),
        Decimal("0"),
    ).quantize(Decimal("0.0001"))

    return OrdenCotizacion(
        id_orden=id_orden,
        proveedor=proveedor,
        texto=texto,
        valor=valor_total,
        moneda=moneda,
        ceco=None,
        cuenta=cuenta,
        servicio=servicio,
        tipo_servicio=tipo_servicio,
        posiciones=len(posiciones_detalle),
        tipo_imputacion=tipo_imputacion,
        template=None,
        adjuntos=adjuntos,
        posiciones_detalle=posiciones_detalle,
    )

