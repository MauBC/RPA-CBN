from __future__ import annotations

import re
from typing import Any


def _texto_error(error: Exception) -> str:
    paso = getattr(error, "paso", "")
    original = getattr(error, "error_original", error)

    return f"{paso} {original}".lower()


def _extraer_posicion(texto: str) -> int | None:
    match = re.search(r"posici[oó]n\s+(\d+)", texto, flags=re.IGNORECASE)

    if not match:
        return None

    try:
        return int(match.group(1))
    except ValueError:
        return None


def _codigo_imputacion_por_posicion(orden: Any, posicion_numero: int | None) -> str | None:
    if orden is None:
        return None

    if posicion_numero is None:
        return getattr(orden, "ceco", None)

    try:
        if getattr(orden, "posiciones_detalle", None):
            posicion = orden.posiciones_detalle[posicion_numero - 1]
            return getattr(posicion, "ceco", None)
    except Exception:
        pass

    return getattr(orden, "ceco", None)


def _nombre_tipo_imputacion(orden: Any) -> str:
    tipo = str(getattr(orden, "tipo_imputacion", "") or "").strip().upper()

    return {
        "K": "CECO",
        "H": "centro de beneficio",
        "F": "orden interna",
        "I": "orden de inversión",
    }.get(tipo, "código de imputación")


def construir_mensaje_usuario(error: Exception, orden: Any | None = None) -> str:
    texto = _texto_error(error)
    paso = str(getattr(error, "paso", "") or "")
    posicion_numero = _extraer_posicion(f"{paso} {texto}")

    proveedor = getattr(orden, "proveedor", None)
    servicio = getattr(orden, "servicio", None)
    cuenta = getattr(orden, "cuenta", None)
    codigo_imputacion = _codigo_imputacion_por_posicion(orden, posicion_numero)
    nombre_imputacion = _nombre_tipo_imputacion(orden)

    if "proveedor" in texto:
        if proveedor:
            return f"No se encontró el proveedor {proveedor} en CBN."
        return "No se encontró el proveedor en CBN."

    if "servicio" in texto or "bien" in texto or "codigo de bien" in texto or "código de bien" in texto:
        if servicio:
            return f"No se encontró el código de bien/servicio {servicio} en CBN."
        return "No se encontró el código de bien/servicio en CBN."

    if "cuenta contable" in texto or "cuenta" in texto:
        if cuenta:
            return f"No se encontró la cuenta contable {cuenta} en CBN."
        return "No se encontró la cuenta contable en CBN."

    if (
        "centro de costo" in texto
        or "ceco" in texto
        or "centro de beneficio" in texto
        or "orden interna" in texto
        or "orden de inversión" in texto
        or "orden de inversion" in texto
        or "asignar imputacion" in texto
        or "asignar imputación" in texto
    ):
        if codigo_imputacion:
            if posicion_numero:
                return (
                    f"No se encontró el {nombre_imputacion} {codigo_imputacion} "
                    f"en CBN para la posición {posicion_numero}."
                )

            return f"No se encontró el {nombre_imputacion} {codigo_imputacion} en CBN."

        return f"No se encontró el {nombre_imputacion} en CBN."

    if "template" in texto or "imputacion multiple" in texto or "imputación múltiple" in texto:
        if posicion_numero:
            return (
                f"No se pudo cargar el TEMPLATE de imputación múltiple "
                f"para la posición {posicion_numero}."
            )

        return "No se pudo cargar el TEMPLATE de imputación múltiple."

    if "adjunt" in texto or "archivo" in texto:
        return "No se pudo adjuntar uno de los archivos indicados."


    if "mandar a cotizar" in texto or "forma de pago" in texto:
        return "No se pudo completar la fase mandar a cotizar en CBN."

    if "cotizar posicion" in texto or "cotizar posición" in texto or "valor unitario" in texto:
        return "No se pudo cotizar una de las posiciones en CBN."

    if "timeout" in texto or "demor" in texto or "expected to be visible" in texto:
        return "CBN demoró demasiado en responder. Intenta nuevamente."

    return "Ocurrió un error inesperado durante el proceso. Se guardó evidencia técnica para revisión."


def construir_detalle_tecnico(error: Exception) -> str:
    paso = getattr(error, "paso", "")
    original = getattr(error, "error_original", error)

    if paso:
        return f"Paso: {paso} | Detalle: {original}"

    return str(error)
