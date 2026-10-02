from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
import re
import traceback

from src.utils.app_paths import (
    obtener_directorio_ejecucion,
    obtener_directorio_salida_compatibilidad,
    obtener_directorio_logs_compatibilidad,
)


def _limpiar_contexto(contexto: str) -> str:
    texto = re.sub(r"[^A-Za-z0-9._-]+", "_", str(contexto or "error"))
    return texto.strip("._-") or "error"


def preparar_carpetas() -> None:
    ejecucion = obtener_directorio_ejecucion()

    if ejecucion is not None:
        (ejecucion / "errores").mkdir(parents=True, exist_ok=True)
        (ejecucion / "temporal").mkdir(parents=True, exist_ok=True)
        return

    salida = obtener_directorio_salida_compatibilidad()
    (salida / "screenshots").mkdir(parents=True, exist_ok=True)
    (salida / "html_debug").mkdir(parents=True, exist_ok=True)
    obtener_directorio_logs_compatibilidad()


def obtener_timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _directorio_error(contexto: str, timestamp: str) -> Path:
    ejecucion = obtener_directorio_ejecucion()
    contexto = _limpiar_contexto(contexto)

    if ejecucion is not None:
        ruta = ejecucion / "errores" / f"{contexto}_{timestamp}"
        ruta.mkdir(parents=True, exist_ok=True)
        return ruta

    # Compatibilidad para scripts antiguos ejecutados fuera del runner.
    ruta = obtener_directorio_salida_compatibilidad() / "errores" / f"{contexto}_{timestamp}"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def _obtener_estado_scroll(
    page: Any,
) -> dict[str, float]:
    """
    Detecta el área vertical con mayor capacidad de scroll.

    Algunas pantallas del portal usan un contenedor interno
    en lugar del scroll principal de window/document.
    """
    resultado = page.evaluate(
        """
        () => {
            const elementos = [
                document.scrollingElement,
                ...Array.from(
                    document.querySelectorAll('*')
                )
            ].filter(Boolean);

            let mejor = document.scrollingElement;
            let mejorRango = mejor
                ? Math.max(
                    0,
                    mejor.scrollHeight
                    - mejor.clientHeight
                )
                : 0;

            for (const elemento of elementos) {
                const estilo = getComputedStyle(
                    elemento
                );

                const overflowY = estilo.overflowY;

                if (
                    overflowY !== 'auto'
                    && overflowY !== 'scroll'
                ) {
                    continue;
                }

                const rango = Math.max(
                    0,
                    elemento.scrollHeight
                    - elemento.clientHeight
                );

                if (rango > mejorRango) {
                    mejor = elemento;
                    mejorRango = rango;
                }
            }

            if (!mejor) {
                return {
                    top: 0,
                    viewport: window.innerHeight || 800,
                    max_top: 0
                };
            }

            if (
                mejor === document.body
                || mejor === document.documentElement
                || mejor === document.scrollingElement
            ) {
                mejor = document.scrollingElement;
            }

            return {
                top: Number(mejor.scrollTop || 0),
                viewport: Number(
                    mejor.clientHeight
                    || window.innerHeight
                    || 800
                ),
                max_top: Number(
                    Math.max(
                        0,
                        mejor.scrollHeight
                        - mejor.clientHeight
                    )
                )
            };
        }
        """
    )

    if not isinstance(
        resultado,
        dict,
    ):
        return {
            "top": 0.0,
            "viewport": 800.0,
            "max_top": 0.0,
        }

    return {
        "top": float(
            resultado.get(
                "top",
                0,
            )
            or 0
        ),
        "viewport": max(
            1.0,
            float(
                resultado.get(
                    "viewport",
                    800,
                )
                or 800
            ),
        ),
        "max_top": max(
            0.0,
            float(
                resultado.get(
                    "max_top",
                    0,
                )
                or 0
            ),
        ),
    }


def _mover_scroll_debug(
    page: Any,
    posicion: float,
) -> None:
    page.evaluate(
        """
        (posicion) => {
            const elementos = [
                document.scrollingElement,
                ...Array.from(
                    document.querySelectorAll('*')
                )
            ].filter(Boolean);

            let mejor = document.scrollingElement;
            let mejorRango = mejor
                ? Math.max(
                    0,
                    mejor.scrollHeight
                    - mejor.clientHeight
                )
                : 0;

            for (const elemento of elementos) {
                const estilo = getComputedStyle(
                    elemento
                );

                const overflowY = estilo.overflowY;

                if (
                    overflowY !== 'auto'
                    && overflowY !== 'scroll'
                ) {
                    continue;
                }

                const rango = Math.max(
                    0,
                    elemento.scrollHeight
                    - elemento.clientHeight
                );

                if (rango > mejorRango) {
                    mejor = elemento;
                    mejorRango = rango;
                }
            }

            if (!mejor) {
                return;
            }

            mejor.scrollTop = Math.max(
                0,
                Math.min(
                    Number(posicion || 0),
                    Math.max(
                        0,
                        mejor.scrollHeight
                        - mejor.clientHeight
                    )
                )
            );
        }
        """,
        float(
            posicion
        ),
    )

    try:
        page.wait_for_timeout(
            120
        )
    except Exception:
        pass


def _capturar_contexto_visual(
    page: Any,
    directorio: Path,
    rutas: dict[str, Path],
) -> None:
    """
    Guarda tres viewports alrededor del punto de error y
    restaura la posición original.

    También conserva captura.png full_page para mantener
    compatibilidad con el reporte existente.
    """
    estado = _obtener_estado_scroll(
        page
    )

    posicion_original = estado[
        "top"
    ]

    viewport = estado[
        "viewport"
    ]

    max_top = estado[
        "max_top"
    ]

    desplazamiento = max(
        200.0,
        viewport * 0.70,
    )

    posicion_arriba = max(
        0.0,
        posicion_original
        - desplazamiento,
    )

    posicion_abajo = min(
        max_top,
        posicion_original
        + desplazamiento,
    )

    ruta_actual = (
        directorio
        / "captura_actual.png"
    )

    page.screenshot(
        path=str(
            ruta_actual
        ),
        full_page=False,
    )

    rutas[
        "screenshot_actual"
    ] = ruta_actual

    try:
        _mover_scroll_debug(
            page,
            posicion_arriba,
        )

        ruta_arriba = (
            directorio
            / "captura_arriba.png"
        )

        page.screenshot(
            path=str(
                ruta_arriba
            ),
            full_page=False,
        )

        rutas[
            "screenshot_arriba"
        ] = ruta_arriba

        _mover_scroll_debug(
            page,
            posicion_abajo,
        )

        ruta_abajo = (
            directorio
            / "captura_abajo.png"
        )

        page.screenshot(
            path=str(
                ruta_abajo
            ),
            full_page=False,
        )

        rutas[
            "screenshot_abajo"
        ] = ruta_abajo

    finally:
        try:
            _mover_scroll_debug(
                page,
                posicion_original,
            )
        except Exception:
            pass

    ruta_completa = (
        directorio
        / "captura.png"
    )

    page.screenshot(
        path=str(
            ruta_completa
        ),
        full_page=True,
    )

    # Mantiene el contrato histórico del resultado Excel.
    rutas[
        "screenshot"
    ] = ruta_completa


def guardar_evidencia_error(
    page: Any | None,
    error: Exception,
    contexto: str = "error",
) -> dict[str, Path]:
    preparar_carpetas()

    timestamp = obtener_timestamp()
    directorio = _directorio_error(contexto, timestamp)

    # Defensa adicional: guardar_evidencia_error garantiza
    # que su destino exista incluso si el helper fue reemplazado
    # o si una integración externa devuelve una ruta aún no creada.
    directorio.mkdir(
        parents=True,
        exist_ok=True,
    )

    rutas: dict[str, Path] = {}

    ruta_log = directorio / "error.log"
    ruta_log.write_text(
        "ERROR:\n"
        f"{repr(error)}\n\n"
        "TRACEBACK:\n"
        f"{traceback.format_exc()}\n",
        encoding="utf-8",
    )
    rutas["log"] = ruta_log

    if page is None:
        return rutas

    try:
        _capturar_contexto_visual(
            page,
            directorio,
            rutas,
        )
    except Exception:
        # La evidencia visual nunca debe ocultar el error real.
        try:
            ruta_screenshot = (
                directorio
                / "captura.png"
            )

            page.screenshot(
                path=str(
                    ruta_screenshot
                ),
                full_page=True,
            )

            rutas[
                "screenshot"
            ] = ruta_screenshot

        except Exception:
            pass

    try:
        ruta_html = directorio / "pagina.html"
        ruta_html.write_text(page.content(), encoding="utf-8")
        rutas["html"] = ruta_html
    except Exception:
        pass

    return rutas
