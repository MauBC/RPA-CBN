from __future__ import annotations

from contextlib import contextmanager, redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Callable, TextIO
import logging
import sys
import threading


EventCallback = Callable[[dict], None]


def _emitir(callback: EventCallback | None, evento: dict) -> None:
    if callback is None:
        return

    try:
        callback(evento)
    except Exception:
        pass


class _TeePorLinea:
    """
    Copia stdout/stderr a:
    - consola original, cuando existe;
    - ejecucion.log;
    - callback de la interfaz.
    """

    def __init__(
        self,
        original: TextIO | None,
        logger: logging.Logger,
        nivel: int,
        callback: EventCallback | None,
    ) -> None:
        self.original = original
        self.logger = logger
        self.nivel = nivel
        self.callback = callback
        self._buffer = ""
        self._lock = threading.RLock()

    @property
    def encoding(self) -> str:
        return "utf-8"

    def isatty(self) -> bool:
        return False

    def writable(self) -> bool:
        return True

    def write(self, texto: str) -> int:
        if texto is None:
            return 0

        texto = str(texto)

        with self._lock:
            if self.original is not None:
                try:
                    self.original.write(texto)
                    self.original.flush()
                except Exception:
                    pass

            self._buffer += texto

            while "\n" in self._buffer:
                linea, self._buffer = self._buffer.split("\n", 1)
                self._procesar_linea(linea)

        return len(texto)

    def flush(self) -> None:
        with self._lock:
            if self._buffer:
                self._procesar_linea(self._buffer)
                self._buffer = ""

            if self.original is not None:
                try:
                    self.original.flush()
                except Exception:
                    pass

    def _procesar_linea(self, linea: str) -> None:
        linea = linea.rstrip("\r")

        if not linea.strip():
            return

        self.logger.log(self.nivel, linea)

        _emitir(
            self.callback,
            {
                "type": "log",
                "level": logging.getLevelName(self.nivel),
                "message": linea,
                "timestamp": datetime.now().isoformat(timespec="seconds"),
            },
        )


def crear_logger_ejecucion(directorio: str | Path) -> tuple[logging.Logger, Path]:
    directorio = Path(directorio)
    directorio.mkdir(parents=True, exist_ok=True)
    ruta_log = directorio / "ejecucion.log"

    nombre = f"rpa_cbn_ejecucion_{id(directorio)}_{threading.get_ident()}"
    logger = logging.getLogger(nombre)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        try:
            handler.close()
        except Exception:
            pass

    handler = logging.FileHandler(ruta_log, encoding="utf-8")
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)-7s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(handler)

    return logger, ruta_log


@contextmanager
def capturar_salida_ejecucion(
    directorio: str | Path,
    callback: EventCallback | None = None,
):
    logger, ruta_log = crear_logger_ejecucion(directorio)

    salida = _TeePorLinea(
        original=getattr(sys, "__stdout__", None),
        logger=logger,
        nivel=logging.INFO,
        callback=callback,
    )
    errores = _TeePorLinea(
        original=getattr(sys, "__stderr__", None),
        logger=logger,
        nivel=logging.ERROR,
        callback=callback,
    )

    try:
        with redirect_stdout(salida), redirect_stderr(errores):
            yield ruta_log
    finally:
        salida.flush()
        errores.flush()

        for handler in list(logger.handlers):
            try:
                handler.flush()
                handler.close()
            except Exception:
                pass
            logger.removeHandler(handler)
