from contextlib import contextmanager
from typing import Iterator


class PasoRPAError(RuntimeError):
    def __init__(self, paso: str, error_original: Exception):
        self.paso = paso
        self.error_original = error_original
        super().__init__(f"Fallo en paso RPA: {paso}. Detalle: {error_original}")


@contextmanager
def paso_rpa(nombre_paso: str) -> Iterator[None]:
    try:
        yield
    except PasoRPAError:
        raise
    except Exception as error:
        raise PasoRPAError(nombre_paso, error) from error
