from __future__ import annotations

from tkinter import Tk, messagebox


def _mostrar_error_inicio(mensaje: str) -> None:
    raiz = Tk()
    raiz.withdraw()

    try:
        messagebox.showerror("RPA CBN", mensaje, parent=raiz)
    finally:
        raiz.destroy()


def main() -> int:
    try:
        import customtkinter as ctk

        from src.utils.app_paths import (
            obtener_directorio_app,
            validar_permisos_escritura,
        )

        validar_permisos_escritura()

        directorio_app = obtener_directorio_app()
        ruta_tema = directorio_app / "assets" / "ransa_olive.json"

        if not ruta_tema.exists():
            raise FileNotFoundError(
                "No se encontró el tema visual de la aplicación:\n\n"
                f"{ruta_tema}\n\n"
                "Verifique que exista el archivo "
                "'assets\\ransa_olive.json'."
            )

        # El tema debe cargarse antes de importar y crear la ventana principal.
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme(str(ruta_tema))

        from src.gui.main_window import iniciar_interfaz

        iniciar_interfaz()
        return 0

    except PermissionError as error:
        _mostrar_error_inicio(str(error))
        return 1

    except FileNotFoundError as error:
        _mostrar_error_inicio(str(error))
        return 1

    except ModuleNotFoundError as error:
        modulo = getattr(error, "name", "")

        if modulo == "customtkinter":
            _mostrar_error_inicio(
                "Falta instalar CustomTkinter. Abra PowerShell dentro del entorno "
                "RPA_CBN y ejecute:\n\n"
                "python -m pip install -r .\\requirements_gui.txt"
            )
            return 1

        if modulo in {"PIL", "Pillow"}:
            _mostrar_error_inicio(
                "Falta instalar Pillow para mostrar el logo de la aplicación.\n\n"
                "Ejecute:\n\n"
                "python -m pip install pillow"
            )
            return 1

        raise

    except Exception as error:
        _mostrar_error_inicio(
            "No se pudo iniciar la aplicación.\n\n"
            f"Detalle: {error}"
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
