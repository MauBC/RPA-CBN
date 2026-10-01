from __future__ import annotations

from pathlib import Path
import ctypes
from ctypes import wintypes
import os

import pytest

from src.excel.file_access import (
    EstadoAccesoExcel,
    diagnosticar_acceso_excel,
)


def test_diagnostico_archivo_inexistente(tmp_path):
    ruta = tmp_path / "NO_EXISTE.xlsx"

    resultado = diagnosticar_acceso_excel(
        ruta
    )

    assert (
        resultado.estado
        == EstadoAccesoExcel.NO_EXISTE
    )

    assert resultado.puede_escribir is False


def test_diagnostico_archivo_libre(tmp_path):
    ruta = tmp_path / "DATA.xlsx"
    ruta.write_bytes(b"prueba")

    resultado = diagnosticar_acceso_excel(
        ruta
    )

    assert (
        resultado.estado
        == EstadoAccesoExcel.DISPONIBLE
    )

    assert resultado.puede_escribir is True


def test_lock_office_es_solo_indicio(tmp_path):
    ruta = tmp_path / "DATA.xlsx"
    ruta.write_bytes(b"prueba")

    lock_office = tmp_path / "~$DATA.xlsx"
    lock_office.write_bytes(b"lock")

    resultado = diagnosticar_acceso_excel(
        ruta
    )

    assert resultado.lock_office_detectado is True

    # Un ~$ por sí solo no debe causar un falso positivo.
    assert (
        resultado.estado
        == EstadoAccesoExcel.DISPONIBLE
    )


@pytest.mark.skipif(
    os.name != "nt",
    reason="Test específico de Windows",
)
def test_detecta_archivo_bloqueado_para_escritura_windows(
    tmp_path,
):
    ruta = tmp_path / "DATA.xlsx"
    ruta.write_bytes(b"prueba")

    kernel32 = ctypes.WinDLL(
        "kernel32",
        use_last_error=True,
    )

    create_file = kernel32.CreateFileW

    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]

    create_file.restype = wintypes.HANDLE

    close_handle = kernel32.CloseHandle

    close_handle.argtypes = [
        wintypes.HANDLE,
    ]

    close_handle.restype = wintypes.BOOL

    GENERIC_READ = 0x80000000

    # Permitimos que otros lean, pero NO que escriban.
    FILE_SHARE_READ = 0x00000001

    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x00000080

    invalid_handle = ctypes.c_void_p(-1).value

    handle = create_file(
        str(ruta),
        GENERIC_READ,
        FILE_SHARE_READ,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None,
    )

    handle_value = ctypes.cast(
        handle,
        ctypes.c_void_p,
    ).value

    assert handle_value != invalid_handle

    try:
        resultado = diagnosticar_acceso_excel(
            ruta
        )

        assert (
            resultado.estado
            == EstadoAccesoExcel.OCUPADO
        )

        assert resultado.puede_escribir is False

        assert resultado.winerror in (
            32,
            33,
        )

    finally:
        close_handle(handle)


def test_diagnostico_no_modifica_archivo(tmp_path):
    ruta = tmp_path / "DATA.xlsx"

    contenido = (
        b"contenido-que-no-debe-cambiar"
    )

    ruta.write_bytes(contenido)

    antes = ruta.read_bytes()

    diagnosticar_acceso_excel(
        ruta
    )

    despues = ruta.read_bytes()

    assert despues == antes
