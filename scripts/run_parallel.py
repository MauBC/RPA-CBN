from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


BASE_PERFIL = Path("perfil_rpa")
WORKERS_DEFAULT = 2


def _limpiar_locks_chrome(carpeta: Path) -> None:
    if not carpeta.exists():
        return

    patrones = [
        "Singleton*",
        "lockfile",
        "DevToolsActivePort",
    ]

    for patron in patrones:
        for path in carpeta.glob(patron):
            try:
                if path.is_dir():
                    shutil.rmtree(path, ignore_errors=True)
                else:
                    path.unlink(missing_ok=True)
            except Exception:
                pass


def _preparar_perfil_worker(worker_id: int) -> Path:
    destino = Path(f"perfil_rpa_worker_{worker_id}")

    if not destino.exists():
        if BASE_PERFIL.exists():
            print(f"Creando {destino} copiando sesión desde {BASE_PERFIL}...")
            shutil.copytree(
                BASE_PERFIL,
                destino,
                ignore=shutil.ignore_patterns(
                    "Singleton*",
                    "lockfile",
                    "DevToolsActivePort",
                    "CrashpadMetrics-active.pma",
                ),
            )
        else:
            print(f"No existe {BASE_PERFIL}. Creando perfil vacío: {destino}")
            destino.mkdir(parents=True, exist_ok=True)

    _limpiar_locks_chrome(destino)

    return destino


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workers", type=int, default=WORKERS_DEFAULT)
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Muestra los navegadores. Por defecto corre headless.",
    )
    parser.add_argument(
        "--slow-mo",
        type=int,
        default=None,
        help="Milisegundos de pausa por acción. Default: 0 headless, 300 visible.",
    )

    args = parser.parse_args()

    if args.workers != 2:
        raise SystemExit("Por ahora usa --workers 2. Luego se puede ampliar.")

    headless = not args.visible
    slow_mo = args.slow_mo

    outputs = Path("outputs")
    outputs.mkdir(parents=True, exist_ok=True)

    procesos = []

    print("Lanzando RPA paralelo")
    print(f"Workers: {args.workers}")
    print(f"Headless: {headless}")
    print("Worker 1 = ID_ORDEN impares")
    print("Worker 2 = ID_ORDEN pares")
    print()

    for worker_id in range(1, args.workers + 1):
        perfil = _preparar_perfil_worker(worker_id)
        log_path = outputs / f"parallel_worker_{worker_id}.log"

        env = os.environ.copy()
        env["RPA_WORKERS"] = str(args.workers)
        env["RPA_WORKER_ID"] = str(worker_id)
        env["RPA_PROFILE"] = str(perfil)
        env["RPA_HEADLESS"] = "1" if headless else "0"

        if slow_mo is not None:
            env["RPA_SLOW_MO"] = str(slow_mo)
        elif headless:
            env["RPA_SLOW_MO"] = "0"
        else:
            env["RPA_SLOW_MO"] = "300"

        log_file = log_path.open("w", encoding="utf-8", errors="replace")

        print(f"Worker {worker_id}: perfil={perfil} log={log_path}")

        proceso = subprocess.Popen(
            [sys.executable, "app.py"],
            cwd=Path.cwd(),
            env=env,
            stdout=log_file,
            stderr=subprocess.STDOUT,
            text=True,
        )

        procesos.append((worker_id, proceso, log_file, log_path))

    print()
    print("Procesos lanzados. Para ver logs en otra terminal:")
    print("Get-Content .\\outputs\\parallel_worker_1.log -Wait")
    print("Get-Content .\\outputs\\parallel_worker_2.log -Wait")
    print()

    exit_code_final = 0

    for worker_id, proceso, log_file, log_path in procesos:
        codigo = proceso.wait()
        log_file.close()

        print(f"Worker {worker_id} terminó con código {codigo}. Log: {log_path}")

        if codigo != 0:
            exit_code_final = codigo

    print()
    print("RPA paralelo terminado.")

    if exit_code_final != 0:
        print("Uno o más workers terminaron con error. Revisa los logs.")

    return exit_code_final


if __name__ == "__main__":
    raise SystemExit(main())
