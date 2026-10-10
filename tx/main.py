import argparse
import os
import subprocess
import sys
from pathlib import Path

def run_script(script_path: str):
    script = Path(script_path)
    if not script.exists():
        print(f"[ERROR] No existe: {script_path}")
        return 2
    print(f"\n=== Ejecutando: {script.name} ===")
    return subprocess.call([sys.executable, str(script)])

def main():
    ap = argparse.ArgumentParser(description="Runner de scripts SAP DATALOG")
    ap.add_argument("--run", action="append", default=[], help="Script .py a ejecutar (puedes repetir)")
    ap.add_argument("--week", default=None, help="Semana destino (ej. 09). Alternativa a SAP_DATALOG_SEMANA")
    ap.add_argument("--base-dir", default=None, help="Base dir hasta ...\\2026 (alternativa a SAP_DATALOG_BASE_DIR)")
    args = ap.parse_args()

    if args.week:
        os.environ["SAP_DATALOG_SEMANA"] = args.week
    if args.base_dir:
        os.environ["SAP_DATALOG_BASE_DIR"] = args.base_dir

    if not args.run:
        print("Nada que ejecutar. Usa --run Script.py")
        return 1

    rc = 0
    for s in args.run:
        r = run_script(s)
        if r != 0:
            rc = r
            print(f"[WARN] Script falló con código {r}: {s}")
    return rc

if __name__ == "__main__":
    raise SystemExit(main())