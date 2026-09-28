from __future__ import annotations

import argparse
from pathlib import Path
import os

# Importamos configuración
from config import (
    RESULT_RAW_TXT,
    RESULT_TSV,
    RESULT_CSV,
    RESULT_XLSX,
    MERGED_TSV,
    MERGED_CSV,
    MERGED_XLSX,
    DEFAULT_DB_CONFIG_PATH,
    DEFAULT_TARGET_TABLE,
    DEFAULT_USUARIO_REGISTRO_ID,
    DEFAULT_SHEET_NAME,
    DATA_DIR,
    
    # Rutas Hardcoded
    PATH_PROMPT_TXT,
    PATH_GUIA_PREGUNTAS_FALLBACK
)

# Importamos módulos
from src.selenium_runner import run_chatgpt_via_edge
from src.tsv_utils import extract_codeblock, parse_table, write_xlsx
from src.merge_runner import merge_tsv_with_questions
from src.loaders import load_excel_to_sql
from src.exporter import export_taxonomia_excel

def cmd_preview(args: argparse.Namespace) -> int:
    """
    1. Lee Prompt Fijo (Config).
    2. Toma Transcripción Variable (CLI).
    3. Genera/Toma Excel (CLI o Config).
    4. Ejecuta Selenium.
    """
    print("--- INICIANDO PROCESO PREVIEW ---")
    
    # 1. LEER PROMPT FIJO
    print(f"[1] Leyendo Prompt Maestro desde: {PATH_PROMPT_TXT}")
    if not os.path.exists(PATH_PROMPT_TXT):
        print(f"[FATAL] No existe el archivo de prompt: {PATH_PROMPT_TXT}")
        return 1
    try:
        prompt_text = Path(PATH_PROMPT_TXT).read_text(encoding="utf-8")
    except Exception as e:
        print(f"[FATAL] Error leyendo prompt: {e}")
        return 1

    # 2. OBTENER TRANSCRIPCIÓN (VARIABLE)
    transcripcion_path = args.transcripcion
    if not transcripcion_path:
        print("[FATAL] Debes proporcionar --transcripcion \"RUTA_ARCHIVO\"")
        return 1
    
    # Limpiar comillas si el usuario las puso dobles en el argumento
    transcripcion_path = transcripcion_path.strip('"').strip("'")
    
    print(f"[2] Transcripción seleccionada: {transcripcion_path}")
    if not os.path.exists(transcripcion_path):
        print(f"[FATAL] No se encuentra el archivo de transcripción.")
        return 1

    # 3. OBTENER EXCEL (Exportado o Fallback)
    excel_path = None
    
    if args.export_taxonomia:
        # Opción A: Generar desde SQL
        print(f"[3] Exportando taxonomía para Evaluación ID: {args.evaluacion_id}...")
        try:
            excel_path = export_taxonomia_excel(
                evaluacion_id=args.evaluacion_id,
                db_config_path=Path(args.db_config),
                output_dir=DATA_DIR
            )
            print(f"    -> Exportado a: {excel_path}")
        except Exception as e:
            print(f"[ERROR] Falló la exportación SQL: {e}")
            return 1
    else:
        # Opción B: Usar archivo fijo
        print(f"[3] Usando Excel Hardcoded (sin exportación SQL)...")
        excel_path = PATH_GUIA_PREGUNTAS_FALLBACK
        if not os.path.exists(excel_path):
            print(f"[WARN] No existe el Excel de respaldo: {excel_path}")
            # No retornamos error fatal, quizás el usuario solo quiere subir la transcripción
            # pero el prompt dice que necesita ambos. Advertimos.

    # 4. PREPARAR LISTA DE ARCHIVOS
    files_to_upload = [transcripcion_path]
    if excel_path and os.path.exists(excel_path):
        files_to_upload.append(str(excel_path))

    # 5. EJECUTAR SELENIUM
    if args.run_selenium:
        print("[4] Iniciando Selenium...")
        res = run_chatgpt_via_edge(
            prompt_text=prompt_text,
            wait_seconds=args.wait_seconds,
            sheets_url=args.sheets_url,
            headless=args.headless,
            driver_path=args.driver_path,
            file_paths=files_to_upload
        )

        if not res.ok:
            print(f"[ERROR] Selenium falló: {res.message}")
            return 2

        # Procesar salida
        raw = res.tsv_text
        RESULT_RAW_TXT.write_text(raw, encoding="utf-8")
        table = extract_codeblock(raw)
        RESULT_TSV.write_text(table, encoding="utf-8")
        
        rows, _ = parse_table(table)
        if rows:
            import csv
            RESULT_CSV.parent.mkdir(parents=True, exist_ok=True)
            with RESULT_CSV.open("w", encoding="utf-8", newline="") as f:
                csv.writer(f).writerows(rows)
            write_xlsx(rows, RESULT_XLSX)
            print(f"[OK] Archivos generados en carpeta data/.")
    else:
        print("[INFO] --run-selenium no detectado. Solo se prepararon los paths.")

    return 0

# --- EL RESTO DE FUNCIONES (cmd_run, cmd_load) SE MANTIENE IGUAL ---
def cmd_run(args: argparse.Namespace) -> int:
    tsv_in = Path(args.input_tsv).resolve()
    df = merge_tsv_with_questions(
        evaluacion_id=args.evaluacion_id,
        db_config_path=args.db_config,
        tsv_path=tsv_in,
    )
    MERGED_TSV.write_text(df.to_csv(sep="\t", index=False), encoding="utf-8")
    df.to_csv(MERGED_CSV, index=False, encoding="utf-8")
    try:
        df.to_excel(MERGED_XLSX, index=False)
    except: pass
    print(f"[OK] Matriz unificada generada: {MERGED_XLSX}")
    return 0

def cmd_load(args: argparse.Namespace) -> int:
    rows = load_excel_to_sql(
        evaluacion_id=args.evaluacion_id,
        informante_id=args.informante_id,
        excel_path=args.excel,
        db_config_path=args.db_config,
        table=args.table,
        usuario_registro_id=args.usuario_registro_id,
        sheet_name=args.sheet_name,
    )
    return 0 if rows >= 0 else 1

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="p13_sistematizador")
    sub = p.add_subparsers(dest="cmd", required=True)

    # --- PREVIEW ---
    p_preview = sub.add_parser("preview")
    
    # ARGUMENTO OBLIGATORIO: Transcripción
    p_preview.add_argument("--transcripcion", required=True, help="Ruta al archivo de transcripción")
    
    # Opcionales
    p_preview.add_argument("--evaluacion-id", type=int, default=0)
    p_preview.add_argument("--db-config", default=str(DEFAULT_DB_CONFIG_PATH))
    p_preview.add_argument("--export-taxonomia", action="store_true", help="Generar Excel desde SQL")
    p_preview.add_argument("--run-selenium", action="store_true", help="Ejecutar navegador")
    
    # Selenium configs
    p_preview.add_argument("--wait-seconds", type=int, default=120)
    p_preview.add_argument("--sheets-url", default=None)
    p_preview.add_argument("--headless", action="store_true")
    p_preview.add_argument("--driver-path", default=None)
    p_preview.add_argument("--prompt-in", help="(Ignorado, usa config)")
    p_preview.add_argument("--prompt-out", help="(Ignorado)")
    p_preview.add_argument("--input-tsv", default=str(RESULT_TSV))

    p_preview.set_defaults(func=cmd_preview)

    # --- RUN ---
    p_run = sub.add_parser("run")
    p_run.add_argument("--evaluacion-id", type=int, required=True)
    p_run.add_argument("--db-config", default=str(DEFAULT_DB_CONFIG_PATH))
    p_run.add_argument("--input-tsv", default=str(RESULT_TSV))
    p_run.set_defaults(func=cmd_run)

    # --- LOAD ---
    p_load = sub.add_parser("load")
    p_load.add_argument("--evaluacion-id", type=int, required=True)
    p_load.add_argument("--informante-id", type=int, required=True)
    p_load.add_argument("--excel", required=True)
    p_load.add_argument("--sheet-name", default=DEFAULT_SHEET_NAME)
    p_load.add_argument("--db-config", default=str(DEFAULT_DB_CONFIG_PATH))
    p_load.add_argument("--table", default=DEFAULT_TARGET_TABLE)
    p_load.add_argument("--usuario-registro-id", type=int, default=DEFAULT_USUARIO_REGISTRO_ID)
    p_load.set_defaults(func=cmd_load)

    return p

def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    try:
        status = args.func(args)
        if status != 0:
            raise SystemExit(status)
    except Exception as e:
        print(f"[FATAL] Error en la ejecución: {e}")
        raise SystemExit(1)

if __name__ == "__main__":
    main()


    