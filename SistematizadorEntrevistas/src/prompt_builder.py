from __future__ import annotations
from pathlib import Path
from datetime import datetime
import json
import pandas as pd
from .db_questions import load_questions

def build_complete_prompt(prompt_in_path: Path, evaluacion_id: int | None = None, transcripcion_path: Path | None = None, db_config_path: Path | None = None) -> str:
    if not prompt_in_path.exists(): raise FileNotFoundError(f"No existe: {prompt_in_path}")
    base_text = prompt_in_path.read_text(encoding="utf-8", errors="ignore")
    
    preguntas_text = ""
    if evaluacion_id and db_config_path:
        df_preg = load_questions(evaluacion_id, db_config_path)
        for _, row in df_preg.iterrows():
            preguntas_text += f"COD: {row['PREGUNTA_CODIGO']} | DESC: {row['PREGUNTA_DESCRIPCION']}\n"

    trans_text = ""
    if transcripcion_path and transcripcion_path.exists():
        trans_text = transcripcion_path.read_text(encoding="utf-8", errors="ignore")

    return f"{base_text}\n\n=== PREGUNTAS ===\n{preguntas_text}\n\n=== TRANSCRIPCION ===\n{trans_text}"

def write_prompt_trabajo(text: str, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return out_path

def write_prompt_metadata(p_in: Path, p_out: Path, eval_id: int = None):
    meta = {"date": datetime.now().isoformat(), "eval_id": eval_id}
    p_out.with_suffix(".meta.json").write_text(json.dumps(meta), encoding="utf-8")