from __future__ import annotations

from pathlib import Path

import pandas as pd

from .db_questions import load_questions
from .tsv_utils import extract_codeblock, parse_table, rows_to_dataframe


def load_tsv_as_df(tsv_path: str | Path) -> pd.DataFrame:
    p = Path(tsv_path)
    if not p.exists():
        raise FileNotFoundError(f"No existe TSV: {p}")
    raw = p.read_text(encoding="utf-8", errors="ignore")
    table = extract_codeblock(raw)
    rows, _ = parse_table(table)
    df = rows_to_dataframe(rows)
    # Normaliza headers
    df.columns = [str(c).strip() for c in df.columns]
    return df


def merge_tsv_with_questions(
    evaluacion_id: int,
    db_config_path: str | Path,
    tsv_path: str | Path,
) -> pd.DataFrame:
    df_gpt = load_tsv_as_df(tsv_path)

    if "PREGUNTA_CODIGO" not in df_gpt.columns:
        raise ValueError(
            f"El TSV no tiene columna PREGUNTA_CODIGO. Columnas: {list(df_gpt.columns)}"
        )

    df_gpt["PREGUNTA_CODIGO"] = df_gpt["PREGUNTA_CODIGO"].astype(str).str.strip()

    df_q = load_questions(evaluacion_id=evaluacion_id, db_config_path=db_config_path)
    df_q["PREGUNTA_CODIGO"] = df_q["PREGUNTA_CODIGO"].astype(str).str.strip()

    # Merge: mantenemos TODAS las preguntas del cuestionario
    df = df_q.merge(df_gpt, on="PREGUNTA_CODIGO", how="left", suffixes=("_SQL", ""))

    # Si GPT trae PREGUNTA_DESCRIPCION o PREGUNTA_SITUACION_DESEADA, preferimos SQL,
    # pero rellenamos si SQL viene vacío.
    for col in ["PREGUNTA_DESCRIPCION", "PREGUNTA_SITUACION_DESEADA"]:
        c_sql = f"{col}_SQL"
        if c_sql in df.columns and col in df.columns:
            df[col] = df[c_sql].replace({None: ""}).astype(str)
            # rellena con GPT si SQL está vacío
            mask = df[col].str.strip().eq("")
            df.loc[mask, col] = df.loc[mask, col].replace({None: ""}).astype(str)
            df = df.drop(columns=[c_sql])
        elif c_sql in df.columns and col not in df.columns:
            df = df.rename(columns={c_sql: col})

    return df
