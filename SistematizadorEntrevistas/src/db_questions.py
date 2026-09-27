from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd

from .db import connect_from_path

# Query canónica (tu vista usualmente expone estos nombres)
SQL_CANON = """
SELECT
  ID_PREGUNTA,
  PREGUNTA_CODIGO,
  PREGUNTA_DESCRIPCION,
  PREGUNTA_SITUACION_DESEADA
FROM dbo.VW_PREGUNTA_POR_EVALUACION
WHERE EvaluacionId = ?
ORDER BY PREGUNTA_CODIGO;
"""


def _fetch_df(conn, sql: str, params: tuple) -> pd.DataFrame:
    cur = conn.cursor()
    cur.execute(sql, params)
    cols = [d[0] for d in cur.description]
    rows = cur.fetchall()
    return pd.DataFrame.from_records(rows, columns=cols)


def _infer_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nombres de columnas de la vista a los canónicos."""
    cols = {c.lower(): c for c in df.columns}

    def pick(*needles: str) -> Optional[str]:
        for n in needles:
            for lc, orig in cols.items():
                if n in lc:
                    return orig
        return None

    c_id = pick("id_preg", "idpreg", "preguntaid", "id_pregunta", "idpregunta", "id")
    c_cod = pick("codigo", "pregunta_codigo", "pregunta cod")
    c_desc = pick("descripcion", "descripción")
    c_sit = pick("situacion_deseada", "situación_deseada", "situacion deseada", "situación deseada", "deseada")

    if not c_id or not c_cod:
        raise RuntimeError(
            "No pude inferir columnas ID/CODIGO en VW_PREGUNTA_POR_EVALUACION. " 
            f"Columnas vistas: {list(df.columns)}"
        )

    out = pd.DataFrame({
        "ID_PREGUNTA": df[c_id],
        "PREGUNTA_CODIGO": df[c_cod],
        "PREGUNTA_DESCRIPCION": df[c_desc] if c_desc else "",
        "PREGUNTA_SITUACION_DESEADA": df[c_sit] if c_sit else "",
    })
    return out


def load_questions(evaluacion_id: int, db_config_path: str | Path) -> pd.DataFrame:
    """Carga preguntas desde dbo.VW_PREGUNTA_POR_EVALUACION para un EvaluacionId."""
    with connect_from_path(db_config_path) as conn:
        # 1) intento canónico
        try:
            df = _fetch_df(conn, SQL_CANON, (evaluacion_id,))
            return _infer_columns(df)
        except Exception:
            # 2) fallback: trae todo y trata de inferir
            df = _fetch_df(
                conn,
                "SELECT TOP (100000) * FROM dbo.VW_PREGUNTA_POR_EVALUACION WHERE EvaluacionId = ?;",
                (evaluacion_id,),
            )
            return _infer_columns(df)
