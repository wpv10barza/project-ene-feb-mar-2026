from __future__ import annotations
from pathlib import Path
from typing import List, Optional
import pandas as pd
from config import DEFAULT_SHEET_NAME

def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    drop_cols = [c for c in df.columns if c.lower().startswith("unnamed")]
    if drop_cols:
        df = df.drop(columns=drop_cols, errors="ignore")
    return df

def _is_empty(val: object) -> bool:
    if val is None: return True
    s = str(val).strip()
    return (not s) or (s.lower() == "nan") or (s.lower() == "none")

def agrupar_columnas(row: pd.Series, columnas: List[str]) -> str:
    textos: List[str] = []
    for c in columnas:
        if c in row.index:
            v = row[c]
            if not _is_empty(v):
                textos.append(str(v).strip())
    return "\n\n".join(textos)

def read_matrix(path: str | Path, sheet_name: str = DEFAULT_SHEET_NAME) -> pd.DataFrame:
    p = Path(path).resolve()
    
    # 1. Detección de cabecera saltando filas vacías o basura
    df_temp = pd.read_excel(p, sheet_name=sheet_name, header=None)
    skip_rows = 0
    for i, row in df_temp.iterrows():
        row_str = [str(cell).strip().upper() for cell in row]
        if "PREGUNTA_CODIGO" in row_str:
            skip_rows = i
            break
    
    # 2. Carga real del DataFrame
    df = pd.read_excel(p, sheet_name=sheet_name, skiprows=skip_rows)
    df = _clean_columns(df)

    # 3. Identificación inteligente de columnas
    col_codigo = next((c for c in df.columns if c.upper() == "PREGUNTA_CODIGO"), None)
    
    # Buscar columnas de Entrevistadores (E. o .A)
    cols_entrev = [c for c in df.columns if c.upper().startswith("E.")]
    #if not cols_entrev:
    #    cols_entrev = [c for c in df.columns if c.upper().endswith("1.")]
        
    # Buscar columnas de Informantes (I. o .B)
    cols_info = [c for c in df.columns if c.upper().startswith("I.")]
    #if not cols_info:
     #  cols_info = [c for c in df.columns if c.upper().endswith("2.")]
        
    # Buscar columna de Análisis
    col_analisis = next((c for c in df.columns if "ANALISIS_DE_COBERTURA" in c.upper()), None)

    # --- DEPURACIÓN SEGURA (Sin error de longitud) ---
    print(f"\n[DEBUG] Columnas detectadas en {p.name}:")
    print(f" > Código: {col_codigo}")
    print(f" > Entrevistadores: {cols_entrev}")
    print(f" > Informantes: {cols_info}")
    print(f" > Análisis: {col_analisis}\n")

    if not col_codigo:
        raise ValueError(f"No se encontró PREGUNTA_CODIGO en {p.name}. Columnas: {list(df.columns)}")

    # 4. Construcción del DataFrame estandarizado para el mapper
    return pd.DataFrame({
        "PREGUNTA_CODIGO": df[col_codigo].astype(str).str.strip(),
        "NOTAS": df.apply(lambda r: agrupar_columnas(r, cols_entrev), axis=1),
        "SITUACION": df.apply(lambda r: agrupar_columnas(r, cols_info), axis=1),
        "ANALISIS": df[col_analisis].fillna("").astype(str).str.strip() if col_analisis else ""
    })

def read_excel_matrix(excel_path: Path, sheet_name: str = DEFAULT_SHEET_NAME) -> pd.DataFrame:
    return read_matrix(excel_path, sheet_name=sheet_name)
