from __future__ import annotations
from pathlib import Path
import pandas as pd
from .db import connect_from_path  # Asumimos que db.py existe en src, igual que en repo.py

def export_taxonomia_excel(
    evaluacion_id: int,
    db_config_path: Path,
    output_dir: Path
) -> Path:
    """
    1. Consulta SQL filtrando por EvaluacionId.
    2. Genera Excel: taxonomia_evaluacion_{id}.xlsx
    """
    
    sql = """
    SELECT 
        PREGUNTA_CODIGO, 
        PREGUNTA_DESCRIPCION, 
        PREGUNTA_SITUACION_DESEADA
    FROM dbo.VW_PREGUNTA_POR_EVALUACION
    WHERE EvaluacionId = ?
    ORDER BY PREGUNTA_CODIGO
    """
    
    # Usamos tu infraestructura existente de conexión
    with connect_from_path(db_config_path) as conn:
        df = pd.read_sql(sql, conn, params=[evaluacion_id])

    # Limpieza básica
    df = df.fillna("")
    
    # Definir nombre y ruta
    filename = f"taxonomia_evaluacion_{evaluacion_id}.xlsx"
    out_path = output_dir / filename
    
    # Exportar con formato bonito
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Preguntas")
        
        # Ajustar ancho de columnas automáticamente
        ws = writer.sheets["Preguntas"]
        ws.column_dimensions["A"].width = 15  # 
        ws.column_dimensions["B"].width = 50  # DESC
        ws.column_dimensions["C"].width = 60  # SITUACION
        
    print(f"[INFO] Taxonomía exportada a: {out_path}")
    return out_path