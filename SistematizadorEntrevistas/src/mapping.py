from __future__ import annotations

from typing import Any, Dict

import pandas as pd


def _get(row: pd.Series, *candidates: str) -> str:
    for c in candidates:
        if c in row.index and row[c] is not None:
            v = str(row[c]).strip()
            if v and v.lower() != "nan":
                return v
    return ""


def normalize_newlines_to_literal(s: str) -> str:
    """
    Convierte saltos reales (\r\n o \n) a secuencia literal \\n para que en SQL
    se vea como texto "\\n".
    Si el texto ya trae \\n literal, se mantiene.
    """
    if not s:
        return ""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    # solo convierte saltos reales:
    return s


def map_row_to_respuesta(row: pd.Series, pregunta_id: int, aplicacion_id: int) -> Dict[str, Any]:
    """
    Mapea una fila del DF (ya estandarizado por excel_reader) a columnas reales de dbo.Respuesta:
      - SituacionActual  <= ANALISIS
      - Notas            <= NOTAS
      - RespuestaLiteral <= SITUACION

    (Si tu Excel no usa estas columnas, excel_reader.py debe estandarizarlo.)
    """
    notas = _get(row, "NOTAS", "CONVERSACION_ENTREVISTADOR_1", "Notas", "NOTAS")
    resp_lit = _get(row, "SITUACION", "CONVERSACION_INFORMANTE_1", "RespuestaLiteral", "RESPUESTA_LITERAL", "RESPUESTA")
    situacion = _get(row, "ANALISIS", "RESULTADO", "ANÁLISIS_DE_COBERTURA", "ANALISIS_DE_COBERTURA", "SituacionActual")

    # Reemplazo recomendado (a los 3)
    notas = normalize_newlines_to_literal(notas)
    resp_lit = normalize_newlines_to_literal(resp_lit)
    situacion = normalize_newlines_to_literal(situacion)

    return {
        "AplicacionId": int(aplicacion_id),
        "PreguntaId": int(pregunta_id),
        "SituacionActual": situacion,
        "Notas": notas,
        "RespuestaLiteral": resp_lit,
    }


# Alias compatibilidad con versiones anteriores
def map_row_to_sql(
    row: pd.Series,
    evaluacion_id: int | None = None,
    informante_id: int | None = None,
    aplicacion_id: int | None = None,
    pregunta_id: int | None = None,
) -> Dict[str, Any]:
    if aplicacion_id is None or pregunta_id is None:
        raise ValueError("map_row_to_sql requiere aplicacion_id y pregunta_id")
    return map_row_to_respuesta(row=row, pregunta_id=pregunta_id, aplicacion_id=aplicacion_id)
