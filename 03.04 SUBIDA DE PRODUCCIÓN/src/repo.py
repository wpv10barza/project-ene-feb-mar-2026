import pandas as pd
from src.sql import SQL_PREGUNTAS, SQL_APLICACION, SQL_EXISTING_RESP_COUNT, SQL_EXISTING_PREGUNTAS

def cargar_preguntas(conn, evaluacion_id: int) -> pd.DataFrame:
    return pd.read_sql(SQL_PREGUNTAS, conn, params=[evaluacion_id])

def obtener_aplicacion_id(conn, evaluacion_id: int, informante_id: int) -> int:
    row = pd.read_sql(SQL_APLICACION, conn, params=[evaluacion_id, informante_id])
    if row.empty:
        raise RuntimeError(f"No existe Aplicacion para EvaluacionId={evaluacion_id} e InformanteId={informante_id}.")
    return int(row.iloc[0]["Id"])

def contar_respuestas(conn, aplicacion_id: int) -> int:
    row = pd.read_sql(SQL_EXISTING_RESP_COUNT, conn, params=[aplicacion_id])
    return int(row.iloc[0]["n"]) if not row.empty else 0

def preguntas_respondidas(conn, aplicacion_id: int) -> set[int]:
    row = pd.read_sql(SQL_EXISTING_PREGUNTAS, conn, params=[aplicacion_id])
    if row.empty:
        return set()
    return set(int(x) for x in row["PreguntaId"].tolist())

