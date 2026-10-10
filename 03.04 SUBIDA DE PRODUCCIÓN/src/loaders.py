from __future__ import annotations
from datetime import datetime
import pandas as pd

from src.db import col_limit
from src.repo import preguntas_respondidas

def insertar_respuestas_solo_faltantes(conn, aplicacion_id: int, df_mapeado: pd.DataFrame, usuario_registro_id: int) -> int:
    existentes = preguntas_respondidas(conn, aplicacion_id)

    df = df_mapeado.copy()
    df["ID_PREGUNTA"] = df["ID_PREGUNTA"].astype(int)
    df = df[~df["ID_PREGUNTA"].isin(existentes)].copy()
    if df.empty:
        return 0

    lim_sit = col_limit(conn, "dbo.Respuesta", "SituacionActual")
    lim_not = col_limit(conn, "dbo.Respuesta", "Notas")

    insert_sql = '''
    INSERT INTO dbo.Respuesta
    (AplicacionId, PreguntaId, SituacionActual, Notas, UsuarioRegistroId, FechaRegistro)
    VALUES (?, ?, ?, ?, ?, ?);
    '''

    now = datetime.now()
    params = []
    for _, r in df.iterrows():
        situ = "" if pd.isna(r["SITUACION"]) else str(r["SITUACION"])
        notas = "" if pd.isna(r["NOTAS"]) else str(r["NOTAS"])

        if lim_sit is not None and len(situ) > lim_sit:
            situ = situ[:lim_sit]
        if lim_not is not None and len(notas) > lim_not:
            notas = notas[:lim_not]

        params.append((int(aplicacion_id), int(r["ID_PREGUNTA"]), situ, notas, int(usuario_registro_id), now))

    cur = conn.cursor()
    cur.fast_executemany = False
    cur.executemany(insert_sql, params)
    conn.commit()
    return len(params)

def merge_concatenar_respuestas(conn, aplicacion_id: int, df_mapeado: pd.DataFrame, usuario_registro_id: int) -> dict:
    for col in ("ID_PREGUNTA", "SITUACION", "NOTAS"):
        if col not in df_mapeado.columns:
            raise ValueError(f"Falta columna {col} en df_mapeado")

    lim_sit = col_limit(conn, "dbo.Respuesta", "SituacionActual")
    lim_not = col_limit(conn, "dbo.Respuesta", "Notas")
    lim_sit_sql = int(lim_sit) if lim_sit is not None else 0
    lim_not_sql = int(lim_not) if lim_not is not None else 0

    df = df_mapeado.copy()
    df["ID_PREGUNTA"] = df["ID_PREGUNTA"].astype(int)
    df["SITUACION"] = df["SITUACION"].fillna("").astype(str).str.strip()
    df["NOTAS"] = df["NOTAS"].fillna("").astype(str).str.strip()

    df = df.groupby("ID_PREGUNTA", as_index=False).agg({
        "SITUACION": lambda s: "\n\n".join([x for x in s.tolist() if x]).strip(),
        "NOTAS":     lambda s: "\n\n".join([x for x in s.tolist() if x]).strip(),
    })

    if lim_sit is not None:
        df["SITUACION"] = df["SITUACION"].str.slice(0, lim_sit)
    if lim_not is not None:
        df["NOTAS"] = df["NOTAS"].str.slice(0, lim_not)

    cur = conn.cursor()
    cur.execute("""
    SET NOCOUNT ON;
    IF OBJECT_ID('tempdb..#SRC_RESP') IS NOT NULL DROP TABLE #SRC_RESP;
    CREATE TABLE #SRC_RESP(
        PreguntaId INT NOT NULL,
        SituacionActual NVARCHAR(MAX) NULL,
        Notas NVARCHAR(MAX) NULL
    );
    """)

    insert_tmp = "INSERT INTO #SRC_RESP(PreguntaId, SituacionActual, Notas) VALUES (?, ?, ?);"
    params = [(int(r.ID_PREGUNTA), r.SITUACION, r.NOTAS) for r in df.itertuples(index=False)]
    cur.fast_executemany = False
    cur.executemany(insert_tmp, params)

    merge_sql = r"""
    SET NOCOUNT ON;
    DECLARE @AplicacionId INT = ?;
    DECLARE @UsuarioRegistroId INT = ?;
    DECLARE @LimSit INT = ?; -- 0 => MAX
    DECLARE @LimNot INT = ?; -- 0 => MAX
    DECLARE @Sep NVARCHAR(10) = CHAR(13)+CHAR(10)+CHAR(13)+CHAR(10);

    IF OBJECT_ID('tempdb..#MERGE_ACTIONS') IS NOT NULL DROP TABLE #MERGE_ACTIONS;
    CREATE TABLE #MERGE_ACTIONS([action] NVARCHAR(10));

    MERGE dbo.Respuesta AS T
    USING #SRC_RESP AS S
      ON T.AplicacionId = @AplicacionId
     AND T.PreguntaId   = S.PreguntaId

    WHEN MATCHED THEN
      UPDATE SET
        SituacionActual =
          CASE
            WHEN @LimSit <= 0 THEN
              CASE
                WHEN S.SituacionActual IS NULL OR LTRIM(RTRIM(S.SituacionActual)) = '' THEN T.SituacionActual
                WHEN T.SituacionActual IS NULL OR LTRIM(RTRIM(T.SituacionActual)) = '' THEN S.SituacionActual
                WHEN CHARINDEX(S.SituacionActual, T.SituacionActual) > 0 THEN T.SituacionActual
                ELSE CONCAT(T.SituacionActual, @Sep, S.SituacionActual)
              END
            ELSE
              LEFT(
                CASE
                  WHEN S.SituacionActual IS NULL OR LTRIM(RTRIM(S.SituacionActual)) = '' THEN T.SituacionActual
                  WHEN T.SituacionActual IS NULL OR LTRIM(RTRIM(T.SituacionActual)) = '' THEN S.SituacionActual
                  WHEN CHARINDEX(S.SituacionActual, T.SituacionActual) > 0 THEN T.SituacionActual
                  ELSE CONCAT(T.SituacionActual, @Sep, S.SituacionActual)
                END,
              @LimSit)
          END,

        Notas =
          CASE
            WHEN @LimNot <= 0 THEN
              CASE
                WHEN S.Notas IS NULL OR LTRIM(RTRIM(S.Notas)) = '' THEN T.Notas
                WHEN T.Notas IS NULL OR LTRIM(RTRIM(T.Notas)) = '' THEN S.Notas
                WHEN CHARINDEX(S.Notas, T.Notas) > 0 THEN T.Notas
                ELSE CONCAT(T.Notas, @Sep, S.Notas)
              END
            ELSE
              LEFT(
                CASE
                  WHEN S.Notas IS NULL OR LTRIM(RTRIM(S.Notas)) = '' THEN T.Notas
                  WHEN T.Notas IS NULL OR LTRIM(RTRIM(T.Notas)) = '' THEN S.Notas
                  WHEN CHARINDEX(S.Notas, T.Notas) > 0 THEN T.Notas
                  ELSE CONCAT(T.Notas, @Sep, S.Notas)
                END,
              @LimNot)
          END

    WHEN NOT MATCHED BY TARGET THEN
      INSERT (AplicacionId, PreguntaId, SituacionActual, Notas, UsuarioRegistroId, FechaRegistro)
      VALUES (@AplicacionId, S.PreguntaId, S.SituacionActual, S.Notas, @UsuarioRegistroId, GETDATE())

    OUTPUT $action INTO #MERGE_ACTIONS;

    SELECT [action], COUNT(*) AS n
    FROM #MERGE_ACTIONS
    GROUP BY [action];
    """

    cur.execute(merge_sql, (int(aplicacion_id), int(usuario_registro_id), lim_sit_sql, lim_not_sql))

    rows = None
    while True:
        if cur.description is not None:
            rows = cur.fetchall()
            break
        if not cur.nextset():
            break

    conn.commit()

    out = {"INSERT": 0, "UPDATE": 0}
    if rows:
        for action, n in rows:
            a = str(action).upper()
            if a in out:
                out[a] = int(n)
    return out

