import pandas as pd

def validar_mapeo(df_excel: pd.DataFrame, df_preguntas: pd.DataFrame) -> pd.DataFrame:
    df_p = df_preguntas.copy()
    df_p["PREGUNTA_CODIGO"] = df_p["PREGUNTA_CODIGO"].astype(str).str.strip()
    map_id = df_p.set_index("PREGUNTA_CODIGO")["ID_PREGUNTA"].to_dict()

    df = df_excel.copy()
    df["PREGUNTA_CODIGO"] = df["PREGUNTA_CODIGO"].astype(str).str.strip()
    df["ID_PREGUNTA"] = df["PREGUNTA_CODIGO"].map(map_id)

    faltantes = df[df["ID_PREGUNTA"].isna()]["PREGUNTA_CODIGO"].value_counts()
    if not faltantes.empty:
        raise RuntimeError(
            "Detén la carga: hay PREGUNTA_CODIGO sin mapeo en VW_PREGUNTA_POR_EVALUACION.\n"
            + faltantes.to_string()
        )

    return df

