from pathlib import Path
import pandas as pd
from config import DEFAULT_SHEET_NAME, COL_PREG_COD, COL_ENTREV1, COL_ENTREV2, COL_INF, COL_RESULT

def leer_excel(excel_paths, sheet_name: str = DEFAULT_SHEET_NAME) -> pd.DataFrame:
    dfs = []
    for p in excel_paths:
        ruta = Path(p)
        if not ruta.exists():
            raise FileNotFoundError(f"No existe el archivo: {ruta}")

        df = pd.read_excel(ruta, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")], errors="ignore")

        requeridas = [COL_PREG_COD, COL_ENTREV1, COL_ENTREV2, COL_INF, COL_RESULT]
        faltan = [c for c in requeridas if c not in df.columns]
        if faltan:
            raise ValueError(f"En {ruta.name} faltan columnas: {faltan}. Columnas: {list(df.columns)}")

        df[COL_PREG_COD] = df[COL_PREG_COD].astype(str).str.strip()

        notas = (
            df[[COL_ENTREV1, COL_ENTREV2]]
            .fillna("")
            .astype(str)
            .agg("\n".join, axis=1)
            .str.replace(r"\n+", "\n", regex=True)
            .str.strip()
        )

        situacion = df[COL_INF].fillna("").astype(str).str.strip()
        resultado = df[COL_RESULT].fillna("").astype(str).str.strip()

        situacion = resultado + "\n\nRESPUESTA LITERAL:\n" + situacion

        out = pd.DataFrame({
            "PREGUNTA_CODIGO": df[COL_PREG_COD],
            "NOTAS": notas,
            "SITUACION": situacion
        })
        dfs.append(out)

    return pd.concat(dfs, ignore_index=True)

