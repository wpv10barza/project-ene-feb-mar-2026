from __future__ import annotations

from pathlib import Path

from .db import connect_from_path


def get_aplicacion_id(evaluacion_id: int, informante_id: int, db_config_path: str | Path) -> int:
    """Obtiene el AplicacionId más reciente para (EvaluacionId, InformanteId)."""
    with connect_from_path(db_config_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT TOP (1) Id
            FROM dbo.Aplicacion
            WHERE EvaluacionId = ? AND InformanteId = ?
            ORDER BY Id DESC;
            """,
            (evaluacion_id, informante_id),
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(
                f"No existe dbo.Aplicacion para EvaluacionId={evaluacion_id} e InformanteId={informante_id}. " 
                "Crea primero la aplicación en el sistema o inserta un registro en dbo.Aplicacion."
            )
        return int(row[0])
