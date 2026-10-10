import json
from pathlib import Path
import pyodbc
from config import DRIVER, DB_CONFIG_FILE

def read_db_config(config_path: Path) -> dict:
    p = Path(config_path)
    if p.exists() and p.is_dir():
        p = p / "db_config.json"
    if not p.exists():
        raise FileNotFoundError(
            f"No existe {p}. Crea db_config.json junto a main.py. "
            'Ejemplo: { "server":"...", "database":"gaeval_v2", "username":"...", "password":"..." }'
        )
    if p.is_dir():
        raise IsADirectoryError(f"La ruta {p} es una carpeta. Debe ser un archivo db_config.json.")

    cfg = json.loads(p.read_text(encoding="utf-8-sig"))
    for k in ["server", "database", "username", "password"]:
        if k not in cfg or not str(cfg[k]).strip():
            raise ValueError(f"Falta '{k}' en {p}.")
    return cfg

def get_connection():
    cfg = read_db_config(DB_CONFIG_FILE)
    conn_str = (
        f"DRIVER={DRIVER};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"
        "Connection Timeout=30;"
    )
    return pyodbc.connect(conn_str)

def col_limit(conn, table: str, column: str):
    sql = '''
    SELECT t.name AS tipo, c.max_length
    FROM sys.columns c
    JOIN sys.types t ON c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID(?) AND c.name = ?;
    '''
    cur = conn.cursor()
    cur.execute(sql, (table, column))
    row = cur.fetchone()
    if not row:
        return None
    tipo, max_len = row[0], row[1]
    if max_len == -1:
        return None
    if str(tipo).lower() == "nvarchar":
        return max_len // 2
    return max_len

