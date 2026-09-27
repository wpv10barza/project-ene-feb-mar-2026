from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict
import pyodbc

# Intenta importar el DRIVER de config, si no existe usa el default 17
try:
    from config import DRIVER
except ImportError:
    DRIVER = "{ODBC Driver 17 for SQL Server}"

def read_db_config(config_path: str | Path) -> Dict[str, Any]:
    """
    Lee db_config.json con tolerancia a BOM (utf-8-sig) y valida campos mínimos.
    """
    p = Path(config_path)

    if p.exists() and p.is_dir():
        p = p / "db_config.json"

    if not p.exists():
        raise FileNotFoundError(f"No existe db_config.json en: {p}")

    # Lectura segura con utf-8-sig para eliminar BOM si existe
    try:
        content = p.read_text(encoding="utf-8-sig")
        data = json.loads(content)
    except json.JSONDecodeError as e:
        raise ValueError(f"Error al decodificar JSON en {p}: {e}")

    # Validaciones mínimas
    if not str(data.get("server", "")).strip():
        raise ValueError(f"Falta 'server' en {p}")
    if not str(data.get("database", "")).strip():
        raise ValueError(f"Falta 'database' en {p}")

    return data


def build_conn_str(cfg: Dict[str, Any]) -> str:
    """
    Construye la cadena de conexión dinámicamente.
    Soporta Azure (SQL Auth) y Local (Windows Auth) solo cambiando el JSON.
    """
    server = str(cfg.get("server", "")).strip()
    database = str(cfg.get("database", "")).strip()
    username = str(cfg.get("username", "")).strip()
    password = str(cfg.get("password", "")).strip()
    timeout = int(cfg.get("timeout", 30))
    
    # Detección automática de entorno para defaults inteligentes
    is_azure = "database.windows.net" in server.lower()

    # --- Lógica de Encriptación ---
    # Si el JSON tiene 'encrypt', lo usa. Si no: Azure=yes, Local=no (para evitar error SSL)
    encrypt_default = "yes" if is_azure else "no"
    encrypt = str(cfg.get("encrypt", encrypt_default)).lower()

    # --- Lógica de Certificados ---
    # Si es Azure, confiamos en el certificado público (Trust=no).
    # Si es Local, solemos necesitar confiar en certificados autofirmados (Trust=yes).
    trust_default = "no" if is_azure else "yes"
    trust = str(cfg.get("trust_server_certificate", trust_default)).lower()

    parts = [
        f"DRIVER={DRIVER}",
        f"SERVER={server}",
        f"DATABASE={database}",
        f"Connection Timeout={timeout}",
    ]

    # --- Lógica de Autenticación ---
    if username and password:
        # Autenticación SQL (Típico Azure o usuario SA local)
        parts.append(f"UID={username}")
        parts.append(f"PWD={password}")
    else:
        # Autenticación Windows (Típico Local / Dominio Corporativo)
        parts.append("Trusted_Connection=yes")

    # Parámetros de Seguridad
    parts.append(f"Encrypt={encrypt}")
    parts.append(f"TrustServerCertificate={trust}")

    return ";".join(parts) + ";"


def connect(cfg: Dict[str, Any]) -> pyodbc.Connection:
    conn_str = build_conn_str(cfg)
    return pyodbc.connect(conn_str)


def connect_from_path(db_config_path: str | Path) -> pyodbc.Connection:
    cfg = read_db_config(db_config_path)
    return connect(cfg)