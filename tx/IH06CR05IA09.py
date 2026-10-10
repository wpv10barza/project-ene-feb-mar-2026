import os
import time
from datetime import datetime
from pathlib import Path, PureWindowsPath
import re

# ============================================================================
# CONFIGURACIÓN (AUTO) — Semana / Carpeta / Fecha (para nombres con ddmmaa)
# - Puedes sobreescribir con variables de entorno:
#   SAP_DATALOG_BASE_DIR, SAP_DATALOG_SEMANA, SAP_DATALOG_SEMANA_ORIGEN
# ============================================================================
BASE_DIR = os.getenv("SAP_DATALOG_BASE_DIR", r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026")
SEMANA_DESTINO = os.getenv("SAP_DATALOG_SEMANA", "09")  # ejemplo: "09"
try:
    _w = int(SEMANA_DESTINO)
    _prev = str(_w-1).zfill(2)
except:
    _prev = "08"
SEMANA_ORIGEN = os.getenv("SAP_DATALOG_SEMANA_ORIGEN", _prev)
WEEK_DIR = os.path.join(BASE_DIR, SEMANA_DESTINO)
FECHA_HOY = datetime.now().strftime("%d%m%y")

def nombre_con_fecha(nombre_base: str, ext: str = "xlsx") -> str:
    """Devuelve nombre_base con sufijo _ddmmaa (reemplaza si ya existe)."""
    base = re.sub(r"_\d{6}$", "", nombre_base)
    return f"{base}_{FECHA_HOY}.{ext}"

def join_week(*parts: str) -> str:
    ruta = os.path.join(WEEK_DIR, *parts)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    return ruta


import pandas as pd
import win32com.client


# =========================
# CONFIG SEMANA 07
# =========================
R07 = Path(r"WEEK_DIR")
R06 = R07.parent / "06"
CONTROL_XLSX = None  # si quieres fijarlo: r"C:\...\2026\07\libro1.xlsx"

# IW39 (según JM5)
PLANNER_GROUP = "qi3pp"
DATE_FROM = "20.12.2025"
DATE_TO = "31.12.9999"
DATE_END = "31.12.9999"
RESPONSABLE = "MCARRANZA03"

KEY_FOLDER = "IW39_CabeceraOrden"   # clave para ubicar carpeta en Libro1
BASENAME = "IW39_CabeceraOrden"     # nombre base archivo


# =========================
# Router (Libro1)
# =========================
def _control_path() -> Path:
    if CONTROL_XLSX:
        p = Path(CONTROL_XLSX)
        if not p.exists():
            raise FileNotFoundError(f"CONTROL_XLSX no existe: {p}")
        return p

    candidates = [
        R07 / "libro1.xlsx",
        R07 / "- CONTROL DE ARCHIVOS.xlsx",
        R07 / "CONTROL DE ARCHIVOS.xlsx",
        R06 / "libro1.xlsx",
        R06 / "- CONTROL DE ARCHIVOS.xlsx",
        R06 / "CONTROL DE ARCHIVOS.xlsx",
        Path(__file__).with_name("libro1.xlsx"),
        Path(__file__).with_name("- CONTROL DE ARCHIVOS.xlsx"),
        Path(__file__).with_name("CONTROL DE ARCHIVOS.xlsx"),
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError("No se encontró Libro1. Copia 'libro1.xlsx' a la raíz de semana 07 o define CONTROL_XLSX.")

def _to07_from_any_path(folder: str) -> PureWindowsPath:
    base06 = PureWindowsPath(str(R06))
    base07 = PureWindowsPath(str(R07))
    f = PureWindowsPath(str(folder).rstrip("\\/"))

    try:
        return base07 / f.relative_to(base06)
    except Exception:
        pass

    s = str(f)
    mark = r"\2026\06\\"
    i = s.lower().find(mark.lower())
    if i != -1:
        rel = s[i + len(mark):]
        return base07 / (PureWindowsPath(rel) if rel else PureWindowsPath())

    return base07

def _mkdir(p: PureWindowsPath) -> str:
    os.makedirs(str(p), exist_ok=True)
    return str(p)

def _split_archivos(cell) -> list[str]:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    s = str(cell).strip()
    if not s:
        return []
    parts = []
    for chunk in s.replace("\n", ";").split(";"):
        t = chunk.strip()
        if t:
            parts.append(t)
    return parts

class Router:
    def __init__(self, mapping: dict[str, str]):
        self.mapping = mapping  # key: filename(lower) or substring -> folderPath or tipo

    def dest_folder(self, key: str) -> str:
        k = key.strip().lower()
        val = self.mapping.get(k)

        if val is None:
            # substring match
            for kk, vv in self.mapping.items():
                if kk and kk in k:
                    val = vv
                    break

        # si sigue None -> raíz 07
        if not val or (isinstance(val, float) and pd.isna(val)):
            return _mkdir(PureWindowsPath(str(R07)))

        s = str(val).strip()

        # ruta absoluta
        if (":\\" in s) or s.startswith("\\\\"):
            dest = _to07_from_any_path(s)
            return _mkdir(dest)

        # tipo/carpeta relativa
        dest = PureWindowsPath(str(R07)) / PureWindowsPath(s)
        return _mkdir(dest)

def load_router() -> Router:
    xlsx = _control_path()
    xls = pd.ExcelFile(xlsx, engine="openpyxl")

    # 1) Preferir hoja con Name/Folder Path
    for sh in xls.sheet_names:
        head = pd.read_excel(xlsx, sheet_name=sh, nrows=0, engine="openpyxl")
        cols = [str(c).strip() for c in head.columns]
        low = [c.lower() for c in cols]
        if "name" in low and "folder path" in low:
            df = pd.read_excel(xlsx, sheet_name=sh, engine="openpyxl").copy()
            df["k"] = df["Name"].astype(str).str.strip().str.lower()
            df["m"] = pd.to_datetime(df.get("Date modified", pd.NaT), errors="coerce")
            df = df.sort_values("m", ascending=False)
            mp = dict(zip(df["k"], df["Folder Path"].astype(str)))
            return Router(mp)

    # 2) Fallback: hoja con Tipo/Archivos
    for sh in xls.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sh, engine="openpyxl")
        low = [str(c).strip().lower() for c in df.columns]
        if "tipo" in low and "archivos" in low:
            mp = {}
            for _, row in df.iterrows():
                tipo = str(row.get("Tipo", "")).strip()
                for fn in _split_archivos(row.get("Archivos", "")):
                    base = fn.strip().lower()
                    mp[base] = tipo
                    # también mapear sin extensión
                    if "." in base:
                        mp[base.rsplit(".", 1)[0]] = tipo
            return Router(mp)

    raise ValueError("Libro1 no tiene (Name/Folder Path) ni (Tipo/Archivos).")

def sap_path(p: str) -> str:
    return p if p.endswith("\\") else p + "\\"


# =========================
# SAP helpers
# =========================
def conectar_sap():
    app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    return app.Children(0).Children(0)

def try_press(session, id_):
    try:
        session.findById(id_, False).press()
        return True
    except Exception:
        return False

def manejar_popups(session):
    # acepta mensajes/seguridad que aparezcan como wnd[1]
    for _ in range(3):
        time.sleep(0.7)
        try:
            if session.findById("wnd[1]", False):
                # primero intentar OK estándar
                if try_press(session, "wnd[1]/tbar[0]/btn[0]"):
                    continue
                # botones típicos de popup SAP
                if try_press(session, "wnd[1]/usr/btnSPOP-OPTION1"):
                    continue
                if try_press(session, "wnd[1]/usr/btn[0]"):
                    continue
        except Exception:
            pass

def export_save_dialog_name_and_folder(session, folder: str, name_no_ext: str):
    # patrón: btn[16] abre diálogo con txt + botón cambiar carpeta
    session.findById("wnd[1]/usr/ssub/1/txt").text = name_no_ext
    try:
        session.findById("wnd[1]/usr/ssub/1/txt").caretPosition = len(name_no_ext)
    except Exception:
        pass
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = folder
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    manejar_popups(session)


# =========================
# MAIN
# =========================
def main():
    router = load_router()

    ddmmyy = datetime.now().strftime("%d%m%y")
    filename_no_ext = f"{BASENAME}_{ddmmyy}"
    out_folder = router.dest_folder(KEY_FOLDER)

    print(f"📁 Carpeta: {out_folder}")
    print(f"🧾 Archivo: {filename_no_ext}.xlsx")

    session = conectar_sap()
    session.findById("wnd[0]").maximize()

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    session.findById("wnd[0]/usr/ctxt[13]").text = PLANNER_GROUP
    session.findById("wnd[0]/usr/ctxt[15]").text = ""
    session.findById("wnd[0]/usr/ctxt[16]").text = DATE_END
    session.findById("wnd[0]/usr/ctxt[57]").text = ""
    session.findById("wnd[0]/usr/ctxt[75]").text = DATE_FROM
    session.findById("wnd[0]/usr/ctxt[76]").text = DATE_TO
    session.findById("wnd[0]/usr/ctxt[131]").text = RESPONSABLE

    session.findById("wnd[0]/tbar[1]/btn[8]").press()   # ejecutar
    time.sleep(3)

    session.findById("wnd[0]/tbar[1]/btn[16]").press()  # export
    time.sleep(1)

    export_save_dialog_name_and_folder(session, sap_path(out_folder), filename_no_ext)

    print("✅ IW39 exportado.")

if __name__ == "__main__":
    main()