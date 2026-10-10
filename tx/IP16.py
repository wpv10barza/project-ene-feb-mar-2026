from datetime import datetime
import win32com.client
import pandas as pd
import time
import sys
import os
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


# ============================================================================
# CONFIGURACIÓN MAESTRA DE RUTAS
# ============================================================================
BASE_DIR = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026"
SEMANA_ORIGEN = "06"
SEMANA_DESTINO = "07"
EXCEL_MAESTRO_RUTAS = os.path.join(BASE_DIR, SEMANA_ORIGEN, "libro1.xlsx")

def resolver_ruta_dinamica(nombre_archivo, carpeta_forzada=None):
    ruta_destino_base = os.path.join(BASE_DIR, SEMANA_DESTINO)
    if carpeta_forzada:
        ruta_final = os.path.join(ruta_destino_base, carpeta_forzada)
    else:
        try:
            df = pd.read_excel(EXCEL_MAESTRO_RUTAS, sheet_name=SEMANA_ORIGEN)
            fila = df[df['Name'].str.contains(nombre_archivo, case=False, na=False)]
            if not fila.empty:
                ruta_06 = str(fila.iloc[0]['Folder Path']).strip()
                ruta_final = ruta_06.replace(f"\\{SEMANA_ORIGEN}\\", f"\\{SEMANA_DESTINO}\\")
            else:
                ruta_final = ruta_destino_base
        except:
            ruta_final = ruta_destino_base
    os.makedirs(ruta_final, exist_ok=True)
    return ruta_final

def manejar_seguridad(session):
    time.sleep(1)
    try:
        if session.findById("wnd[1]", False): session.findById("wnd[1]/tbar[0]/btn[0]").press() 
    except: pass

def exportar_alv(session, nombre_archivo):
    ruta_destino = resolver_ruta_dinamica(nombre_archivo)
    print(f"💾 Exportando {nombre_archivo} a: {ruta_destino}")
    
    try: session.findById("wnd[0]/tbar[1]/btn[45]").press() 
    except: session.findById("wnd[0]/tbar[1]/btn[43]").press()
    time.sleep(2)
    
    try:
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except: pass
    time.sleep(3)
    
    box_ruta = session.findById("wnd[1]/usr/ctxtDY_PATH", False)
    box_archivo = session.findById("wnd[1]/usr/ctxtDY_FILENAME", False)
    
    if box_ruta and box_archivo:
        box_ruta.text = ruta_destino
        box_archivo.text = f"{nombre_archivo}.xlsx"
        session.findById("wnd[1]/tbar[0]/btn[0]").press() 
        manejar_seguridad(session) 
        manejar_seguridad(session) 
        time.sleep(4) 
        if session.findById("wnd[1]", False):
            try: session.findById("wnd[1]").close()
            except: pass
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

def main():
    try:
        print("🚀 Iniciando extracción múltiple UBICACION...")
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        session = sap_gui_auto.GetScriptingEngine.Children(0).Children(0)

        # 1. IH06 -> UbicacionesAAQ
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIH06"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        # TODO: Clics de filtros IH06
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(3)
        exportar_alv(session, "UbicacionesAAQ")

        # 2. CR05 -> PuestoTrabajo_AAQ
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nCR05"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        # TODO: Clics de filtros CR05
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(3)
        exportar_alv(session, "PuestoTrabajo_AAQ")

        # 3. IA09 -> IA09_CabeceraHojaRuta_AAQ
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIA09"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        # TODO: Clics de filtros IA09
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(3)
        exportar_alv(session, "IA09_CabeceraHojaRuta_AAQ")

        print("✅ Descarga múltiple completada.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()