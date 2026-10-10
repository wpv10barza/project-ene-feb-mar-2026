from datetime import datetime
import os
import win32com.client
import sys
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


def run_sap_script():
    try:
        # Conexión al motor de SAP GUI
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # Inicio de Transacción IA10
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIA10"
        session.findById("wnd[0]").sendVKey(0)

        # Desmarcar Checkboxes (Limpieza de selección)
        for i in [3, 9, 10, 7, 4, 5, 2, 8, 11]:
            session.findById(f"wnd[0]/usr/chk[{i}]").selected = False
        
        session.findById("wnd[0]/usr/ctxt[9]").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press() # Enter en mensaje

        # Carga de datos específicos
        session.findById("wnd[0]/usr/ctxt[4]").text = "68983"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        # --- SECCIÓN 1: PAQUETES ---
        session.findById("wnd[0]/tbar[1]/btn[18]").press()
        session.findById("wnd[0]/mbar/menu[4]/menu[2]/menu[9]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        # Exportar Layout Paquetes
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ssub/1/txt").text = "IA10_PaquetesHR_AAQ"
        session.findById("wnd[1]/tbar[0]/btn[20]").press() # Guardar
        
        # --- SECCIÓN 2: MATERIAL ---
        session.findById("wnd[0]/tbar[0]/btn[3]").press() # Regresar
        session.findById("wnd[0]/usr/chk[6]").selected = False
        session.findById("wnd[0]/usr/chk[4]").selected = True
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/mbar/menu[4]/menu[2]/menu[7]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ssub/1/txt").text = "IA10_MaterialHR_AAQ"
        session.findById("wnd[1]/tbar[0]/btn[20]").press()

        # --- SECCIÓN 3: MAF ---
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/usr/chk[4]").selected = False
        session.findById("wnd[0]/usr/chk[5]").selected = True
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/mbar/menu[4]/menu[2]/menu[8]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ssub/1/txt").text = "IA10_MafHR_AAQ"
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        print("✅ Proceso IA10 finalizado exitosamente.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_sap_script()