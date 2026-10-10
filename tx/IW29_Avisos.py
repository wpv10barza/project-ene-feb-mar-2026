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
        # Conexión con el motor de scripting de SAP GUI
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # Configuración de ventana
        session.findById("wnd[0]").resizeWorkingPane(76, 22, False)
        
        # Transacción IW29
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW29"
        session.findById("wnd[0]").sendVKey(0)

        # Llenado de parámetros de selección
        session.findById("wnd[0]/usr/ctxt[15]").text = "25.01.2025"
        session.findById("wnd[0]/usr/ctxt[16]").text = "25.01.2026"
        session.findById("wnd[0]/usr/ctxt[42]").text = "QI3PP"
        session.findById("wnd[0]/usr/ctxt[46]").text = "C601"
        session.findById("wnd[0]/usr/ctxt[124]").text = "/MCARRANZA02"
        
        # Interacción y navegación
        session.findById("wnd[0]/usr/ctxt[124]").setFocus()
        session.findById("wnd[0]/usr/ctxt[124]").caretPosition = 12
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Limpieza de campo y nueva ejecución
        session.findById("wnd[0]/usr/ctxt[42]").setFocus()
        session.findById("wnd[0]/usr/ctxt[42]").caretPosition = 5
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[1]").sendVKey(4)
        session.findById("wnd[2]").close()
        session.findById("wnd[1]").close()
        
        session.findById("wnd[0]/usr/ctxt[42]").text = ""
        session.findById("wnd[0]/usr/ctxt[42]").caretPosition = 0
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Selección en lista
        session.findById("wnd[0]/usr/ctxt[124]").setFocus()
        session.findById("wnd[0]/usr/ctxt[124]").caretPosition = 8
        session.findById("wnd[0]").sendVKey(4)
        
        session.findById("wnd[1]/usr").verticalScrollbar.position = 646
        session.findById("wnd[1]/usr/lbl[1,19]").setFocus()
        session.findById("wnd[1]/usr/lbl[1,19]").caretPosition = 6
        session.findById("wnd[1]").sendVKey(2)
        
        # Ejecutar y manejar el Grid
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
        grid.firstVisibleRow = 0
        
        session.findById("wnd[0]").resizeWorkingPane(89, 22, False)
        session.findById("wnd[0]").sendVKey(32)
        
        if len(connection.Children) > 1: # Verificar si se abrió ventana emergente
             session.findById("wnd[1]").close()
             
        session.findById("wnd[0]").maximize()
        
        grid.setCurrentCell(-1, "ERDAT")
        grid.selectAll()
        grid.contextMenu()

    except Exception as e:
        print(f"Error detectado: {e}")

if __name__ == "__main__":
    run_sap_script()