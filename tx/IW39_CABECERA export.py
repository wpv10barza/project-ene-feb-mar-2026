from datetime import datetime
import os
import win32com.client
import pandas as pd
import pyperclip
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


def ejecutar_iw49n_con_excel():
    try:
        # --- CONEXIÓN A SAP ---
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            print("No se encontró la instancia de SAP GUI.")
            return
        
        application = sap_gui_auto.GetScriptingEngine
        if not application:
            print("No se pudo obtener el Scripting Engine.")
            return
            
        connection = application.Children(0)
        session = connection.Children(0)
        
    except Exception as e:
        print(f"Error crítico al conectar con SAP: {e}")
        return

    try:
        # --- INICIO DE NAVEGACIÓN ---
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW49N"
        session.findById("wnd[0]").sendVKey(0)

        # Selección de Checkboxes según el VBS original
        session.findById("wnd[0]/usr/chk[2]").selected = True
        session.findById("wnd[0]/usr/chk[3]").selected = True
        session.findById("wnd[0]/usr/chk[3]").setFocus()

        # Abrir Ventana de Selección Múltiple (Botón al lado del campo de Orden)
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub/1/2/btn[0]").press()

        # --- LECTURA DE EXCEL Y PORTAPAPELES ---
        # Ruta proporcionada en tu solicitud
        excel_path = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026\05\IW39_CABECERA.xlsx"
        
        # Leemos la columna F (index 5) del Excel
        df = pd.read_excel(excel_path, usecols=[5], header=None)
        
        # Limpiamos valores nulos y convertimos a string
        lista_ordenes = df.iloc[:, 0].dropna().astype(str).tolist()
        orders_string = '\r\n'.join(lista_ordenes) # SAP prefiere saltos de línea Windows
        
        # Copiar al portapapeles
        pyperclip.copy(orders_string)

        # --- PEGAR EN SAP (Ventana wnd[1]) ---
        # Asegurarse de estar en la pestaña correcta si el VBS lo indicaba
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select()
        
        # El botón 24 en SAP suele ser 'Pegar del Portapapeles'
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        
        # El botón 8 es 'Ejecutar/Continuar' (el check verde)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Cerramos cualquier ventana residual según el flujo VBS
        try:
            session.findById("wnd[1]").close()
        except:
            pass # Si ya se cerró al presionar ejecutar, ignoramos el error

        # --- CONTINUACIÓN DEL FLUJO VBS (Tabs y Focus) ---
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub/1/2/ctxt[5]").setFocus()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub/1/2/ctxt[5]").caretPosition = 5
        session.findById("wnd[0]").sendVKey(4) # Abre ayuda de búsqueda (F4)

        # Interacción con la ventana de ayuda de búsqueda
        session.findById("wnd[1]/usr").verticalScrollbar.position = 926
        session.findById("wnd[1]/usr/lbl[1,19]").setFocus()
        session.findById("wnd[1]/usr/lbl[1,19]").caretPosition = 9
        session.findById("wnd[1]").sendVKey(2) # Doble click / Seleccionar

        # Cambio a Tab 5 y reinicio de transacción (según tu VBS final)
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5").select()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW49N"
        
        print("Proceso completado con éxito.")

    except Exception as e:
        print(f"Error durante la ejecución de los pasos: {e}")

if __name__ == "__main__":
    ejecutar_iw49n_con_excel()