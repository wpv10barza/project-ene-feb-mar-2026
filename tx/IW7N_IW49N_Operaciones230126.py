from datetime import datetime
import win32com.client
import pandas as pd
import pyperclip
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


# --- FUNCIÓN 1: DEPURACIÓN DE DATOS ---
def depurar_y_preparar_datos(ruta_archivo):
    print("\n" + "="*60)
    print("🔎 PASO 1: DEPURACIÓN DE DATOS EXCEL")
    print("="*60, flush=True)
    if not os.path.exists(ruta_archivo):
        print(f"❌ ERROR: No existe el archivo en: {ruta_archivo}")
        return None
    try:
        print(f"Leyendo Excel...", end=" ", flush=True)
        df = pd.read_excel(ruta_archivo)
        col = 'Orden' if 'Orden' in df.columns else df.columns[5]
        lista_limpia = (df[col].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist())
        print(f"✅ {len(lista_limpia)} órdenes listas.")
        return lista_limpia
    except Exception as e:
        print(f"❌ ERROR: {e}")
        return None

# --- FUNCIÓN 2: AUTOMATIZACIÓN INTEGRADA ---
def ejecutar_proceso_sap(lista_ordenes):
    print("\n" + "="*60)
    print("🤖 PASO 2: EJECUCIÓN INTEGRADA EN SAP GUI")
    print("="*60, flush=True)

    try:
        # CONEXIÓN
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        session = application.Children(0).Children(0)
        session.findById("wnd[0]").maximize()

        # --- SUB-PROCESO A: IW39 (DESCARGA CABECERA) ---
        print("STEP A: Iniciando IW39...", end=" ", flush=True)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
        session.findById("wnd[0]").sendVKey(0)
        
        session.findById("wnd[0]/usr/ctxt[13]").text = "qi3pp"
        session.findById("wnd[0]/usr/ctxt[75]").text = "01.01.2025"
        session.findById("wnd[0]/usr/ctxt[76]").text = "31.12.2999"
        
        # Selección de Layout
        session.findById("wnd[0]/usr/ctxt[131]").setFocus()
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[1]/usr").verticalScrollbar.position = 684
        session.findById("wnd[1]/usr/lbl[1,19]").setFocus()
        session.findById("wnd[1]").sendVKey(2)
        
        session.findById("wnd[0]/tbar[1]/btn[8]").press() # Ejecutar
        
        # Exportar IW39_CABECERA
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        session.findById("wnd[1]/usr/ssub/1/txt").text = "IW39_CABECERA"
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        print("✅ Archivo IW39 generado.")

        # --- SUB-PROCESO B: IW49N (CARGA Y EXPORTACIÓN OPERACIONES) ---
        print("STEP B: Iniciando IW49N con Carga Masiva...", end=" ", flush=True)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW49N"
        session.findById("wnd[0]").sendVKey(0)
        
        # Configuración de checks
        session.findById("wnd[0]/usr/chk[2]").selected = True
        session.findById("wnd[0]/usr/chk[3]").selected = True
        
        # CARGA DE LAS 802 ÓRDENES
        pyperclip.copy('\r\n'.join(lista_ordenes))
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub/1/2/btn[0]").press()
        time.sleep(1)
        session.findById("wnd[1]/tbar[0]/btn[25]").press() # Limpiar
        session.findById("wnd[1]/tbar[0]/btn[24]").press() # Pegar
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Validar
        
        # Filtros Pestaña 9 (TAB9)
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub/1/2/ctxt[5]").setFocus()
        session.findById("wnd[0]").sendVKey(4)
        
        # Selección Matchcode (F4)
        session.findById("wnd[1]").sendVKey(82)
        session.findById("wnd[1]/usr").verticalScrollbar.position = 925
        session.findById("wnd[1]/usr/lbl[1,19]").setFocus()
        session.findById("wnd[1]").sendVKey(2)
        
        # Ejecutar y Exportar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        # Manejo de diálogo de exportación
        try:
             session.findById("wnd[1]/tbar[0]/btn[20]").press()
        except:
             session.findById("wnd[1]/usr/ssub/1/txt").text = "IW7N_IW49N_Operaciones230126"
             session.findById("wnd[1]/tbar[0]/btn[20]").press()
        
        print("✅ Archivo IW49N generado.")

        print("\n" + "="*60)
        print("✨ PROCESO COMPLETO FINALIZADO CON ÉXITO ✨")
        print("="*60)

    except Exception as e:
        print(f"\n❌ ERROR EN SAP GUI: {e}")

if __name__ == "__main__":
    RUTA = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026\05\IW39_CABECERA.xlsx"
    
    ordenes = depurar_y_preparar_datos(RUTA)
    if ordenes:
        ejecutar_proceso_sap(ordenes)