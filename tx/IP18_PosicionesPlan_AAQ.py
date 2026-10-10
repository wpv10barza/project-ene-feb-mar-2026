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
# CONFIGURACIÓN DE RUTAS
# ============================================================================
RUTA_BASE = r"WEEK_DIR"
NOMBRE_ARCHIVO = "IP18_PosicionesPlan_AAQ"
RUTA_COMPLETA_EXCEL = os.path.join(RUTA_BASE, f"{NOMBRE_ARCHIVO}.xlsx")

def manejar_seguridad(session):
    """Cierra las ventanas emergentes de seguridad de Windows/SAP"""
    time.sleep(1)
    try:
        if session.findById("wnd[1]", False):
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
    except: pass

def limpiar_basura_excel():
    """Abre el Excel descargado y elimina las filas sin Grupo de Hoja de Ruta"""
    print("\n🧹 Iniciando limpieza automática del Excel...")
    try:
        df = pd.read_excel(RUTA_COMPLETA_EXCEL)
        total_original = len(df)
        
        df_limpio = df.dropna(subset=['Grupo hojas ruta'])
        
        total_limpio = len(df_limpio)
        basura_eliminada = total_original - total_limpio
        
        df_limpio.to_excel(RUTA_COMPLETA_EXCEL, index=False)
        print(f"   ✅ Limpieza exitosa: Se eliminaron {basura_eliminada} registros basura.")
        print(f"   📊 Registros válidos listos para IA10: {total_limpio}")
    except Exception as e:
        print(f"   ⚠️ No se pudo limpiar el Excel automáticamente: {e}")

def run_sap_script():
    try:
        print("🚀 Iniciando extracción optimizada de IP18...")
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        session = application.Children(0).Children(0)

        # 1. INICIO
        session.findById("wnd[0]").maximize()
        try: session.findById("wnd[0]/tbar[0]/btn[3]").press() # Asegurar volver al inicio
        except: pass
        
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIP18"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # 2. PARÁMETROS BÁSICOS
        print("⚙️ Configurando parámetros (Centro C601)...")
        session.findById("wnd[0]/usr/ctxt[16]").text = "C601" 
        
        # 3. SELECCIÓN DE VARIANTE (Una sola vez)
        print("📋 Aplicando Layout/Variante...")
        session.findById("wnd[0]/usr/ctxt[80]").setFocus()
        session.findById("wnd[0]").sendVKey(4) # F4
        time.sleep(1.5) 
        
        # Manejo de popups molestos de variantes
        try:
            if session.findById("wnd[2]", False): session.findById("wnd[2]").close()
            if session.findById("wnd[1]/tbar[0]/btn[32]", False): session.findById("wnd[1]/tbar[0]/btn[32]").press()
        except: pass
        time.sleep(0.5)

        # Scroll y doble clic a la variante
        session.findById("wnd[1]/usr").verticalScrollbar.position = 70
        session.findById("wnd[1]/usr/lbl[1,20]").setFocus()
        session.findById("wnd[1]").sendVKey(2) # F2
        time.sleep(1)
        
        # 4. EJECUCIÓN DIRECTA (Se eliminó el ciclo inútil de F8 -> F3 -> F8)
        print("⏳ Ejecutando reporte...")
        session.findById("wnd[0]/tbar[1]/btn[8]").press() # F8
        time.sleep(4)
        session.findById("wnd[0]").maximize()
        
        # 5. EXPORTACIÓN INMEDIATA ALV
        print("💾 Exportando resultados a Excel...")
        try: session.findById("wnd[0]/tbar[1]/btn[45]").press() 
        except: session.findById("wnd[0]/tbar[1]/btn[43]").press()
        time.sleep(2)
        
        try:
            session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass
        time.sleep(3)
        
        # Guardado estructurado
        box_ruta = session.findById("wnd[1]/usr/ctxtDY_PATH", False)
        box_archivo = session.findById("wnd[1]/usr/ctxtDY_FILENAME", False)
        
        if box_ruta and box_archivo:
            box_ruta.text = RUTA_BASE
            box_archivo.text = f"{NOMBRE_ARCHIVO}.xlsx"
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
            manejar_seguridad(session) 
            manejar_seguridad(session) 
            time.sleep(4) 
            
            if session.findById("wnd[1]", False):
                try: session.findById("wnd[1]").close()
                except: pass
        else:
            print("⚠️ Error: No se encontraron los campos de guardado.")

        # 6. CIERRE
        print("   🔙 Retornando al menú principal...")
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except: pass

        print("✅ Descarga IP18 completada. Procesando limpieza...")
        limpiar_basura_excel()

    except Exception as e:
        print(f"❌ Error crítico detectado: {e}")

if __name__ == "__main__":
    if not os.path.exists(RUTA_BASE):
        os.makedirs(RUTA_BASE)
    run_sap_script()