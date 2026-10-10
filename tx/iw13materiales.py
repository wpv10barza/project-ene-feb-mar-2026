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


# ============================================================================
# CONFIGURACIÓN DE RUTAS
# ============================================================================
RUTA_BASE = r"WEEK_DIR\MaterialesHR - copia"

# OJO: Puse IA10 porque ahí estaban escondidos tus materiales "AM". 
# Si realmente los vas a sacar de IA09, cambia este nombre.
ARCHIVO_ENTRADA = os.path.join(RUTA_BASE, "IA10_MaterialHR_AAQ.xlsx")

# ============================================================================
# FUNCIÓN 1: DEPURACIÓN Y FILTRADO (EXCEL -> PORTAPAPELES)
# ============================================================================
def depurar_y_preparar_datos(ruta_archivo):
    print(f"🔎 1. Analizando archivo: {os.path.basename(ruta_archivo)}")
    
    if not os.path.exists(ruta_archivo):
        print("❌ Archivo no encontrado. Verifica la ruta.")
        sys.exit()
        
    try:
        # Leer Excel
        df = pd.read_excel(ruta_archivo, header=None, engine='openpyxl')
        contenido = df.to_string(index=False, header=False)
        
        # Extraer códigos AM... y 10...
        patron_materiales = re.findall(r'\bAM\d{5,10}\b|\b10\d{8,15}\b', contenido)
        materiales_unicos = list(set(patron_materiales))
        
        if materiales_unicos:
            # Copiar al portapapeles
            texto_portapapeles = '\r\n'.join(materiales_unicos)
            pyperclip.copy(texto_portapapeles)
            print(f"✅ ÉXITO: {len(materiales_unicos)} materiales copiados al portapapeles.")
            return True
        else:
            print("⚠️ No se encontraron materiales con el formato correcto en el Excel.")
            sys.exit()
            
    except PermissionError:
        print("❌ ERROR: El archivo Excel está abierto. Ciérralo y vuelve a ejecutar.")
        sys.exit()
    except Exception as e:
        print(f"❌ Error al procesar el Excel: {e}")
        sys.exit()

# ============================================================================
# FUNCIÓN 2: AUTOMATIZACIÓN SAP (TRADUCCIÓN DE JM3.VBS)
# ============================================================================
def ejecutar_sap_mm60():
    print("\n🤖 2. Conectando con SAP GUI...")
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        print("✅ Conexión a SAP exitosa.")
    except Exception as e:
        print("❌ No se pudo conectar a SAP. Asegúrate de tener SAP Login abierto y conectado.")
        sys.exit()

    try:
        print("🚀 Ejecutando transacción /nMM60...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM60"
        session.findById("wnd[0]").sendVKey(0) # Enter
        time.sleep(1) # Pausa breve para que SAP cargue la pantalla
        
        print("📋 Abriendo Matchcode y pegando datos...")
        # 1. Clic en botón de Selección Múltiple (Flecha amarilla del campo Material)
        # (Sacado de tu VBS: session.findById("wnd[0]/usr/btn[0]").press)
        session.findById("wnd[0]/usr/btn[0]").press()
        time.sleep(0.5)
        
        # 2. Clic en "Pegar del Portapapeles" (Botón 24 = Shift+F12)
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        time.sleep(0.5)
        
        # 3. Clic en "Tomar/Copiar" (Botón 8 = F8) para cerrar la ventana emergente
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(0.5)
        
        print("⚙️ Ejecutando reporte MM60...")
        # 4. Clic en Ejecutar reporte principal (Botón 8 de la barra 1 = F8)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        print("🎉 ¡Proceso SAP completado con éxito! El reporte está cargando en tu pantalla...")
        
    except Exception as e:
        print(f"\n❌ Error durante la ejecución de los clics en SAP:\n{e}")
        print("👉 Tip: Asegúrate de estar en la pantalla principal de SAP antes de correr el script.")

# ============================================================================
# EJECUCIÓN PRINCIPAL
# ============================================================================
if __name__ == "__main__":
    print("="*60)
    print("🚀 INICIANDO BOT: EXCEL -> PORTAPAPELES -> SAP MM60")
    print("="*60)
    
    # 1. Leer el Excel y mandar al portapapeles
    if depurar_y_preparar_datos(ARCHIVO_ENTRADA):
        # 2. Si el paso 1 fue exitoso, controlar SAP
        ejecutar_sap_mm60()