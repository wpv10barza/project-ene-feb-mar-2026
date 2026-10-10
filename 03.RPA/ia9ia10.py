import win32com.client
import pandas as pd
import pyperclip
import time
import sys
import os
import glob

# Asegurar que la consola muestre los mensajes correctamente
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# ============================================================================
# 0. PARÁMETROS DINÁMICOS
# ============================================================================
FECHA_DIA_FIJADO = "13.03.2026"  # <- Cambia esta fecha cuando lo necesites

# ============================================================================
# 1. RECIBIR LA RUTA DESDE LA TERMINAL Y CREAR CARPETAS
# ============================================================================
if len(sys.argv) < 2:
    print("❌ ERROR: No ingresaste ninguna ruta en la terminal.")
    sys.exit(1)

RUTA_BASE = sys.argv[1].strip()
if RUTA_BASE.startswith("--"):
    RUTA_BASE = RUTA_BASE[2:]

os.makedirs(RUTA_BASE, exist_ok=True)

# ✅ Subcarpeta fija para organizar descargas
CARPETA_DESCARGAS = os.path.join(RUTA_BASE, "Hoja de ruta")
os.makedirs(CARPETA_DESCARGAS, exist_ok=True)

print(f"===========================================================")
print(f" 📁 RUTA TERMINAL DETECTADA: {RUTA_BASE}")
print(f" 📂 CARPETA DE DESCARGAS: {CARPETA_DESCARGAS}")
print(f"===========================================================")

# ============================================================================
# 📚 DICCIONARIO DE REPORTES (IA10) - INTACTO
# ============================================================================
DICCIONARIO_REPORTES = {
    "IA10_OperacionHR_AAQ": [0, 1],       
    "IA10_PaquetesHR_AAQ": [0, 1, 3],     
    "IA10_MaterialHR_AAQ": [0, 1, 4],     
    "IA10_MafHR_AAQ": [0, 1, 5]           
}

# ============================================================================
# FUNCIONES SAP GENERALES (TUS ORIGINALES)
# ============================================================================
def conectar_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        return application.Children(0).Children(0)
    except Exception as e:
        print(f"❌ Error conectando a SAP: {e}")
        sys.exit(1)

def manejar_seguridad(session):
    time.sleep(1)
    try:
        if session.findById("wnd[1]", False):
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
    except: pass

# ============================================================================
# PASO 1: EJECUCIÓN IA09 (CON IDs REALES GRABADOS)
# ============================================================================
def ejecutar_ia09(session, carpeta):
    print("\n" + "="*60)
    print("🤖 PASO 1: EXTRACCIÓN CABECERA DE HOJA DE RUTA (IA09)")
    print("="*60)
    
    nombre_archivo = "IA09_CabeceraHojaRuta_AAQ"
    ruta_completa = os.path.join(carpeta, f"{nombre_archivo}.xlsx")
    
    # ✅ LÓGICA DE SALTAR: Si el archivo ya existe, no entra a SAP
    if os.path.exists(ruta_completa):
        print(f"   ⏩ El archivo {nombre_archivo}.xlsx ya existe. Saltando extracción...")
        return ruta_completa
        
    print("🚀 Abriendo Transacción IA09...")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nIA09"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)
    
    print(f"⚙️ Inyectando Día fijado ({FECHA_DIA_FIJADO}) y Centro (C601)...")
    
    try: session.findById("wnd[0]/usr/ctxt[6]").text = FECHA_DIA_FIJADO
    except Exception as e: print(f"⚠️ Error al ingresar fecha: {e}")
    
    try: session.findById("wnd[0]/usr/ctxt[9]").text = "C601"
    except Exception as e: print(f"⚠️ Error al ingresar centro: {e}")
    
    session.findById("wnd[0]").sendVKey(8) # F8 para ejecutar
    time.sleep(4)
    
    print(f"   -> Exportando {nombre_archivo}...")
    try: session.findById("wnd[0]/tbar[1]/btn[16]").press()
    except:
        try: session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        except: pass
    time.sleep(2)
    
    try:
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except: pass
    time.sleep(3)
    
    try:
        if session.findById("wnd[1]/usr/ssub/1/txt", False):
            session.findById("wnd[1]/usr/ssub/1/txt").text = nombre_archivo
            session.findById("wnd[1]/tbar[0]/btn[20]").press() 
            time.sleep(2)
    except: pass

    try:
        ruta_ingresada = False
        box_ruta = session.findById("wnd[1]/usr/ctxt[0]", False)
        if box_ruta:
            box_ruta.text = carpeta
            box_archivo = session.findById("wnd[1]/usr/ctxt[1]", False)
            if box_archivo and box_archivo.text == "":
                box_archivo.text = f"{nombre_archivo}.xlsx"
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
            ruta_ingresada = True
        elif session.findById("wnd[1]/usr/ctxtDY_PATH", False):
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = carpeta
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{nombre_archivo}.xlsx"
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
            ruta_ingresada = True
            
        if ruta_ingresada:
            manejar_seguridad(session) 
            manejar_seguridad(session) 
            print(f"   ✅ Archivo guardado exitosamente.")
            time.sleep(3) 
    except: pass
    
    return ruta_completa

# ============================================================================
# PASO 2: DEPURACIÓN EXCEL IA09 (TU LÓGICA ORIGINAL)
# ============================================================================
def depurar_y_preparar_datos(ruta_archivo):
    print("\n" + "="*60)
    print("🔎 PASO 2: DEPURACIÓN DE DATOS EXCEL (IA09)")
    print("="*60, flush=True)
    if not os.path.exists(ruta_archivo):
        print(f"❌ ERROR: No existe el archivo base en:\n{ruta_archivo}")
        return False
    try:
        print("   📂 Leyendo archivo Excel...", end=" ", flush=True)
        df = pd.read_excel(ruta_archivo)
        print("✅ OK")
        
        col_grupo = None
        nombres_exactos = ["Grupo hojas ruta", "Grupo", "Grupo HR", "GrHjRuta"]
        
        for nombre in nombres_exactos:
            if nombre in df.columns:
                col_grupo = nombre
                break
        
        if not col_grupo:
            for col in df.columns:
                col_lower = str(col).lower()
                if "grupo" in col_lower and "planif" not in col_lower:
                    col_grupo = col
                    break
                    
        if not col_grupo:
            print("\n❌ ERROR: No se encontró la columna de Grupo de Hoja de Ruta.")
            return False
            
        print(f"   🎯 Columna detectada para extraer Grupos HR: '{col_grupo}'")
        
        col_pto = next((c for c in df.columns if "puesto" in str(c).lower()), None)
        if col_pto:
            print(f"   ⚙️ Filtrando por Puesto de trabajo 'QI3PP'...", end=" ", flush=True)
            df_filtered = df[df[col_pto].astype(str).str.contains('QI3PP', na=False)].copy()
        else:
            df_filtered = df.copy()

        grupos_brutos = df_filtered[col_grupo].dropna().unique()
        
        if len(grupos_brutos) == 0:
            print("\n⚠️ AVISO: No se encontraron grupos para copiar.")
            return False
            
        grupos_limpios = [str(int(g)) if isinstance(g, float) else str(g).strip() for g in grupos_brutos]
        pyperclip.copy("\r\n".join(grupos_limpios))
        
        print(f"\n   ✅ OK ({len(grupos_limpios)} grupos copiados al portapapeles)")
        return True
    except Exception as e:
        print(f"\n❌ ERROR procesando el Excel: {e}")
        return False

# ============================================================================
# PASO 3: FUNCIONES DE DESCARGA IA10 (TUS CLICS ORIGINALES INTACTOS)
# ============================================================================
def configurar_checkboxes(session, casillas_activas):
    for i in range(12): 
        try:
            estado = True if i in casillas_activas else False
            session.findById(f"wnd[0]/usr/chk[{i}]").selected = estado
        except: pass 

def retornar_a_filtros(session):
    print("   🔙 Retornando a la pantalla de selección...")
    reintentos = 0
    while reintentos < 5:
        if session.findById("wnd[0]/usr/btn[2]", False):
            return True
        try:
            if session.findById("wnd[1]", False):
                try: session.findById("wnd[1]").close()
                except: session.findById("wnd[1]/tbar[0]/btn[12]").press()
        except: pass

        try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except: pass
        
        time.sleep(1)
        reintentos += 1
    return False

def descargar_reporte(session, nombre_archivo, carpeta):
    print(f"   -> Exportando como: {nombre_archivo}.xlsx...")
    
    try: session.findById("wnd[0]/tbar[1]/btn[45]").press() 
    except:
        try: session.findById("wnd[0]/tbar[1]/btn[43]").press() 
        except: 
            print("   ⚠️ No se encontró el botón de exportar en esta vista.")
            return False

    time.sleep(2) 
    
    try:
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except: pass 

    time.sleep(3) 
    
    try:
        if session.findById("wnd[1]/usr/ssub/1/txt", False):
            session.findById("wnd[1]/usr/ssub/1/txt").text = nombre_archivo
            session.findById("wnd[1]/tbar[0]/btn[20]").press() 
            time.sleep(2)
    except: pass

    try:
        ruta_ingresada = False
        box_ruta = session.findById("wnd[1]/usr/ctxt[0]", False)
        
        if box_ruta:
            box_ruta.text = carpeta  
            
            box_archivo = session.findById("wnd[1]/usr/ctxt[1]", False)
            if box_archivo and box_archivo.text == "":
                box_archivo.text = f"{nombre_archivo}.xlsx"
                
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
            ruta_ingresada = True
            
        elif session.findById("wnd[1]/usr/ctxtDY_PATH", False):
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = carpeta 
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{nombre_archivo}.xlsx"
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
            ruta_ingresada = True
            
        if ruta_ingresada:
            manejar_seguridad(session) 
            manejar_seguridad(session) 
            print(f"   ✅ Archivo guardado exitosamente.")
            time.sleep(3) 
        else:
            print("   ⚠️ No aparecieron los campos de texto para la Ruta.")
            
    except Exception as e:
        print(f"   ⚠️ Error crítico al guardar: {e}")

# ============================================================================
# FLUJO PRINCIPAL MAESTRO
# ============================================================================
def main():
    session = conectar_sap()
    
    # 1. Ejecutar IA09 primero (Guardado en la nueva carpeta)
    archivo_ia09 = ejecutar_ia09(session, CARPETA_DESCARGAS)
    
    # 2. Leer Excel y preparar datos
    if not depurar_y_preparar_datos(archivo_ia09): 
        sys.exit(1)

    print("\n" + "="*60)
    print("🤖 PASO 3: AUTOMATIZACIÓN SAP IA10 (CICLO CONTINUO)")
    print("="*60)
    
    print("🚀 Abriendo Transacción IA10...")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nIA10"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)
    
    print("⚙️ Aplicando filtros base y pegando portapapeles...")
    
    # ✅ INYECCIÓN DE LA FECHA (Descubierta en fia10.vbs)
    try: session.findById("wnd[0]/usr/ctxt[6]").text = FECHA_DIA_FIJADO
    except Exception as e: print(f"⚠️ Error al ingresar fecha en IA10: {e}")
    
    session.findById("wnd[0]/usr/ctxt[9]").text = "C601"  
    session.findById("wnd[0]/usr/ctxt[13]").text = "4"    
    
    session.findById("wnd[0]/usr/btn[2]").press() 
    time.sleep(1)
    session.findById("wnd[1]/tbar[0]/btn[24]").press() 
    time.sleep(1)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()  
    time.sleep(1)

    contador = 1
    total = len(DICCIONARIO_REPORTES)

    for nombre_archivo, casillas in DICCIONARIO_REPORTES.items():
        print(f"\n" + "-"*40)
        print(f"📋 Ejecutando {contador}/{total}: {nombre_archivo}")
        
        # ✅ LÓGICA DE SALTAR: Si el reporte de IA10 ya existe, omitir
        if os.path.exists(os.path.join(CARPETA_DESCARGAS, f"{nombre_archivo}.xlsx")):
            print(f"   ⏩ El archivo ya existe en la carpeta. Saltando extracción...")
            contador += 1
            continue
        
        configurar_checkboxes(session, casillas)
        
        session.findById("wnd[0]").sendVKey(8) 
        time.sleep(4)
        
        if session.findById("wnd[0]/usr/btn[2]", False):
            mensaje_sap = ""
            try: mensaje_sap = session.findById("wnd[0]/sbar").text
            except: pass
            print(f"   ⚠️ No hay datos para este ciclo. SAP dice: {mensaje_sap}")
            print("   ⏩ Saltando al siguiente reporte...")
        else:
            # ✅ Exportando directo a la carpeta de descargas
            descargar_reporte(session, nombre_archivo, CARPETA_DESCARGAS)
            retornar_a_filtros(session) 
            
        contador += 1
    
    print("\n" + "="*60)
    print("🏁 AUTOMATIZACIÓN FINALIZADA CORRECTAMENTE.")
    print("="*60)

if __name__ == "__main__":
    main()