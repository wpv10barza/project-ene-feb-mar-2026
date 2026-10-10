import win32com.client
import pandas as pd
import pyperclip
import time
import sys
import os
import glob
from datetime import datetime
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


# ================= CONFIGURACIÓN MAESTRA =================
BASE_DIR = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026"
SEMANA_ORIGEN = "06"
SEMANA_DESTINO = "07"
EXCEL_MAESTRO = os.path.join(BASE_DIR, SEMANA_ORIGEN, "libro1.xlsx")
FECHA_HOY = datetime.now().strftime("%d%m%y")
NOMBRE_ARCHIVO = f"IW47_Notificaciones_{FECHA_HOY}"

def resolver_ruta_dinamica(nombre_base, carpeta_fallback):
    try:
        df = pd.read_excel(EXCEL_MAESTRO, sheet_name=SEMANA_ORIGEN)
        fila = df[df['Name'].str.contains(nombre_base, case=False, na=False)]
        if not fila.empty:
            ruta_06 = str(fila.iloc[0]['Folder Path']).strip()
            ruta_final = ruta_06.replace(f"\\{SEMANA_ORIGEN}\\", f"\\{SEMANA_DESTINO}\\")
            os.makedirs(ruta_final, exist_ok=True)
            return ruta_final
    except: pass
    ruta_final = os.path.join(BASE_DIR, SEMANA_DESTINO, carpeta_fallback)
    os.makedirs(ruta_final, exist_ok=True)
    return ruta_final

def obtener_archivo_iw39():
    """Busca dinámicamente el Excel más reciente de CabeceraOrden"""
    print("   🔎 Buscando el Excel de IW39...")
    ruta_dinamica = resolver_ruta_dinamica("IW39_CabeceraOrden", "Ordenes")
    nombre_base = "IW39_CabeceraOrden"
    
    # Buscar archivos que coincidan
    archivos = glob.glob(os.path.join(ruta_dinamica, f"{nombre_base}*.xlsx"))
    if not archivos:
        ruta_raiz = os.path.join(BASE_DIR, SEMANA_DESTINO)
        archivos = glob.glob(os.path.join(ruta_raiz, f"{nombre_base}*.xlsx"))
        
    if archivos:
        # Tomar el más reciente en caso de haber varios
        archivo_mas_reciente = max(archivos, key=os.path.getctime)
        print(f"   📂 Extrayendo Órdenes desde: {archivo_mas_reciente}")
        return archivo_mas_reciente
    else:
        print(f"   ❌ ERROR: No se encontró ningún Excel de IW39 en las rutas.")
        sys.exit(1)

def busca_columna_orden_en_archivo(archivo):
    try:
        df = pd.read_excel(archivo)
        
        # 1. Prioridad Absoluta: Buscar la columna que se llame EXACTAMENTE "Orden"
        for col in df.columns:
            if str(col).strip().lower() == 'orden':
                return col
                
        # 2. Búsqueda Segura: Que contenga 'orden' pero que NO sea 'clase de orden' ni 'tipo'
        for col in df.columns:
            col_lower = str(col).lower()
            if 'orden' in col_lower and 'clase' not in col_lower and 'tipo' not in col_lower:
                return col
                
        # 3. Fallback en caso de emergencia
        for col in df.columns:
            if 'orden' in str(col).lower():
                return col
                
        return None
    except Exception as e:
        print(f"   ❌ Error al buscar columna de orden: {e}")
        return None

def copiar_columna_ordenes_encontrada_en_portapapeles(archivo):
    col_orden = busca_columna_orden_en_archivo(archivo)
    if not col_orden:
        print(f"   ❌ No se encontró una columna relacionada con 'Orden' en el archivo.")
        sys.exit(1)
    try:
        df = pd.read_excel(archivo)
        # Extraemos, limpiamos decimales y nos aseguramos de que sean valores numéricos reales
        ordenes_crudas = df[col_orden].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist()
        
        # Filtrar para quedarnos solo con códigos numéricos de orden (Ej: 4015848)
        ordenes = [o for o in ordenes_crudas if o.isdigit()]
        
        pyperclip.copy("\r\n".join(ordenes))
        print(f"   📋 ¡{len(ordenes)} órdenes copiadas al portapapeles desde la columna '{col_orden}'!")
        
        if ordenes:
            # Imprime las 5 primeras filas como ejemplo
            print(f"   🔍 Ejemplo de Órdenes Correctas: {ordenes[:5]}")
            
    except Exception as e:
        print(f"   ❌ Error al copiar órdenes: {e}")
        sys.exit(1)        

def exportador_jm5(session, nombre_archivo, ruta_destino):
    """Exportador que emula exactamente el menú y los clics del archivo JM5.vbs"""
    print(f"   💾 Exportando usando menú clásico hacia: {ruta_destino}")
    time.sleep(2)
    
    try:
        # Menú: Lista -> Exportar -> Fichero Local
        session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select()
        time.sleep(1)
        
        # Seleccionar 'Hoja de cálculo'
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1.5)
        
        # Llenar la ruta y el nombre clásico de guardado (.xls / .xlsx)
        session.findById("wnd[1]/usr/ctxt[0]").text = ruta_destino
        session.findById("wnd[1]/usr/ctxt[1]").text = f"{nombre_archivo}.xlsx"
        session.findById("wnd[1]/tbar[0]/btn[0]").press() 
        time.sleep(1)
        
        # Confirmar sobrescritura si aparece la ventana
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press() 
        except: pass
        
    except Exception as e:
        print(f"   ⚠️ Error en flujo de exportación de menú: {e}")

    archivo_completo = os.path.join(ruta_destino, f"{nombre_archivo}.xlsx")
    print("   ⏳ Verificando creación de archivo...")
    intentos = 0
    while not os.path.exists(archivo_completo) and intentos < 15:
        time.sleep(1)
        intentos += 1
    if os.path.exists(archivo_completo): print(f"   ✅ ¡Éxito! Archivo guardado correctamente.")
    else: print(f"   ⚠️ SAP tardó mucho. Verifica la carpeta manualmente.")
    
    try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
    except: pass

def main():
    print(f"🚀 Iniciando Script 3: IW47 (Notificaciones)...")
    
    # Llamamos a las funciones actualizadas de búsqueda y copia
    archivo_iw39 = obtener_archivo_iw39()
    copiar_columna_ordenes_encontrada_en_portapapeles(archivo_iw39)
    
    ruta_destino = resolver_ruta_dinamica("IW47_Notificaciones", "Notificaciones")
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    session = sap_gui_auto.GetScriptingEngine.Children(0).Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NIW47"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    
    # Inyectar órdenes en Matchcode
    print("   🎯 Pegando datos del portapapeles en el matchcode...")
    try:
        session.findById("wnd[0]/usr/btn[0]").press() # Botón Selección múltiple
        time.sleep(1)
        
        # Botón borrar selecciones antiguas
        try: session.findById("wnd[1]/tbar[0]/btn[25]").press() 
        except: pass
        
        session.findById("wnd[1]/tbar[0]/btn[24]").press() # Botón Pegar
        time.sleep(1)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Botón Ejecutar
        time.sleep(1)
    except Exception as e:
        print(f"   ⚠️ Fallo al inyectar las órdenes: {e}")

    # Borrar fechas para que no limiten los resultados
    print("   🧹 Limpiando fechas de búsqueda...")
    try:
        session.findById("wnd[0]/usr/ctxt[15]").text = ""
        session.findById("wnd[0]/usr/ctxt[16]").text = ""
    except: pass

    # Plantilla MCARRANZA01 forzada directamente
    print("   🎨 Aplicando Variante MCARRANZA01...")
    try:
        session.findById("wnd[0]/usr/ctxt[43]").text = "MCARRANZA01"
    except Exception as e:
        print(f"   ⚠️ No se encontró la casilla ctxt[43] para el Layout: {e}")

    # Ejecutar reporte
    print("   ⏳ Cargando la tabla ALV...")
    session.findById("wnd[0]/tbar[1]/btn[8]").press() 
    time.sleep(6)
    
    # Exportar usando el método de JM5
    exportador_jm5(session, NOMBRE_ARCHIVO, ruta_destino)

if __name__ == "__main__":
    main()