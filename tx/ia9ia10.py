import win32com.client
import pandas as pd
import pyperclip
import time
import sys
import os
from datetime import datetime

# ============================================================================
# CONFIGURACIÓN MAESTRA Y RUTAS (AUTO)
# - NO importa desde qué carpeta lo ejecutes.
# - Puedes pasar la carpeta de semana como 1er argumento:
#     python ia9ia10.py "...\2026\08"
# - O puedes definir variables de entorno:
#     SAP_DATALOG_BASE_DIR  (ruta hasta ...\2026)
#     SAP_DATALOG_SEMANA    (ej. 08)
#     SAP_DATALOG_SEMANA_ORIGEN (ej. 07)
# ============================================================================
DEFAULT_BASE_DIR = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026"

def _calcular_semana_origen(semana_destino: str) -> str:
    try:
        w = int(str(semana_destino).strip())
        return str(max(w - 1, 1)).zfill(2)
    except:
        return "01"

def _cargar_config_desde_env():
    base_dir = os.getenv("SAP_DATALOG_BASE_DIR", DEFAULT_BASE_DIR)
    semana_destino = os.getenv("SAP_DATALOG_SEMANA", "07")
    semana_origen = os.getenv("SAP_DATALOG_SEMANA_ORIGEN", _calcular_semana_origen(semana_destino))
    excel_maestro = os.path.join(base_dir, semana_origen, "libro1.xlsx")
    fecha_hoy = datetime.now().strftime("%d%m%y")
    return base_dir, semana_origen, semana_destino, excel_maestro, fecha_hoy

BASE_DIR, SEMANA_ORIGEN, SEMANA_DESTINO, EXCEL_MAESTRO, FECHA_HOY = _cargar_config_desde_env()

NOMBRE_ARCHIVO = "IA10_MafHR_AAQ" # Nombre extraído de tu JM1.vbs

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

def leer_ia09_y_copiar():
    print("   🔎 Buscando el Excel de IA09_Cabecera...")
    ruta_dinamica = resolver_ruta_dinamica("IA09_CabeceraHojaRuta", "HojasRuta")
    ruta_raiz = os.path.join(BASE_DIR, SEMANA_DESTINO)
    
    posibles_nombres = [f"IA09_CabeceraHojaRuta_{FECHA_HOY}.xlsx", "IA09_CabeceraHojaRuta_AAQ.xlsx"]
    archivo_ia09 = None
    
    for nombre in posibles_nombres:
        if os.path.exists(os.path.join(ruta_dinamica, nombre)):
            archivo_ia09 = os.path.join(ruta_dinamica, nombre)
            break
        elif os.path.exists(os.path.join(ruta_raiz, nombre)):
            archivo_ia09 = os.path.join(ruta_raiz, nombre)
            break

    if not archivo_ia09:
        print(f"❌ ERROR: No se encontró el Excel de IA09 para sacar los Grupos de Hoja de Ruta.")
        sys.exit(1)
        
    print(f"   📂 Leyendo datos desde: {archivo_ia09}")
    try:
        df = pd.read_excel(archivo_ia09)
        col_grupo = next((col for col in df.columns if "grupo" in str(col).lower()), df.columns[0])
        grupos = df[col_grupo].dropna().apply(lambda x: str(x).replace('.0', '').strip()).tolist()
        grupos_unicos = list(dict.fromkeys(grupos)) 
        pyperclip.copy("\r\n".join(grupos_unicos))
        print(f"   📋 ¡{len(grupos_unicos)} Grupos copiados al portapapeles!")
    except Exception as e:
        print(f"   ❌ Error leyendo IA09: {e}")
        sys.exit(1)

def main():
    print("🚀 Iniciando Extracción Única: IA10 MAF (Medios Auxiliares de Fabricación)...")
    leer_ia09_y_copiar()
    ruta_destino = resolver_ruta_dinamica("IA10_MafHR", "HojasRuta")
    print(f"   📂 Ruta de guardado: {ruta_destino}")
    
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        session = sap_gui_auto.GetScriptingEngine.Children(0).Children(0)
    except:
        print("❌ Error: SAP no está abierto.")
        sys.exit(1)

    # 1. ENTRAR A LA TRANSACCIÓN
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nia10"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    
    # 2. FILTROS Y MATCHCODE
    print("   ⚙️ Pegando códigos en el Matchcode...")
    try:
        session.findById("wnd[0]/usr/ctxt[9]").text = "C601"  
        session.findById("wnd[0]/usr/ctxt[13]").text = "4"    
        
        session.findById("wnd[0]/usr/btn[2]").press() 
        time.sleep(1)
        session.findById("wnd[1]/tbar[0]/btn[24]").press() 
        time.sleep(1)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  
        time.sleep(1)
    except Exception as e:
        print(f"   ⚠️ Problema inyectando datos iniciales: {e}")

    # 3. CONFIGURAR CHECKBOXES PARA "MAF" EXCLUSIVAMENTE
    print("   ☑️ Seleccionando casillas exclusivas de MAF...")
    for i in range(7): # Desmarcar todas por seguridad
        try: session.findById(f"wnd[0]/usr/chk[{i}]").selected = False
        except: pass
    
    # Activar Cabecera (0), Operación (1) y MAF (5) tal cual tu VBS
    for i in [0, 1, 5]: 
        try: session.findById(f"wnd[0]/usr/chk[{i}]").selected = True
        except: pass

    # 4. EJECUTAR
    session.findById("wnd[0]").sendVKey(8) 
    print("   ⏳ Cargando tabla ALV (esperando 8 segundos para evitar errores)...")
    time.sleep(8) # Pausa larga asegurando que la tabla renderice
    
    # 5. EXPORTAR (Basado calcado en JM1.vbs y JM2.vbs)
    print("   💾 Iniciando exportación...")
    try:
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        time.sleep(2)
        
        # Seleccionar Hoja de Cálculo
        session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)
        
        # Llenar Rutas Modernas DY_PATH
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta_destino
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{NOMBRE_ARCHIVO}_{FECHA_HOY}.xlsx"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        # Confirmar sobrescritura si existe
        time.sleep(1)
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass
    except Exception as e:
        print(f"   ❌ Error en el flujo de exportación de SAP: {e}")

    # 6. VERIFICACIÓN ROBUSTA EN DISCO
    archivo_completo = os.path.join(ruta_destino, f"{NOMBRE_ARCHIVO}.xlsx")
    print(f"   🔍 Buscando el archivo en Windows: {archivo_completo}")
    
    intentos = 0
    descargado = False
    while intentos < 20: # Espera hasta 20 segundos
        if os.path.exists(archivo_completo):
            descargado = True
            break
        time.sleep(1)
        intentos += 1

    if descargado:
        print("\n✅ ¡ARCHIVO MAF DESCARGADO CON ÉXITO EN TU DISCO DURO!")
    else:
        print("\n⚠️ ALERTA: Pasaron 20 segundos y Windows no detecta el archivo.")

    # 7. PAUSA DE APROBACIÓN HUMANA ANTES DE RETROCEDER
    print("\n" + "="*60)
    input("🛑 PAUSA DE APROBACIÓN: Por favor, ve a la carpeta y verifica visualmente que el archivo exista.\n👉 Presiona ENTER aquí en la consola cuando estés listo para continuar...")
    print("="*60 + "\n")

    # 8. RETORNAR
    print("   🔙 Retornando a la pantalla principal...")
    try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
    except: pass

# ============================================================================
# PUNTO DE ENTRADA (para que SÍ se ejecute al correr "python ia9ia10.py ...")
# ============================================================================
def _aplicar_weekdir_por_arg():
    '''
    Si pasas una ruta tipo ...\2026\08 como 1er argumento,
    el script la usa como SEMANA_DESTINO y recalcula SEMANA_ORIGEN.
    '''
    if len(sys.argv) < 2:
        return
    weekdir = sys.argv[1].strip().strip('"').strip("'")
    if not weekdir:
        return
    if os.path.isdir(weekdir):
        base_dir = os.path.dirname(weekdir)
        semana = os.path.basename(weekdir)
        os.environ["SAP_DATALOG_BASE_DIR"] = base_dir
        os.environ["SAP_DATALOG_SEMANA"] = semana
        os.environ.pop("SAP_DATALOG_SEMANA_ORIGEN", None)  # recalcular

if __name__ == "__main__":
    _aplicar_weekdir_por_arg()
    BASE_DIR, SEMANA_ORIGEN, SEMANA_DESTINO, EXCEL_MAESTRO, FECHA_HOY = _cargar_config_desde_env()
    print(f"📌 BaseDir  : {BASE_DIR}")
    print(f"📌 Semana   : {SEMANA_DESTINO} (origen {SEMANA_ORIGEN})")
    print(f"📌 ExcelLog : {EXCEL_MAESTRO}")
    main()
