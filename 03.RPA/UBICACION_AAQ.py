import time
import os
import sys
import re
from pathlib import Path, PureWindowsPath

import pandas as pd
import win32com.client

# =========================
# CONFIG (solo se usa si NO pasas ruta por terminal)
# =========================
R07 = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026\07")
R06 = R07.parent / "06"

# Ruta de respaldo en caso de fallo (solo si NO pasas ruta por terminal)
RUTA_FALLBACK_08 = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026\08")
CONTROL_XLSX = None

# ✅ NUEVO PARÁMETRO: Fecha para la transacción IA09
FECHA_DIA_FIJADO = "13.03.2026"

# =========================
# CLASE ROUTER MANUAL
# =========================
class ManualRouter:
    """Se usa cuando pasas la ruta por la terminal, omite libro1.xlsx"""
    def __init__(self, ruta_fija: str):
        self.ruta_fija = os.path.abspath(str(ruta_fija).strip())
    
    def get_row_by_index(self, _):
        """Simula devolver un 'row' (fila) donde todo apunta a la ruta que pasaste."""
        class DummyRow:
            def __init__(self, rf):
                self.carpeta = rf
                # Para evitar problemas de atributo no encontrado:
                self.estado = "EN PROCESO"
                self.error = ""
        return DummyRow(self.ruta_fija)
    
    def update_row(self, idx, estado, msg):
        print(f"-> [ESTADO MANUAL] {estado}: {msg}")

# =========================
# FUNCIONES SAP
# =========================
def conectar_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        return session
    except Exception as e:
        print(f"❌ Error al conectar con SAP: {e}")
        return None

def limpiar_string_para_sap(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip()

def guardar_dialogo(session, out_dir, nombre_archivo, fallback_dir=None):
    time.sleep(1)
    # Intenta dar Enter por si sale un popup intermedio
    try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except: pass
    time.sleep(1)
    
    # Intenta poner el nombre del archivo si hay un campo 'txt'
    try: session.findById("wnd[1]/usr/ssub/1/txt").text = nombre_archivo
    except: pass
    
    # Intenta darle al botón 20
    try: session.findById("wnd[1]/tbar[0]/btn[20]").press()
    except: pass
    time.sleep(1)
    
    # Rellenar Directorio y Nombre
    ruta_ingresada = False
    
    # Modo 1: Campos DY_PATH y DY_FILENAME
    try:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = out_dir
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{nombre_archivo}.xlsx"
        ruta_ingresada = True
    except: pass
    
    # Modo 2: ctxt[0] y ctxt[1]
    if not ruta_ingresada:
        try:
            session.findById("wnd[1]/usr/ctxt[0]").text = out_dir
            session.findById("wnd[1]/usr/ctxt[1]").text = f"{nombre_archivo}.xlsx"
            ruta_ingresada = True
        except: pass
    
    # Si logramos meter la ruta y el nombre, damos click en el botón "Generar" (btn[0])
    if ruta_ingresada:
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception as e: print(f"Aviso al presionar btn[0]: {e}")
        time.sleep(2)
        
        # Clics de seguridad por si pide sobrescribir
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass

        # Validamos si se guardó realmente, si no, intentamos con fallback
        archivo_esperado = os.path.join(out_dir, f"{nombre_archivo}.xlsx")
        if fallback_dir and not os.path.exists(archivo_esperado):
            print(f"⚠️ No se encontró {archivo_esperado}. Guardando en carpeta de respaldo: {fallback_dir}")
            try: session.findById("wnd[1]/usr/ctxtDY_PATH").text = fallback_dir
            except: 
                try: session.findById("wnd[1]/usr/ctxt[0]").text = fallback_dir
                except: pass
            try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass
            time.sleep(2)
            try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass
            try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass
            
    else:
        print("⚠️ No se encontraron los campos para ingresar el directorio de guardado.")

# =========================
# PRINCIPAL
# =========================
def main():
    # 1) Definir modo: Manual (terminal) o Automático (libro1.xlsx)
    router = None
    if len(sys.argv) > 1:
        arg = sys.argv[1].strip()
        if arg.startswith("--"):
            arg = arg[2:]
        if arg:
            print(f"🔍 MODO MANUAL: usando ruta -> {arg}")
            router = ManualRouter(arg)
    
    # Conectamos a SAP
    session = conectar_sap()
    if not session:
        return

    session.findById("wnd[0]").maximize()

    # Vamos a iterar. 
    # Si es router, iteramos 1 vez. 
    max_idx = 1
    
    for idx in range(max_idx):
        if router:
            row = router.get_row_by_index(idx)
        else:
            break
            
        print(f"\n==============================================")
        print(f"PROCESANDO (idx={idx}): {row.carpeta}")
        print(f"==============================================")
        
        carpeta_actual = row.carpeta
        # Validar si la carpeta existe. Si no, la creamos (o fallamos, según prefieras).
        if not os.path.isdir(carpeta_actual):
            print(f"⚠️ La carpeta {carpeta_actual} no existe, se creará.")
            try:
                os.makedirs(carpeta_actual, exist_ok=True)
            except Exception as e:
                print(f"❌ Error al crear carpeta: {e}")
                continue
                
        # --- LÓGICA DE IL06 Y IA09 ---
        
        # ==========================================
        # TRANSCACCIÓN 1: IL06
        # ==========================================
        out1 = "IL06_Ubicaciones"
        carpeta1 = os.path.join(carpeta_actual, "Ubicaciones")
        os.makedirs(carpeta1, exist_ok=True)
        fallback_dir = str(RUTA_FALLBACK_08) if not router else carpeta1

        print(f"   -> Guardando {out1} en -> {carpeta1}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nIL06"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            
            try: session.findById("wnd[0]/usr/ctxt[7]").text = ""
            except Exception: pass
            
            try: session.findById("wnd[0]/usr/ctxt[9]").text = "C601"
            except Exception: pass
            
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(1.5)
            
            # Exportar local (Ctrl+Shift+F9 = btn[45] o btn[16])
            try: session.findById("wnd[0]/tbar[1]/btn[45]").press()
            except: session.findById("wnd[0]/tbar[1]/btn[16]").press()
            time.sleep(1)
            
            # Formato Excel local
            try: session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
            except: 
                try: session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select()
                except: pass
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            guardar_dialogo(session, carpeta1, out1, fallback_dir=fallback_dir)
            print("   ✅ IL06 finalizada.")
        except Exception as e:
            print(f"   ❌ Ocurrió un error en IL06: {e}")

        # ==========================================
        # TRANSCACCIÓN 2: IA09 (HOJA DE RUTA)
        # ==========================================
        carpeta3 = os.path.join(carpeta_actual, "Hoja de ruta")
        os.makedirs(carpeta3, exist_ok=True)
        ruta_ia09 = os.path.join(carpeta3, "IA09_CabeceraHojaRuta_AAQ.xlsx")
        
        # ✅ NUEVA LÓGICA DE REANUDACIÓN: Saltar si IA09 ya existe
        if os.path.exists(ruta_ia09) and os.path.getsize(ruta_ia09) > 0:
            print(f"   ⏩ El archivo IA09 ya existe en {carpeta3}. Saltando extracción de SAP...")
        else:
            out3 = "IA09_CabeceraHojaRuta_AAQ"
            print(f"   -> Guardando {out3} en -> {carpeta3}")
            try:
                # Botones de retroceso por si la transacción anterior dejó SAP bloqueado
                try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
                except Exception: pass
                try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
                except Exception: pass

                session.findById("wnd[0]/tbar[0]/okcd").text = "/nIA09"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
                
                # ✅ INYECCIÓN DE FECHA Y CENTRO CON IDs REALES DEL VBS
                try: session.findById("wnd[0]/usr/ctxt[6]").text = FECHA_DIA_FIJADO
                except Exception as e: print(f"⚠️ Error al ingresar fecha en IA09: {e}")
                
                try: session.findById("wnd[0]/usr/ctxt[9]").text = "C601"
                except Exception as e: print(f"⚠️ Error al ingresar centro en IA09: {e}")
                
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                time.sleep(1.5)
                
                session.findById("wnd[0]/tbar[1]/btn[16]").press()
                time.sleep(1)
                
                # Formato Excel
                try:
                    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except: pass
                time.sleep(1)
                
                guardar_dialogo(session, carpeta3, "IA09_CabeceraHojaRuta_AAQ", fallback_dir=fallback_dir)
                print("   ✅ IA09 finalizada correctamente.")
                
            except Exception as e:
                # Si falla la IA09, el script simplemente avisa y continúa
                print(f"   ❌ Ocurrió un error en IA09: {e}")

        # Marcar la fila como completada si todo salió bien
        if router:
            router.update_row(idx, "OK", "Ejecutado correctamente")
            
    print("\nFIN DEL PROCESO.")

if __name__ == "__main__":
    main()