import win32com.client
import pandas as pd
import pyperclip
import time
import sys
import os
import glob
import io
import signal
from datetime import datetime

# ============================================================
# CONFIGURACIÓN
# ============================================================
DEBUG = False
SHORT_SLEEP = 0.20
REPORT_TIMEOUT = 160
EXPORT_TIMEOUT = 45

FECHA_HOY = datetime.now().strftime("%d%m%y")

# ✅ RANGOS DE FECHAS PARA BÚSQUEDA (IW39, IW47, IW13)
FECHA_INICIO = "13.03.2025"
FECHA_FIN = "13.03.2026"

def dprint(msg: str):
    if DEBUG:
        print(msg)

def snooze(mult=1.0):
    time.sleep(SHORT_SLEEP * mult)

# ============================================================
# RUTA BASE
# ============================================================
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

if len(sys.argv) < 2:
    print("❌ ERROR: No ingresaste ninguna ruta en la terminal.")
    sys.exit(1)

RUTA_BASE = " ".join(sys.argv[1:]).strip().strip('"').strip("'")
if RUTA_BASE.startswith("--"):
    RUTA_BASE = RUTA_BASE[2:]
RUTA_BASE = os.path.abspath(RUTA_BASE)

CARPETA_ORDENES = os.path.join(RUTA_BASE, "Ordenes")
CARPETAS = {
    "IW49N": os.path.join(RUTA_BASE, "Operaciones"),
    "IW47":  os.path.join(RUTA_BASE, "notificación"),
    "IW13":  os.path.join(RUTA_BASE, "Material Orden")
}

for ruta in CARPETAS.values():
    os.makedirs(ruta, exist_ok=True)

# ============================================================
# EXCEL IW39 -> ÓRDENES
# ============================================================
def obtener_archivo_ordenes():
    print(f"   🔎 Buscando archivo IW39 en: {CARPETA_ORDENES}")
    # Busca cualquier archivo que contenga "IW39" sin importar si es xls, xlsx
    patron = os.path.join(CARPETA_ORDENES, "*IW39*.*")
    archivos = glob.glob(patron)
    
    # Filtrar solo archivos válidos (excluir PDFs u ocultos temporales como ~$)
    archivos_validos = [f for f in archivos if (f.lower().endswith('.xlsx') or f.lower().endswith('.xls')) and not os.path.basename(f).startswith('~$')]
    
    if archivos_validos:
        # Devuelve el archivo modificado más recientemente
        archivo_encontrado = max(archivos_validos, key=os.path.getmtime)
        print(f"   ✅ Archivo encontrado: {os.path.basename(archivo_encontrado)}")
        return archivo_encontrado
    
    # Si no lo encuentra, mostramos qué hay en la carpeta para auditar
    try:
        print(f"   ❌ No se encontró ningún archivo IW39 válido. Archivos presentes: {os.listdir(CARPETA_ORDENES)}")
    except Exception:
        print("   ❌ La carpeta aún no existe o está vacía.")
    return None

def detectar_columna_orden(df: pd.DataFrame):
    for c in df.columns:
        if str(c).strip().lower() == "orden": return c
    for c in df.columns:
        if "orden" in str(c).strip().lower(): return c
    return None

def cargar_ordenes_desde_excel(ruta_archivo) -> list[str]:
    df = pd.read_excel(ruta_archivo)
    col = detectar_columna_orden(df)
    if col is not None:
        serie = df[col].dropna().astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    else:
        if len(df.columns) < 6:
            raise RuntimeError("El archivo IW39 no tiene columna suficiente.")
        serie = df.iloc[:, 5].dropna().astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

    ordenes = [x for x in serie.tolist() if x.isdigit()]
    seen = set()
    out = []
    for o in ordenes:
        if o not in seen:
            seen.add(o)
            out.append(o)
    if not out:
        raise RuntimeError("No se encontraron órdenes válidas en IW39.")
    return out

def copiar_ordenes(ordenes: list[str]):
    pyperclip.copy("\r\n".join(ordenes))
    print(f"📋 {len(ordenes)} órdenes copiadas al portapapeles.")

# ============================================================
# SAP helpers
# ============================================================
def conectar_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        return app.Children(0).Children(0)
    except Exception as e:
        print(f"❌ ERROR: SAP GUI no detectado. {e}")
        sys.exit(1)

def exists(session, id_str: str) -> bool:
    try:
        return session.findById(id_str, False) is not None
    except Exception:
        return False

def get(session, id_str: str):
    return session.findById(id_str)

def wait_ready(session, timeout=60):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if not session.Busy: return True
        except: pass
        time.sleep(0.2)
    return False

def cerrar_popups(session, max_cierres=10):
    for _ in range(max_cierres):
        snooze(1.2)
        try:
            if exists(session, "wnd[1]"):
                if exists(session, "wnd[1]/tbar[0]/btn[0]"):
                    try: get(session, "wnd[1]/tbar[0]/btn[0]").press(); continue
                    except: pass
                try: get(session, "wnd[1]").close(); continue
                except: pass
        except: pass
        break

def resetear_pantalla_sap(session):
    cerrar_popups(session, 10)
    try:
        get(session, "wnd[0]/tbar[0]/okcd").text = "/n"
        get(session, "wnd[0]").sendVKey(0)
        wait_ready(session, 30)
    except: pass

def set_checks_0_9(session):
    for i in range(0, 10):
        try: get(session, f"wnd[0]/usr/chk[{i}]").selected = True
        except: pass

def pegar_en_matchcode(session, btn_candidates: list[str], ordenes: list[str]) -> bool:
    copiar_ordenes(ordenes)
    for btn_id in btn_candidates:
        try:
            if not exists(session, btn_id): continue
            get(session, btn_id).press()
            snooze(5)
            if not exists(session, "wnd[1]"): continue
            try: get(session, "wnd[1]/tbar[0]/btn[25]").press()
            except: pass
            snooze(2)
            try: get(session, "wnd[1]/tbar[0]/btn[24]").press()
            except: get(session, "wnd[1]/tbar[0]/btn[23]").press()
            snooze(2)
            try: get(session, "wnd[1]/tbar[0]/btn[8]").press()
            except: get(session, "wnd[1]/tbar[0]/btn[0]").press()
            snooze(2)
            cerrar_popups(session, 3)
            return True
        except: continue
    return False

def sbar_text(session) -> str:
    try: return get(session, "wnd[0]/sbar").Text
    except: return ""

def no_hay_datos(session) -> bool:
    txt = (sbar_text(session) or "").lower()
    claves = ["no se ha seleccionado", "no se encontraron", "sin datos", "no hay datos", "ningún registro"]
    return any(k in txt for k in claves)

# ============================================================
# FUNCIÓN DE EXPORTACIÓN NATIVA CON DEPURACIÓN EXTREMA (PARA IW49N)
# ============================================================
def exportar_excel_con_debug(session, carpeta, base_name_sin_ext):
    print(f"\n   [DEBUG] 🛠️ INICIANDO DIAGNÓSTICO DE GUARDADO EN SAP...")
    os.makedirs(carpeta, exist_ok=True)
    
    export_opened = False
    for btn in ["wnd[0]/tbar[1]/btn[45]", "wnd[0]/tbar[1]/btn[43]", "wnd[0]/tbar[1]/btn[16]"]:
        if exists(session, btn):
            print(f"   [DEBUG] Botón de exportación encontrado: {btn}")
            get(session, btn).press()
            export_opened = True
            break
            
    if not export_opened:
        try:
            get(session, "wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
            print("   [DEBUG] Menú de exportación encontrado por barra de menú.")
            export_opened = True
        except: pass

    if not export_opened:
        print("   [DEBUG] ❌ FALLO: No se encontró forma de abrir el diálogo de exportación.")
        return False

    time.sleep(4)

    if exists(session, "wnd[1]/usr/sub/2/sub/2/1/rad[1,0]"):
        print("   [DEBUG] Formato de hoja de cálculo detectado en rad[1,0].")
        get(session, "wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select()
        get(session, "wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)
    elif exists(session, "wnd[1]/usr/sub/2/sub/2/1/rad[2,0]"):
        print("   [DEBUG] Formato de hoja de cálculo detectado en rad[2,0].")
        get(session, "wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
        get(session, "wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)

    print("   [DEBUG] Evaluando ventana actual para asignar nombre y ruta...")
    guardado_ok = False
    
    if exists(session, "wnd[1]/usr/ctxtDY_PATH") and exists(session, "wnd[1]/usr/ctxtDY_FILENAME"):
        print("   [DEBUG] ✅ Variante A: (ctxtDY_PATH / ctxtDY_FILENAME)")
        try:
            get(session, "wnd[1]/usr/ctxtDY_PATH").text = carpeta
            get(session, "wnd[1]/usr/ctxtDY_FILENAME").text = f"{base_name_sin_ext}.xlsx"
            get(session, "wnd[1]/tbar[0]/btn[0]").press()
            guardado_ok = True
        except Exception as e:
            print(f"   [DEBUG] ❌ Error en A: {e}")

    elif exists(session, "wnd[1]/usr/ctxt[0]") and exists(session, "wnd[1]/usr/ctxt[1]"):
        print("   [DEBUG] ✅ Variante B: (ctxt[0] / ctxt[1])")
        try:
            get(session, "wnd[1]/usr/ctxt[0]").text = carpeta
            get(session, "wnd[1]/usr/ctxt[1]").text = f"{base_name_sin_ext}.xlsx"
            get(session, "wnd[1]/tbar[0]/btn[0]").press()
            guardado_ok = True
        except Exception as e:
            print(f"   [DEBUG] ❌ Error en B: {e}")

    elif exists(session, "wnd[1]/usr/ssub/1/txt"):
        print("   [DEBUG] ✅ Variante C: Ventana en 2 pasos")
        try:
            get(session, "wnd[1]/usr/ssub/1/txt").text = base_name_sin_ext
            get(session, "wnd[1]/tbar[0]/btn[20]").press()
            time.sleep(2)
            if exists(session, "wnd[1]/usr/ctxt[0]"):
                get(session, "wnd[1]/usr/ctxt[0]").text = carpeta
                get(session, "wnd[1]/tbar[0]/btn[0]").press()
                guardado_ok = True
            else:
                print("   [DEBUG] ❌ FALLO: No apareció la ruta.")
        except Exception as e:
            print(f"   [DEBUG] ❌ Error en C: {e}")
            
    else:
        print("   [DEBUG] ❌ FALLO CRÍTICO: SAP abrió ventana irreconocible.")

    time.sleep(2)
    try: 
        if exists(session, "wnd[2]/tbar[0]/btn[0]"):
            get(session, "wnd[2]/tbar[0]/btn[0]").press()
    except: pass

    time.sleep(4)
    print(f"   [DEBUG] 🛠️ FIN DE DIAGNÓSTICO DE GUARDADO.\n")
    return guardado_ok

def verificar_archivo_creado(ruta_carpeta, nombre_base, timeout=EXPORT_TIMEOUT):
    print(f"   ⏳ Verificando descarga en: {ruta_carpeta} ...")
    t0 = time.time()
    while time.time() - t0 < timeout:
        for ext in (".xlsx", ".xls", ".XLSX", ".XLS"):
            p = os.path.join(ruta_carpeta, nombre_base + ext)
            if os.path.exists(p) and not os.path.basename(p).startswith("~$"):
                print("   ✅ ¡Archivo detectado!\n")
                return True
        time.sleep(1.5)
    print(f"   ⚠️ ALERTA: El archivo no apareció o tardó más de {timeout} seg.\n")
    return False
# ============================================================
# EXTRACCIÓN ROBUSTA BLINDADA Y ORDENADA (PARA IW47 e IW13)
# ============================================================
def copiar_y_guardar_grilla(session, grid_id, ruta_completa, cabeceras=None):
    try:
        grid = get(session, grid_id)
        
        # 1. EXTRAER CABECERAS PRIMERO (Para no colapsar la selección luego)
        if cabeceras is None:
            cabeceras = []
            try:
                for col in grid.ColumnOrder:
                    texto_col = str(col)
                    try:
                        titulos = grid.GetColumnTitles(col)
                        for t in reversed(titulos):
                            if t and str(t).strip():
                                texto_col = str(t).strip()
                                break
                    except: pass
                    cabeceras.append(texto_col)
            except: pass

        # 2. PREPARAR VENTANA Y PORTAPAPELES
        shell = win32com.client.Dispatch("WScript.Shell")
        pyperclip.copy("") 
        
        try:
            titulo_sap = get(session, "wnd[0]").Text
            shell.AppActivate(titulo_sap)
        except: 
            try: shell.AppActivate("SAP") 
            except: pass
            
        time.sleep(1.5)
        
        # 3. SELECCIONAR TODO (En el último segundo)
        grid.setFocus()
        grid.setCurrentCell(-1, "")
        grid.selectAll()
        time.sleep(1) # Damos 1 segundo entero para que SAP sombree la tabla
        
        print("   📋 Enviando comando de copia (CTRL+C) a SAP...")
        
        # === BLINDAJE ABSOLUTO CONTRA KEYBOARD INTERRUPT ===
        original_sigint = signal.getsignal(signal.SIGINT)
        signal.signal(signal.SIGINT, signal.SIG_IGN)
        
        shell.SendKeys("^c")
        
        # Esperamos los datos
        datos_copiados = ""
        for i in range(20):
            time.sleep(1)
            datos_copiados = pyperclip.paste()
            if datos_copiados.strip():
                print(f"   ⚡ ¡Datos recibidos en memoria después de {i+1} segs!")
                break
                
        # REINTENTO CON MÉTODO ALTERNATIVO (CTRL + INSERT)
        if not datos_copiados.strip():
            print("   ⚠️ El portapapeles sigue vacío. Forzando foco y reintentando con Ctrl+Insert...")
            try: shell.AppActivate(titulo_sap)
            except: pass
            time.sleep(1)
            
            grid.setFocus()
            grid.selectAll()
            time.sleep(1)
            shell.SendKeys("^{INSERT}") # Atajo alternativo infalible de Windows
            
            for i in range(10):
                time.sleep(1)
                datos_copiados = pyperclip.paste()
                if datos_copiados.strip():
                    print(f"   ⚡ ¡Datos recibidos en el reintento!")
                    break

        # Restauramos el comportamiento normal de Ctrl+C en la terminal
        signal.signal(signal.SIGINT, original_sigint)
        # ===================================================

        # === NUEVA PROTECCIÓN ANTI-CRASH DE PANDAS ===
        if not datos_copiados.strip():
            print("   ❌ ERROR: SAP bloqueó el copiado o se perdió el foco de la ventana. No hay datos para procesar.")
            return False # Retornamos False aquí para que el script siga vivo y pase a la siguiente TX
        # ===================================================

        print("   💾 Convirtiendo datos y guardando Excel...")
        df = pd.read_csv(io.StringIO(datos_copiados), sep='\t', header=None, dtype=str)
        df.dropna(how='all', axis=1, inplace=True)
        
        if cabeceras and len(cabeceras) == len(df.columns): df.columns = cabeceras
        elif cabeceras and len(cabeceras) > len(df.columns): df.columns = cabeceras[:len(df.columns)]
        else: df.columns = [f"Col_{i+1}" for i in range(len(df.columns))]
        
        df.to_excel(ruta_completa, index=False)
        print(f"   ✅ ¡ÉXITO! Archivo guardado en: {ruta_completa}")
        return True
        
    except Exception as e:
        try: signal.signal(signal.SIGINT, original_sigint) 
        except: pass
        print(f"   ❌ Falló el proceso de copia rápida: {e}")
        return False

# ============================================================
# TRANSACCIONES
# ============================================================
def ejecutar_iw49n(session, ordenes: list[str]) -> bool:
    print("🚀 Ejecutando IW49N...")
    resetear_pantalla_sap(session)
    get(session, "wnd[0]/tbar[0]/okcd").text = "/nIW49N"
    get(session, "wnd[0]").sendVKey(0)
    wait_ready(session, 60)
    
    set_checks_0_9(session)
    try:
        get(session, "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select()
        get(session, "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub/1/2/ctxt[5]").text = "MCARRANZA03"
        get(session, "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").select()
    except: pass

    print("   🎯 Inyectando órdenes IW49N...")
    if not pegar_en_matchcode(session, ["wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub/1/2/btn[0]", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH"], ordenes):
        return False

    get(session, "wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session, REPORT_TIMEOUT)
    cerrar_popups(session, 6)

    if no_hay_datos(session):
        print("   ⚠️ IW49N sin datos.")
        return False

    print("   💾 Exportando IW49N...")
    return exportar_excel_con_debug(session, CARPETAS["IW49N"], f"IW7N_IW49N_Operaciones_{FECHA_HOY}")


def ejecutar_iw47(session, ordenes: list[str]) -> bool:
    print("🚀 Ejecutando IW47 (Exportación rápida vía Portapapeles)...")
    resetear_pantalla_sap(session)
    get(session, "wnd[0]").maximize()
    get(session, "wnd[0]/tbar[0]/okcd").text = "/nIW47"
    get(session, "wnd[0]").sendVKey(0)
    wait_ready(session, 60)

    print("   🎯 Inyectando órdenes IW47...")
    if not pegar_en_matchcode(session, ["wnd[0]/usr/btn[0]"], ordenes):
        return False

    set_checks_0_9(session)
    
    # ✅ INYECTANDO VARIABLES DE FECHA DE BÚSQUEDA (Desde el config)
    print(f"   ⚙️ Configurando Fechas: {FECHA_INICIO} a {FECHA_FIN}")
    try: get(session, "wnd[0]/usr/ctxt[15]").text = FECHA_INICIO
    except: pass
    try: get(session, "wnd[0]/usr/ctxt[16]").text = FECHA_FIN
    except: pass
    
    try: get(session, "wnd[0]/usr/ctxt[43]").text = "MCARRANZA01"
    except: pass

    get(session, "wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session, REPORT_TIMEOUT)
    cerrar_popups(session, 6)

    if no_hay_datos(session):
        print("   ⚠️ IW47 sin datos.")
        return False

    ruta = os.path.join(CARPETAS["IW47"], f"IW47_notificaciones_{FECHA_HOY}.xlsx")
    return copiar_y_guardar_grilla(session, "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell", ruta, cabeceras=None)


def ejecutar_iw13(session, ordenes: list[str]) -> bool:
    print("🚀 Ejecutando IW13 (Exportación rápida vía Portapapeles)...")
    resetear_pantalla_sap(session)
    get(session, "wnd[0]").maximize()
    get(session, "wnd[0]/tbar[0]/okcd").text = "/nIW13"
    get(session, "wnd[0]").sendVKey(0)
    wait_ready(session, 60)

    set_checks_0_9(session)
    print("   🎯 Inyectando órdenes IW13...")
    if not pegar_en_matchcode(session, ["wnd[0]/usr/btn[1]"], ordenes):
        return False

    # ✅ INYECTANDO FECHA INICIAL EN IW13 (Desde el config)
    print(f"   ⚙️ Configurando Día Fijado: {FECHA_INICIO}")
    try: get(session, "wnd[0]/usr/ctxt[18]").text = FECHA_INICIO
    except: pass
    
    try: get(session, "wnd[0]/usr/ctxt[20]").text = "CARRANZA02"
    except: pass

    get(session, "wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session, REPORT_TIMEOUT)
    cerrar_popups(session, 6)

    if no_hay_datos(session):
        print("   ⚠️ IW13 sin datos.")
        return False

    cabeceras_iw13 = [
        "Orden", "Material", "Descripción del material", "UM", 
        "Reservado", "Ref.reser.", "S/ref.rese", "Fe.contab.", "TpP", "Tomados"
    ]

    ruta = os.path.join(CARPETAS["IW13"], f"IW13_MaterialesOrden_{FECHA_HOY}.xlsx")
    return copiar_y_guardar_grilla(session, "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell", ruta, cabeceras=cabeceras_iw13)

# ============================================================
# MAIN BLINDADO
# ============================================================
def main():
    print("\n" + "="*60)
    print("🤖 INICIANDO AUTOMATIZACIÓN SAP (BLINDAJE ANTIMUERTE Y DIAGNÓSTICO)")
    print(f"📁 Directorio Base: {RUTA_BASE}")
    print("="*60)

    ruta_iw39 = obtener_archivo_ordenes()
    if not ruta_iw39:
        print("❌ ERROR: No se encontró el archivo IW39_CabeceraOrden en /Ordenes.")
        sys.exit(1)

    try:
        ordenes = cargar_ordenes_desde_excel(ruta_iw39)
    except Exception as e:
        print(f"❌ ERROR: No se pudo leer órdenes desde IW39: {e}")
        sys.exit(1)

    session = conectar_sap()

    # --- 1) IW49N ---
    nombre_iw49n = f"IW7N_IW49N_Operaciones_{FECHA_HOY}"
    print("\n" + "-"*40)
    print(f"📋 Ejecutando IW49N: {nombre_iw49n}")
    if not verificar_archivo_creado(CARPETAS["IW49N"], nombre_iw49n, timeout=2):
        try:
            ok49n = ejecutar_iw49n(session, ordenes)
            if not ok49n or not verificar_archivo_creado(CARPETAS["IW49N"], nombre_iw49n):
                print("   ⚠️ ADVERTENCIA: IW49N no se exportó/verificó correctamente. Saltando a IW47...")
            else:
                print("   ✅ IW49N finalizada correctamente.")
        except Exception as e:
            print(f"   ❌ ERROR CRÍTICO en IW49N: {e}. Saltando a IW47...")
            try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except: pass
    else: 
        print("   ⏩ IW49N ya existe. Saltando...")

    # --- 2) IW47 ---
    nombre_iw47 = f"IW47_notificaciones_{FECHA_HOY}"
    print("\n" + "-"*40)
    print(f"📋 Ejecutando IW47: {nombre_iw47}")
    if not verificar_archivo_creado(CARPETAS["IW47"], nombre_iw47, timeout=2):
        try:
            ok47 = ejecutar_iw47(session, ordenes)
            if not ok47:
                print("   ⚠️ ADVERTENCIA: IW47 reportó fallo en ejecución. Saltando a IW13...")
            else:
                print("   ✅ IW47 finalizada o ejecutada.")
                verificar_archivo_creado(CARPETAS["IW47"], nombre_iw47)
        except Exception as e:
            print(f"   ❌ ERROR CRÍTICO en IW47: {e}. Saltando a IW13...")
            try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except: pass
    else: 
        print("   ⏩ IW47 ya existe. Saltando...")

    # --- 3) IW13 ---
    nombre_iw13 = f"IW13_Material Orden_{FECHA_HOY}"
    print("\n" + "-"*40)
    print(f"📋 Ejecutando IW13: {nombre_iw13}")
    if not verificar_archivo_creado(CARPETAS["IW13"], nombre_iw13, timeout=2):
        try:
            ok13 = ejecutar_iw13(session, ordenes)
            if not ok13:
                print("   ⚠️ ADVERTENCIA: IW13 reportó fallo en ejecución.")
            else:
                print("   ✅ IW13 finalizada o ejecutada.")
                verificar_archivo_creado(CARPETAS["IW13"], nombre_iw13)
        except Exception as e:
            print(f"   ❌ ERROR CRÍTICO en IW13: {e}.")
            try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except: pass
    else: 
        print("   ⏩ IW13 ya existe. Saltando...")

    print("\n" + "="*60)
    print("🏁 PROCESO COMPLETADO")
    print("="*60)

if __name__ == "__main__":
    main()