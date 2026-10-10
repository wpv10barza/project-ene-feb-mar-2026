import win32com.client
import pandas as pd
import pyperclip
import time
import sys
import os
import glob
from datetime import datetime

# Asegurar codificación correcta en consola
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# ============================================================================
# 1. CONFIGURACIÓN DE LOTES Y MATCHCODES
# ============================================================================
MATCHCODE_IP16 = "wnd[0]/usr/btn[0]"  # <- Botón múltiple para IP16
MATCHCODE_IP24 = "wnd[0]/usr/btn[3]"  # <- Botón múltiple para IP24
TAMANO_LOTE = 1000                    # <- Límite de registros por lote

# ============================================================================
# 2. GESTIÓN DE RUTA DESDE LA TERMINAL
# ============================================================================
if len(sys.argv) < 2:
    print("❌ ERROR: No ingresaste ninguna ruta en la terminal.")
    sys.exit(1)

RUTA_BASE = " ".join(sys.argv[1:]).strip('\"').strip('\'')
if RUTA_BASE.startswith("--"):
    RUTA_BASE = RUTA_BASE[2:]

RUTA_BASE = os.path.abspath(RUTA_BASE)

CARPETAS = {
    "IP18": os.path.join(RUTA_BASE, "Posiciones"),
    "IP16": os.path.join(RUTA_BASE, "Planes"),
    "IP24": os.path.join(RUTA_BASE, "Programacion")
}

for ruta in CARPETAS.values():
    os.makedirs(ruta, exist_ok=True)

FECHA_HOY = datetime.now().strftime("%d%m%Y")

print("\n" + "="*60)
print("🤖 INICIANDO AUTOMATIZACIÓN SAP POR LOTES (IP18 -> IP16 -> IP24)")
print(f"📁 Directorio Base: {RUTA_BASE}")
print("="*60)

# ============================================================================
# FUNCIONES DE CONEXIÓN (100% ORIGINAL TUYA) Y SUPERVISIÓN
# ============================================================================
def conectar_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        return app.Children(0).Children(0)
    except Exception:
        print("❌ ERROR: No se detectó SAP GUI abierto.")
        sys.exit(1)

def archivo_ya_existe(ruta_carpeta, codigo_tx):
    archivos = glob.glob(os.path.join(ruta_carpeta, f"*{codigo_tx}*.xlsx"))
    archivos = [f for f in archivos if not os.path.basename(f).startswith("~$") and "Lote" not in f]
    return len(archivos) > 0

def obtener_ruta_archivo_creado(ruta_carpeta, codigo_tx):
    archivos = glob.glob(os.path.join(ruta_carpeta, f"*{codigo_tx}*.xlsx"))
    archivos = [f for f in archivos if not os.path.basename(f).startswith("~$") and "Lote" not in f]
    if archivos: return max(archivos, key=os.path.getmtime)
    return None

def verificar_archivo_creado(ruta_carpeta, codigo_tx, timeout=4000):
    print(f"   ⏳ Verificando si existe el reporte {codigo_tx} en: {ruta_carpeta} ...")
    inicio = time.time()
    while time.time() - inicio < timeout:
        ruta = obtener_ruta_archivo_creado(ruta_carpeta, codigo_tx)
        if ruta:
            print(f"   ✅ ¡Archivo {codigo_tx} detectado!\n")
            return ruta
        time.sleep(1.5)
    print(f"   ⚠️ ALERTA: El archivo {codigo_tx} no apareció.\n")
    return None

def esperar_archivo_escrito(ruta_archivo, timeout=60):
    print(f"   ⏳ Esperando escritura de disco para evitar 0 bytes en {os.path.basename(ruta_archivo)}...")
    inicio = time.time()
    while time.time() - inicio < timeout:
        if os.path.exists(ruta_archivo):
            if os.path.getsize(ruta_archivo) > 0:
                try:
                    with open(ruta_archivo, 'a'): pass
                    return True
                except IOError: pass
        time.sleep(1)
    print("   ⚠️ Tiempo agotado. El archivo podría estar corrupto o bloqueado.")
    return False

# ============================================================================
# LÓGICA DE LOTES Y MATCHCODE (USO DE TECLADO NATIVO)
# ============================================================================
def pegar_en_matchcode(session, btn_id):
    try:
        session.findById(btn_id).press()
        time.sleep(1)
        
        try: session.findById("wnd[1]").sendVKey(16)
        except: pass
        
        try: session.findById("wnd[1]").sendVKey(24)
        except: pass
        time.sleep(1)
        
        try: session.findById("wnd[1]").sendVKey(8)
        except: pass
        time.sleep(1)
        
        return True
    except Exception as e:
        print(f"   ❌ Error en el menú de pegado (Matchcode): {e}")
        return False

def chunk_list(lista, n):
    return [lista[i:i + n] for i in range(0, len(lista), n)]

def unir_excels_temporales(carpeta, patron_busqueda, nombre_final):
    print(f"   🔄 Uniendo archivos temporales '{patron_busqueda}' en {carpeta}...")
    
    # 1. Buscar archivos temporales (Acepta .xlsx y .xls por si SAP cambia el formato)
    archivos = glob.glob(os.path.join(carpeta, f"*{patron_busqueda}*.xlsx"))
    archivos.extend(glob.glob(os.path.join(carpeta, f"*{patron_busqueda}*.xls")))
    
    # 2. 🔴 CORRECCIÓN CLAVE: Solo excluir el archivo final EXACTO, no los lotes
    nombre_exacto_xlsx = f"{nombre_final}.xlsx".lower()
    nombre_exacto_xls = f"{nombre_final}.xls".lower()
    
    archivos_validos = []
    for f in archivos:
        nombre_base = os.path.basename(f).lower()
        # Ignorar archivos ocultos y el archivo final si ya existe
        if not nombre_base.startswith("~$") and nombre_base != nombre_exacto_xlsx and nombre_base != nombre_exacto_xls:
            archivos_validos.append(f)

    if not archivos_validos:
        print(f"   ⚠️ No se encontraron archivos temporales válidos para unir ({patron_busqueda}).")
        return False

    # 3. Leer y Unir
    dataframes = []
    for f in archivos_validos:
        try:
            df = pd.read_excel(f)
            dataframes.append(df)
            print(f"      📄 Leído con éxito: {os.path.basename(f)}")
        except Exception as e:
            print(f"      ❌ Error al leer {os.path.basename(f)}: {e}")

    # 4. Guardar archivo consolidado
    if dataframes:
        print("   🛠️ Concatenando datos...")
        df_final = pd.concat(dataframes, ignore_index=True)
        ruta_final = os.path.join(carpeta, f"{nombre_final}.xlsx")

        try:
            with pd.ExcelWriter(ruta_final, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            print(f"   ✅ Archivo consolidado guardado como: {nombre_final}.xlsx")

            # 5. Limpieza: Borrar los lotes temporales que ya usamos
            for f in archivos_validos:
                try:
                    os.remove(f)
                    print(f"      🗑️ Borrado temporal: {os.path.basename(f)}")
                except:
                    pass
            return True
            
        except Exception as e:
            print(f"   ❌ Error al guardar el archivo consolidado: {e}")
            return False
    else:
        return False
# ============================================================================
# LÓGICA DE NEGOCIO: FILTRADO Y EXTRACCIÓN
# ============================================================================
def leer_y_filtrar_ip18(ruta_archivo):
    print(f"   📂 Leyendo archivo IP18: {ruta_archivo}")
    try:
        # 1. Intenta leer como Excel normal
        try:
            df = pd.read_excel(ruta_archivo)
        except Exception:
            # 2. Si SAP lo guardó como HTML disfrazado de Excel, usamos read_html
            print(f"   ⚠️ Formato nativo SAP detectado. Leyendo estructura alternativa...")
            df_list = pd.read_html(ruta_archivo, decimal=',', thousands='.')
            df = df_list[0]
        
        if df.empty:
            print("   ⚠️ El archivo fue leído pero está vacío.")
            return None

        print(f"   ✅ Filas totales leídas: {len(df)}")
        
        # Búsqueda flexible de la columna de Ubicación (por si SAP cambia el nombre)
        col_ubi = next((col for col in df.columns if 'ubicac' in str(col).lower() or 'ubic.' in str(col).lower()), None)
        
        if col_ubi:
            df_filtrado = df[df[col_ubi].astype(str).str.startswith('A-AAQ', na=False)]
            print(f"   ✅ Filas retenidas tras filtro 'A-AAQ' en '{col_ubi}': {len(df_filtrado)}")
        else:
            print(f"   ⚠️ No se encontró la columna de 'Ubicación técnica'. Se usarán todas las filas.")
            df_filtrado = df.copy()

        if df_filtrado.empty:
            print("   🛑 La tabla quedó vacía tras el filtro. Se saltarán los procesos dependientes.")
            return None
            
        return df_filtrado

    except Exception as e:
        print(f"   ❌ ERROR crítico al leer el archivo IP18: {e}")
        return None

def extraer_planes_ip16(df_filtrado):
    print("   ⚙️ Extrayendo PLANES (Col. B) para la IP16...")
    try:
        columna_b = df_filtrado.iloc[:, 1].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist()
        planes_unicos = list(dict.fromkeys(columna_b))
        print(f"   ✅ Se extrajeron {len(planes_unicos)} PLANES únicos.")
        return planes_unicos
    except Exception: return []

def extraer_posiciones_ip24(df_filtrado):
    print("   ⚙️ Extrayendo POSICIONES DE MANTENIMIENTO (Col. C) para la IP24...")
    try:
        columna_c = df_filtrado.iloc[:, 2].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist()
        posiciones_unicas = list(dict.fromkeys(columna_c))
        print(f"   ✅ Se extrajeron {len(posiciones_unicas)} POSICIONES únicas.")
        return posiciones_unicas
    except Exception: return []

# ============================================================================
# EXPORTACIÓN
# ============================================================================
def exportar_excel(session, carpeta, nombre_archivo):
    print(f"   💾 Intentando guardar: {nombre_archivo}")
    time.sleep(2)
    try:
        try: session.findById("wnd[0]/tbar[1]/btn[16]").press()
        except: session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select() 
        time.sleep(2.5)
        
        try:
            session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select()
            session.findById("wnd[1]").sendVKey(0) 
            time.sleep(1.5)
        except: pass
        
        try:
            session.findById("wnd[1]/usr/ssub/1/txt").text = nombre_archivo
            session.findById("wnd[1]").sendVKey(0) 
            time.sleep(1.5)
        except: pass

        try:
            session.findById("wnd[1]/usr/ctxt[0]").text = carpeta
            session.findById("wnd[1]/usr/ctxt[1]").text = nombre_archivo
        except: pass

        try: session.findById("wnd[1]").sendVKey(0) 
        except: pass
        
        try: session.findById("wnd[1]/tbar[0]/btn[11]").press() 
        except: pass
        
        time.sleep(1.5)
        print("   ⏳ Dando instrucción a disco...")
        time.sleep(4) 
    except Exception as e:
        print(f"   ❌ Fallo en la secuencia de exportación: {e}")
        return False
    return True

# ============================================================================
# EJECUCIÓN DE TRANSACCIONES 
# ============================================================================
def ejecutar_ip18(session):
    print("\n🚀 Iniciando IP18 (Descarga de Base Maestra)...")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nip18"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        
        try: session.findById("wnd[0]/usr/ctxt[16]").text = "C601"
        except: pass
        try: session.findById("wnd[0]/usr/ctxt[80]").text = "MCARRANZA03"
        except: pass
        
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(5)
        
        exportar_excel(session, CARPETAS["IP18"], f"IP18_Posiciones_{FECHA_HOY}.xlsx")
    except Exception as e:
        print(f"   ❌ Error en IP18: {e}")

def ejecutar_ip16_por_lotes(session, lotes):
    print(f"\n🚀 Iniciando IP16 (En {len(lotes)} lotes para no saturar SAP)...")
    for idx, lote in enumerate(lotes, 1):
        nombre_lote = f"IP16_Planes_Lote_{idx}_{FECHA_HOY}.xlsx"
        print(f"  Procesando lote {idx}/{len(lotes)} ({len(lote)} planes)...")
        pyperclip.copy("\r\n".join(lote))
        
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nip16"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)

            pegar_en_matchcode(session, MATCHCODE_IP16) 
            
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(5)
            
            exportar_excel(session, CARPETAS["IP16"], nombre_lote)
            esperar_archivo_escrito(os.path.join(CARPETAS["IP16"], nombre_lote))
            
        except Exception as e:
            print(f"   ❌ Error en lote {idx} de IP16: {e}")

def ejecutar_ip24_por_lotes(session, lotes):
    print(f"\n🚀 Iniciando IP24 (En {len(lotes)} lotes para no saturar SAP)...")
    
    archivos_temporales = [] 

    for idx, lote in enumerate(lotes, 1):
        nombre_lote = f"IP24_Programacion_Lote_{idx}_{FECHA_HOY}.xlsx"
        ruta_guardada = os.path.join(CARPETAS["IP24"], nombre_lote)
        
        # ==========================================================
        # ✅ NUEVA LÓGICA: Saltar descarga si el lote ya existe
        # ==========================================================
        if os.path.exists(ruta_guardada) and os.path.getsize(ruta_guardada) > 0:
            print(f"  El lote {idx}/{len(lotes)} ya está descargado ({nombre_lote}). Saltando SAP...")
            archivos_temporales.append(ruta_guardada)
            continue  # <-- Esto hace que salte directo al siguiente lote sin tocar SAP
        
        # Si no existe, entonces sí ejecuta SAP
        print(f"  Procesando lote {idx}/{len(lotes)} ({len(lote)} posiciones)...")
        pyperclip.copy("\r\n".join(lote))
        
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nip24"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            
            pegar_en_matchcode(session, MATCHCODE_IP24) 
            
            try: session.findById("wnd[0]/usr/ctxt[90]").text = "MCARRNZA01" 
            except: pass

            session.findById("wnd[0]").sendVKey(8)
            time.sleep(5)
            
            exportar_excel(session, CARPETAS["IP24"], nombre_lote)
            esperar_archivo_escrito(ruta_guardada)
            
            archivos_temporales.append(ruta_guardada)
            
        except Exception as e:
            print(f"   ❌ Error en lote {idx} de IP24: {e}")

    # ==========================================================
    # LÓGICA DE UNIFICACIÓN DE BLOQUES
    # ==========================================================
    print("\n" + "="*60)
    print("🧩 UNIFICANDO BLOQUES DE IP24 EN UN SOLO ARCHIVO MAESTRO...")
    
    dataframes = []
    for archivo in archivos_temporales:
        if os.path.exists(archivo):
            try:
                df_temp = pd.read_excel(archivo)
                dataframes.append(df_temp)
            except Exception as e:
                print(f"    No se pudo leer el temporal {archivo}: {e}")
            
    if dataframes:
        df_maestro = pd.concat(dataframes, ignore_index=True)
        df_maestro.drop_duplicates(inplace=True)
        
        archivo_final = os.path.join(CARPETAS["IP24"], f"IP24_Maestro_Completo_{FECHA_HOY}.xlsx")
        df_maestro.to_excel(archivo_final, index=False)
        print(f"    Archivo final consolidado creado: {os.path.basename(archivo_final)}")
        
        for archivo in archivos_temporales:
            try:
                os.remove(archivo)
            except Exception as e:
                print(f"    No se pudo borrar el temporal {archivo}: {e}")
        print("   🧹 Archivos temporales eliminados correctamente.")
    else:
        print("    Error: No se pudo unificar. No se extrajeron datos en los lotes temporales.")
    def ejecutar_ip24_por_lotes(session, lista_completa, carpeta_destino):
        tamaño_bloque = 200
        print(f"\n🚀 Iniciando IP24 (En bloques de {tamaño_bloque} para no saturar SAP)...")
        
        archivos_temporales = []
        total_bloques = (len(lista_completa) // tamaño_bloque) + 1

        for i in range(0, len(lista_completa), tamaño_bloque):
            bloque_actual = lista_completa[i : i + tamaño_bloque]
            num_bloque = (i // tamaño_bloque) + 1
            nombre_lote = f"IP24_Temp_Bloque_{num_bloque}_{FECHA_HOY}.xlsx"
            ruta_lote = os.path.join(carpeta_destino, nombre_lote)
            
            print(f"   Procesando bloque {num_bloque}/{total_bloques} ({len(bloque_actual)} posiciones)...")
            pyperclip.copy("\r\n".join(bloque_actual))
            
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nip24"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
                
                pegar_en_matchcode(session, MATCHCODE_IP24) 
                
                try: session.findById("wnd[0]/usr/ctxt[90]").text = "MCARRNZA01" 
                except: pass

                session.findById("wnd[0]").sendVKey(8)
                time.sleep(5)
                
                exportar_excel(session, carpeta_destino, nombre_lote)
                esperar_archivo_escrito(ruta_lote)
                
                archivos_temporales.append(ruta_lote)
                
            except Exception as e:
                print(f"   ❌ Error en bloque {num_bloque} de IP24: {e}")

    # === UNIFICACIÓN FINAL ===
    print("\n" + "="*60)
    print(" UNIFICANDO BLOQUES DE IP24 EN UN SOLO ARCHIVO MAESTRO...")
    dataframes = []
    for archivo in archivos_temporales:
        if os.path.exists(archivo) and os.path.getsize(archivo) > 0:
            try: dataframes.append(pd.read_excel(archivo))
            except: pass
            
    if dataframes:
        df_maestro = pd.concat(dataframes, ignore_index=True)
        df_maestro.drop_duplicates(inplace=True) # Borra repetidos si SAP se cruzó
        
        nombre_final = f"IP24_Programacion_{FECHA_HOY}.xlsx"
        ruta_final = os.path.join(carpeta_destino, nombre_final)
        df_maestro.to_excel(ruta_final, index=False)
        print(f"   ✅ Archivo maestro creado exitosamente: {nombre_final}")
        
        # Limpiar la basura temporal
        for archivo in archivos_temporales:
            try: os.remove(archivo)
            except: pass

# ============================================================================
# BUCLE PRINCIPAL
# ============================================================================
def main():
    session = conectar_sap()
    session.findById("wnd[0]").maximize() 
    
    # --- 1. IP18 ---
    ruta_ip18 = obtener_ruta_archivo_creado(CARPETAS["IP18"], "IP18_Posiciones")
    if ruta_ip18:
        print(f"\n   ⏩ El archivo maestro IP18 YA EXISTE. Saltando ejecución de SAP...")
    else:
        ejecutar_ip18(session)
        ruta_ip18 = verificar_archivo_creado(CARPETAS["IP18"], "IP18_Posiciones")

    if ruta_ip18:
        df_filtrado = leer_y_filtrar_ip18(ruta_ip18)
        
        if df_filtrado is not None and not df_filtrado.empty:
            
            # --- 2. IP16 (Usa PLANES) ---
            if archivo_ya_existe(CARPETAS["IP16"], "IP16_Planes"):
                print(f"\n   ⏩ El archivo final IP16_Planes YA EXISTE. Saltando ejecución...")
                exito_ip16 = True
            else:
                planes = extraer_planes_ip16(df_filtrado)
                if planes:
                    lotes_planes = chunk_list(planes, TAMANO_LOTE)
                    ejecutar_ip16_por_lotes(session, lotes_planes)
                    unir_excels_temporales(CARPETAS["IP16"], "IP16_Planes", "IP16_Planes")
                    exito_ip16 = bool(verificar_archivo_creado(CARPETAS["IP16"], "IP16_Planes", timeout=10))
                else:
                    exito_ip16 = False
                    print("\n🛑 DETENIDO: No se extrajeron PLANES.")
                
            # --- 3. IP24 (Usa POSICIONES) ---
            if exito_ip16:
                if archivo_ya_existe(CARPETAS["IP24"], "IP24_Programacion"):
                    print(f"\n   ✅ El archivo final IP24_Programacion YA EXISTE. Ciclo completo.")
                else:
                    posiciones = extraer_posiciones_ip24(df_filtrado)
                    if posiciones:
                        lotes_posiciones = chunk_list(posiciones, TAMANO_LOTE)
                        ejecutar_ip24_por_lotes(session, lotes_posiciones)
                        unir_excels_temporales(CARPETAS["IP24"], "IP24_Programacion", "IP24_Programacion")
                        verificar_archivo_creado(CARPETAS["IP24"], "IP24_Programacion", timeout=10)
                    else:
                        print("\n🛑 DETENIDO: No se extrajeron POSICIONES.")
            else:
                 print("\n🛑 DETENIDO: Se canceló la IP24 porque no se completó la IP16.")
        else:
            print("\n🛑 DETENIDO: No se pudo leer el archivo IP18 o el filtro dejó la tabla vacía.")
    else:
        print("\n🛑 DETENIDO: Se canceló la IP16 e IP24 porque no se detectó el archivo de la IP18.")

    print("\n" + "="*60)
    print("🏁 AUTOMATIZACIÓN FINALIZADA.")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()