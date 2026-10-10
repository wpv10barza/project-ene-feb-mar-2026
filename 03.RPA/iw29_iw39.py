import win32com.client
import time
import sys
import os
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================
PUESTO_TRABAJO = "QI3PP"

# Layouts (según tu hallazgo: sin "/")
LAYOUT_IW29 = "MCARRANZA02"
LAYOUT_IW39 = "MCARRANZA03"

FECHA_INI_IW29 = "01.01.2025"
FECHA_FIN_IW29 = "31.12.2999"

# Debug (solo para diagnóstico; no pausas largas)
DEBUG = False
SHORT_SLEEP = 0.15   # base para que SAP alcance a dibujar
LONG_WAIT = 120      # timeout de espera por reportes

def dprint(msg: str):
    if DEBUG:
        print(msg)

def snooze(mult: float = 1.0):
    time.sleep(SHORT_SLEEP * mult)

# ============================================================
# RUTA + CARPETAS + NOMBRES (ddmmyy)
# ============================================================
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

if len(sys.argv) < 2:
    print("❌ ERROR: No ingresaste ninguna ruta en la terminal.")
    print('Uso correcto: python Proceso_IW29_IW39_AUTO.py "C:\\Tu\\Ruta\\08"')
    sys.exit(1)

RUTA_BASE = sys.argv[1].strip()
if RUTA_BASE.startswith("--"):
    RUTA_BASE = RUTA_BASE[2:]

CARPETA_AVISOS = os.path.join(RUTA_BASE, "Avisos")
CARPETA_ORDENES = os.path.join(RUTA_BASE, "Ordenes")
os.makedirs(CARPETA_AVISOS, exist_ok=True)
os.makedirs(CARPETA_ORDENES, exist_ok=True)

FECHA_HOY = datetime.now().strftime("%d%m%y")  # ddmmyy (ej: 200226)
NOMBRE_IW29 = f"IW29_Avisos_{FECHA_HOY}"
NOMBRE_IW39 = f"IW39_CabeceraOrden_{FECHA_HOY}"

print("=" * 80)
print(f"📁 RUTA BASE: {RUTA_BASE}")
print(f"📂 Avisos  -> {os.path.join(CARPETA_AVISOS, NOMBRE_IW29 + '.xlsx')}")
print(f"📂 Ordenes -> {os.path.join(CARPETA_ORDENES, NOMBRE_IW39 + '.xlsx')}")
print("=" * 80)

# ============================================================
# UTILS SAP
# ============================================================
def conectar_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        return app.Children(0).Children(0)
    except Exception as e:
        print(f"❌ Error conectando a SAP: {e}")
        sys.exit(1)

def exists(session, id_str: str) -> bool:
    try:
        obj = session.findById(id_str, False)
        return obj is not None
    except Exception:
        return False

def get(session, id_str: str):
    return session.findById(id_str)

def wait_ready(session, timeout=60) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if not session.Busy:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False

def sbar_text(session) -> str:
    try:
        return session.findById("wnd[0]/sbar").Text
    except Exception:
        return ""

def sbar_type(session) -> str:
    try:
        return session.findById("wnd[0]/sbar").MessageType  # S/W/E/A
    except Exception:
        return ""

def cerrar_popups(session, max_cierres=8):
    # cierra/acepta popups típicos
    for _ in range(max_cierres):
        snooze(1.5)
        try:
            if exists(session, "wnd[1]"):
                # 1) intentar OK/Continuar
                if exists(session, "wnd[1]/tbar[0]/btn[0]"):
                    try:
                        get(session, "wnd[1]/tbar[0]/btn[0]").press()
                        continue
                    except Exception:
                        pass
                # 2) cerrar
                try:
                    get(session, "wnd[1]").close()
                    continue
                except Exception:
                    pass
        except Exception:
            pass
        break

def confirmar_sobrescritura(session):
    # popup de sobrescritura típicamente: SPOP-OPTION1 = Sí
    for i in range(1, 4):
        try:
            if exists(session, f"wnd[{i}]/usr/btnSPOP-OPTION1"):
                get(session, f"wnd[{i}]/usr/btnSPOP-OPTION1").press()
                return True
        except Exception:
            pass
    return False

def safe_set_text(session, id_str: str, value: str) -> bool:
    try:
        if exists(session, id_str):
            get(session, id_str).text = value
            return True
    except Exception:
        pass
    return False

def set_layout_safe(session, field_id: str, layout: str):
    """
    Evita el error de 'no se ha podido encontrar ninguna disposición':
    prueba layout sin / y con /, valida con Enter.
    si falla, limpia para no bloquear.
    """
    candidates = [layout.lstrip("/"), "/" + layout.lstrip("/")]
    seen = set()
    candidates = [c for c in candidates if not (c in seen or seen.add(c))]

    for v in candidates:
        try:
            get(session, field_id).text = v
            get(session, "wnd[0]").sendVKey(0)  # Enter valida el layout
            if sbar_type(session) not in ("E", "A"):
                return v
        except Exception:
            pass

    # si falla: limpiar para no bloquear ejecución
    try:
        get(session, field_id).text = ""
        get(session, "wnd[0]").sendVKey(0)
    except Exception:
        pass

    print(f"   ⚠️ Layout no válido: {layout}. Se ejecutará sin disposición.")
    if sbar_text(session):
        print(f"      SAP: {sbar_text(session)}")
    return None

def archivo_existe(carpeta: str, nombre_base: str):
    for ext in (".xlsx", ".xls", ".XLSX", ".XLS"):
        p = os.path.join(carpeta, nombre_base + ext)
        if os.path.exists(p):
            return p
    return None

# ============================================================
# EXPORTACIÓN (AUTO) - soporta 2 variantes
#   A) Clásica (tu VBS): wnd[1]/usr/ssub/1/txt + btn[20] + ctxt[0] + btn[0]
#   B) ALV "Exportar como": DY_FILENAME + luego Save dialog DY_PATH/DY_FILENAME
# ============================================================
def press_export(session) -> bool:
    # según tu VBS: btn[16]. Dejamos otros por compatibilidad.
    for bid in ("wnd[0]/tbar[1]/btn[16]", "wnd[0]/tbar[1]/btn[45]", "wnd[0]/tbar[1]/btn[43]"):
        try:
            if exists(session, bid):
                get(session, bid).press()
                return True
        except Exception:
            pass
    return False

def wait_wnd(session, idx: int, timeout=10) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout:
        if exists(session, f"wnd[{idx}]"):
            return True
        time.sleep(0.2)
    return False

def exportar_excel(session, carpeta_destino: str, nombre_base: str, reintentos=2) -> bool:
    ya = archivo_existe(carpeta_destino, nombre_base)
    if ya:
        print(f"   ⏩ Ya existe: {ya}. Saltando exportación.")
        return True

    for intento in range(1, reintentos + 1):
        dprint(f"[EXPORT] intento {intento}/{reintentos}")
        wait_ready(session, 60)
        snooze(2)

        if not press_export(session):
            print("   ❌ No encontré botón Exportar en la barra.")
            return False

        # Espera a que aparezca alguna ventana (Exportar como / Guardar)
        if not wait_wnd(session, 1, 10):
            cerrar_popups(session)
            continue

        snooze(2)

        # ---------- VARIANTE A: CLÁSICA (la de tu VBS) ----------
        # wnd[1]/usr/ssub/1/txt  (nombre)
        # wnd[1]/tbar[0]/btn[20] (aceptar nombre / pasar a ruta)
        # wnd[1]/usr/ctxt[0]     (ruta)
        # wnd[1]/tbar[0]/btn[0]  (guardar)
        try:
            if exists(session, "wnd[1]/usr/ssub/1/txt"):
                get(session, "wnd[1]/usr/ssub/1/txt").text = nombre_base
                snooze(1.5)
                if exists(session, "wnd[1]/tbar[0]/btn[20]"):
                    get(session, "wnd[1]/tbar[0]/btn[20]").press()
                snooze(2)

                # set ruta
                if exists(session, "wnd[1]/usr/ctxt[0]"):
                    get(session, "wnd[1]/usr/ctxt[0]").text = carpeta_destino
                snooze(1.5)

                # guardar
                if exists(session, "wnd[1]/tbar[0]/btn[0]"):
                    get(session, "wnd[1]/tbar[0]/btn[0]").press()

                snooze(2)
                confirmar_sobrescritura(session)
                cerrar_popups(session)

                # verificar archivo
                t0 = time.time()
                while time.time() - t0 < 25:
                    ya = archivo_existe(carpeta_destino, nombre_base)
                    if ya:
                        print(f"   ✅ Exportado OK: {ya}")
                        return True
                    time.sleep(1)

                dprint("[EXPORT] Variante A no verificó archivo.")
                continue
        except Exception:
            pass

        # ---------- VARIANTE B: "Exportar como" (ALV moderno) ----------
        # 1) wnd[1] tiene DY_FILENAME (nombre)
        # 2) confirmar exportar (btn[0])
        # 3) aparece Save dialog con DY_PATH y DY_FILENAME (wnd[1] o wnd[2])
        try:
            if exists(session, "wnd[1]/usr/txtDY_FILENAME") or exists(session, "wnd[1]/usr/ctxtDY_FILENAME"):
                if exists(session, "wnd[1]/usr/txtDY_FILENAME"):
                    get(session, "wnd[1]/usr/txtDY_FILENAME").text = nombre_base
                else:
                    get(session, "wnd[1]/usr/ctxtDY_FILENAME").text = nombre_base

                snooze(1.5)

                # confirmar "Exportar a..."
                if exists(session, "wnd[1]/tbar[0]/btn[0]"):
                    get(session, "wnd[1]/tbar[0]/btn[0]").press()
                else:
                    # fallback: intentar Enter
                    try:
                        get(session, "wnd[1]").sendVKey(0)
                    except Exception:
                        pass

                snooze(2)

                # buscar ventana de guardado (a veces queda en wnd[1], a veces pasa a wnd[2])
                save_wnd = None
                for idx in (1, 2, 3):
                    if exists(session, f"wnd[{idx}]/usr/ctxtDY_PATH") or exists(session, f"wnd[{idx}]/usr/ctxt[0]"):
                        save_wnd = idx
                        break

                if save_wnd is None:
                    cerrar_popups(session)
                    continue

                # set path + filename
                if exists(session, f"wnd[{save_wnd}]/usr/ctxtDY_PATH"):
                    get(session, f"wnd[{save_wnd}]/usr/ctxtDY_PATH").text = carpeta_destino
                elif exists(session, f"wnd[{save_wnd}]/usr/ctxt[0]"):
                    get(session, f"wnd[{save_wnd}]/usr/ctxt[0]").text = carpeta_destino

                # nombre final con extensión (para asegurar .xlsx)
                if exists(session, f"wnd[{save_wnd}]/usr/ctxtDY_FILENAME"):
                    get(session, f"wnd[{save_wnd}]/usr/ctxtDY_FILENAME").text = f"{nombre_base}.xlsx"
                elif exists(session, f"wnd[{save_wnd}]/usr/txtDY_FILENAME"):
                    get(session, f"wnd[{save_wnd}]/usr/txtDY_FILENAME").text = f"{nombre_base}.xlsx"
                elif exists(session, f"wnd[{save_wnd}]/usr/ctxt[1]"):
                    get(session, f"wnd[{save_wnd}]/usr/ctxt[1]").text = f"{nombre_base}.xlsx"

                snooze(1.5)

                # guardar (btn[11] suele ser guardar en file dialog; fallback btn[0])
                if exists(session, f"wnd[{save_wnd}]/tbar[0]/btn[11]"):
                    get(session, f"wnd[{save_wnd}]/tbar[0]/btn[11]").press()
                elif exists(session, f"wnd[{save_wnd}]/tbar[0]/btn[0]"):
                    get(session, f"wnd[{save_wnd}]/tbar[0]/btn[0]").press()

                snooze(2)
                confirmar_sobrescritura(session)
                cerrar_popups(session)

                # verificar archivo
                t0 = time.time()
                while time.time() - t0 < 25:
                    ya = archivo_existe(carpeta_destino, nombre_base)
                    if ya:
                        print(f"   ✅ Exportado OK: {ya}")
                        return True
                    time.sleep(1)

                dprint("[EXPORT] Variante B no verificó archivo.")
                continue
        except Exception:
            pass

        # si ninguna variante calzó, intenta limpiar y reintentar
        cerrar_popups(session)

    print("   ❌ No se pudo exportar/guardar automáticamente (revisar IDs del popup de exportación).")
    return False

# ============================================================
# IW29 (según VBS, sin repeticiones)
# ============================================================
def ejecutar_iw29(session):
    ya = archivo_existe(CARPETA_AVISOS, NOMBRE_IW29)
    if ya:
        print(f"⏩ IW29 ya existe: {ya}")
        return True

    print("\n🚀 IW29 (Avisos)")
    get(session, "wnd[0]").maximize()

    # ir a transacción
    get(session, "wnd[0]/tbar[0]/okcd").text = "/nIW29"
    get(session, "wnd[0]").sendVKey(0)
    wait_ready(session, 60)

    # filtros (según VBS)
    safe_set_text(session, "wnd[0]/usr/ctxt[15]", FECHA_INI_IW29)
    safe_set_text(session, "wnd[0]/usr/ctxt[16]", FECHA_FIN_IW29)
    safe_set_text(session, "wnd[0]/usr/ctxt[42]", PUESTO_TRABAJO)
    set_layout_safe(session, "wnd[0]/usr/ctxt[124]", LAYOUT_IW29)

    # checks (según VBS)
    try:
        get(session, "wnd[0]/usr/chk[3]").selected = True
        get(session, "wnd[0]/usr/chk[1]").selected = True
    except Exception:
        pass

    # ejecutar (una vez)
    get(session, "wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session, LONG_WAIT)
    cerrar_popups(session)

    # exportar + guardar automático
    ok = exportar_excel(session, CARPETA_AVISOS, NOMBRE_IW29)
    return ok

# ============================================================
# IW39 (según VBS, sin repeticiones)
# ============================================================
def ejecutar_iw39(session):
    ya = archivo_existe(CARPETA_ORDENES, NOMBRE_IW39)
    if ya:
        print(f"⏩ IW39 ya existe: {ya}")
        return True

    print("\n🚀 IW39 (Cabecera Orden)")

    # ir a transacción
    get(session, "wnd[0]/tbar[0]/okcd").text = "/nIW39"
    get(session, "wnd[0]").sendVKey(0)
    wait_ready(session, 60)

    # filtros (según VBS)
    safe_set_text(session, "wnd[0]/usr/ctxt[13]", PUESTO_TRABAJO)
    safe_set_text(session, "wnd[0]/usr/ctxt[15]", "")
    safe_set_text(session, "wnd[0]/usr/ctxt[57]", "")
    set_layout_safe(session, "wnd[0]/usr/ctxt[131]", LAYOUT_IW39)

    # ejecutar
    get(session, "wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session, LONG_WAIT)
    cerrar_popups(session)

    # exportar + guardar automático
    ok = exportar_excel(session, CARPETA_ORDENES, NOMBRE_IW39)
    return ok

# ============================================================
# MAIN
# ============================================================
def main():
    print("\n" + "=" * 80)
    print("🤖 INICIANDO AUTOMATIZACIÓN SAP (IW29 -> IW39) AUTOMÁTICA")
    print("=" * 80)

    session = conectar_sap()

    ok1 = ejecutar_iw29(session)
    ok2 = ejecutar_iw39(session)

    print("\n" + "=" * 80)
    if ok1 and ok2:
        print("🏁 LISTO: IW29 e IW39 exportados y guardados correctamente.")
    else:
        print("⚠️ Terminado con advertencias. Revisa mensajes arriba (export/layout).")
    print("=" * 80 + "\n")

if __name__ == "__main__":
    main()