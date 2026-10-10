import os
from pathlib import Path

# --- RUTAS BASE ---
BASE_DIR = Path(__file__).resolve().parent
DRIVERS_DIR = BASE_DIR / "drivers"
DATA_DIR = BASE_DIR / "data"

# Crear carpetas si no existen
for d in [DRIVERS_DIR, DATA_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# --- CONFIGURACIÓN CHATGPT ---
CHATGPT_URL = "https://chatgpt.com/"
CHATGPT_TEXTAREA_ID = "prompt-textarea"

# --- CONFIGURACIÓN EDGE ---
# Intenta obtener el LocalAppData para cargar el perfil de usuario de Edge
_local_app_data = os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local")
EDGE_USER_DATA_DIR = os.path.join(_local_app_data, "Microsoft", "Edge", "User Data")
EDGE_PROFILE_DIR = "Default"
DEFAULT_DRIVER_PATH = DRIVERS_DIR / "msedgedriver.exe"

# ==============================================================================
# RUTAS DEL PROYECTO (Hardcoded según requerimiento)
# ==============================================================================

# 1. Donde están las INSTRUCCIONES (Prompt)
PATH_INSTRUCCIONES = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 05.01 PEÑ 2026\02. Archivos\C. Promp para gpt\Prompt transcripción.txt"

# 2. Archivo WHISPER por defecto
PATH_WHISPER_DEFAULT = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 05.01 PEÑ 2026\02. Archivos\A. Audio-transcripción\Texto de Whisper.txt"

# 3. CARPETA DE SALIDA (Donde se guardará el resultado procesado)
PATH_SALIDA_GPT = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 05.01 PEÑ 2026\02. Archivos\C. Promp para gpt"