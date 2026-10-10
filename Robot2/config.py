import os
from pathlib import Path

# --- RUTAS DE TRABAJO ---
BASE_DIR = Path(__file__).resolve().parent
DRIVERS_DIR = BASE_DIR / "drivers"
DATA_DIR = BASE_DIR / "data"

for d in [DRIVERS_DIR, DATA_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# --- CONFIGURACIÓN CHATGPT ---
CHATGPT_URL = "https://chatgpt.com/"
CHATGPT_TEXTAREA_ID = "prompt-textarea"
CHATGPT_SEND_BUTTON_SELECTOR = "[data-testid='send-button']"

# --- CONFIGURACIÓN EDGE ---
_local_app_data = os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local")
EDGE_USER_DATA_DIR = os.path.join(_local_app_data, "Microsoft", "Edge", "User Data")
EDGE_PROFILE_DIR = "Default"
DEFAULT_DRIVER_PATH = DRIVERS_DIR / "msedgedriver.exe"

# ==============================================================================
# RUTAS HARDCODED (Reglas y Archivos)
# ==============================================================================
PATH_INSTRUCCIONES = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\05.00 CIERRE\05.01 PEÑ 2026\02. Archivos\C. Promp para gpt\Prompt transcripción.txt"

PATH_WHISPER_DEFAULT = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\05.00 CIERRE\05.01 PEÑ 2026\02. Archivos\A. Audio-transcripción\Texto de Whisper.txt"
