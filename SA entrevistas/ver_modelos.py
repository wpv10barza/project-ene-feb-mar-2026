from google import genai
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
API_KEY_FILES = list(BASE_DIR.glob("AIza*.txt"))

if API_KEY_FILES:
    key = API_KEY_FILES[0].read_text(encoding="utf-8").strip()
    client = genai.Client(api_key=key)
    
    print("🔍 Modelos disponibles para tu llave:")
    try:
        # Paginamos para ver todos
        for m in client.models.list():
            # Filtramos solo los Flash y Pro
            if "flash" in m.name or "pro" in m.name:
                print(f"   ✅ {m.name}")
    except Exception as e:
        print(f"❌ Error: {e}")
else:
    print("Falta la API Key")