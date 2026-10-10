import argparse
import os
import sys
import config
# Asegúrate de que existe src/selenium_runner.py
from src.selenium_runner import run_chatgpt_via_edge

def main():
    # --- 1. CONFIGURACIÓN DE ARGUMENTOS ---
    parser = argparse.ArgumentParser(description="Robot 2: Subida de Transcripción a ChatGPT")
    
    # Argumento opcional. Si no se da, usa el default de config.py
    parser.add_argument(
        "--transcripcion", 
        type=str, 
        default=None, 
        help="Ruta absoluta del archivo .txt a subir"
    )
    
    args = parser.parse_args()

    print("\n" + "="*50)
    print("       INICIANDO PROCESO: ROBOT 2 (LOG DE EJECUCIÓN)")
    print("="*50 + "\n")

    # --- 2. VALIDACIÓN DE LA RUTA DEL ARCHIVO (WHISPER) ---
    ruta_whisper = args.transcripcion
    
    # A. Si no hay argumento, usamos la ruta Hardcoded
    if not ruta_whisper:
        print("[LOG] No se proporcionó argumento --transcripcion.")
        print(f"[LOG] Usando ruta por defecto definida en Config: \n      -> {config.PATH_WHISPER_DEFAULT}")
        ruta_whisper = config.PATH_WHISPER_DEFAULT
    else:
        print(f"[LOG] Argumento recibido: {ruta_whisper}")

    # B. Limpieza de comillas (común en PowerShell "Copiar como ruta de acceso")
    ruta_whisper = ruta_whisper.strip('"').strip("'")
    
    # C. Verificación de existencia
    if not os.path.exists(ruta_whisper):
        print(f"\n[ERROR FATAL] El archivo no existe en la ruta:\n      {ruta_whisper}")
        print("Revisa que la ruta sea correcta y que el archivo .txt esté ahí.")
        sys.exit(1)
    
    print("[LOG] Archivo de transcripción validado: OK")

    # --- 3. CARGA DEL PROMPT MAESTRO (INSTRUCCIONES) ---
    ruta_instrucciones = config.PATH_INSTRUCCIONES
    print(f"[LOG] Buscando archivo de instrucciones en: \n      -> {ruta_instrucciones}")

    if not os.path.exists(ruta_instrucciones):
        print(f"\n[ERROR FATAL] No se encuentra el archivo de Reglas/Prompt.")
        sys.exit(1)
    
    try:
        with open(ruta_instrucciones, "r", encoding="utf-8") as f:
            prompt_texto = f.read()
        print(f"[LOG] Instrucciones cargadas ({len(prompt_texto)} caracteres). OK")
    except Exception as e:
        print(f"[ERROR] Falló la lectura del archivo de instrucciones: {e}")
        sys.exit(1)

    # --- 4. EJECUCIÓN DE SELENIUM ---
    print("\n" + "-"*30)
    print(" INICIANDO NAVEGADOR (EDGE)")
    print("-"*30)
    
    # Llamamos al módulo que abre el navegador
    resultado = run_chatgpt_via_edge(
        prompt_text=prompt_texto,       # El texto de las reglas
        file_paths=[ruta_whisper],      # La lista con tu archivo Whisper
        wait_seconds=60                 # Tiempo para que GPT procese
    )

    # --- 5. REPORTE FINAL ---
    print("\n" + "="*50)
    if resultado.ok:
        print(" [ÉXITO] El proceso finalizó correctamente.")
        print("         1. Archivo subido.")
        print("         2. Prompt pegado y enviado.")
        print("         3. Revisa la ventana del navegador.")
    else:
        print(f" [FALLO] Ocurrió un error: {resultado.message}")
    print("="*50 + "\n")

if __name__ == "__main__":
    main()