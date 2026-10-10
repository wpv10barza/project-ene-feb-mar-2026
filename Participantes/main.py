import argparse
import os
import sys
import time
import config
from src.selenium_runner import run_chatgpt_via_edge

def main():
    parser = argparse.ArgumentParser(description="Robot de Clasificación con ChatGPT")
    parser.add_argument("--transcripcion", type=str, default=None, help="Ruta al archivo de Whisper")
    args = parser.parse_args()

    print("\n" + "="*60)
    print("      ROBOT DE CLASIFICACIÓN -> PROCESAMIENTO GPT")
    print("="*60 + "\n")

    # 1. Resolución de Rutas
    ruta_whisper = args.transcripcion if args.transcripcion else config.PATH_WHISPER_DEFAULT
    # Limpiar comillas accidentales
    ruta_whisper = ruta_whisper.strip('"').strip("'")
    
    if not os.path.exists(ruta_whisper):
        print(f"[ERROR] No se encuentra el archivo de transcripción: {ruta_whisper}")
        sys.exit(1)

    if not os.path.exists(config.PATH_INSTRUCCIONES):
        print(f"[ERROR] No se encuentra el archivo de Prompt (Reglas): {config.PATH_INSTRUCCIONES}")
        sys.exit(1)

    # 2. Leer Instrucciones (Prompt Maestro)
    print(f"[INFO] Leyendo reglas desde: {os.path.basename(config.PATH_INSTRUCCIONES)}")
    try:
        with open(config.PATH_INSTRUCCIONES, "r", encoding="utf-8") as f:
            prompt_texto = f.read()
    except UnicodeDecodeError:
        with open(config.PATH_INSTRUCCIONES, "r", encoding="latin-1") as f:
            prompt_texto = f.read()

    # 3. Ejecutar Selenium
    print(f"[INFO] Archivo a procesar: {os.path.basename(ruta_whisper)}")
    
    # Tiempo de espera sugerido para respuestas largas
    TIEMPO_ESPERA_GPT = 220 
    
    resultado = run_chatgpt_via_edge(
        prompt_text=prompt_texto,
        file_paths=[ruta_whisper],
        wait_seconds=TIEMPO_ESPERA_GPT
    )

    # 4. Guardar Resultado Final
    print("\n" + "-"*60)
    if resultado.ok and resultado.text:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        nombre_resultado = f"RESULTADO_GPT_{timestamp}.txt"
        
        # Usar la ruta de salida de config.py
        ruta_final = os.path.join(config.PATH_SALIDA_GPT, nombre_resultado)
        
        try:
            # Asegurar que el directorio de salida existe
            os.makedirs(config.PATH_SALIDA_GPT, exist_ok=True)
            
            with open(ruta_final, "w", encoding="utf-8") as f_out:
                f_out.write(resultado.text)
            
            print(f" [EXITO] Archivo generado exitosamente:")
            print(f" >> {ruta_final}")
            print(f" >> Longitud: {len(resultado.text)} caracteres.")
            
        except Exception as e:
            print(f" [ERROR AL GUARDAR] {e}")
            # Guardado de emergencia
            backup_path = f"EMERGENCIA_{nombre_resultado}"
            with open(backup_path, "w", encoding="utf-8") as f_err:
                f_err.write(resultado.text)
            print(f" >> Se guardó un backup de emergencia en: {os.path.abspath(backup_path)}")
            
    else:
        print(f" [FALLO] No se pudo obtener la respuesta de ChatGPT.")
        print(f" Detalle: {resultado.message}")
            
    print("="*60)

if __name__ == "__main__":
    main()