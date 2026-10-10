import os
import subprocess
from fastapi import FastAPI, BackgroundTasks

app = FastAPI(title="API RPA - DX SAP GAMSAC")

RUTA_SCRIPTS = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026\11\RPA"

def ejecutar_bot(nombre_script: str, ruta_destino: str):
    script_path = os.path.join(RUTA_SCRIPTS, nombre_script)
    comando = f'python "{script_path}" "{ruta_destino}"'
    print(f"🤖 API ejecutando: {comando}")
    subprocess.run(comando, shell=True)

@app.post("/api/ejecutar-ip16-ip24")
def lanzar_ip16_ip24(ruta_destino: str, background_tasks: BackgroundTasks):
    background_tasks.add_task(ejecutar_bot, "sap_ip18.py", ruta_destino)
    return {"status": "success", "mensaje": "✅ Bot de IP16/IP24 iniciado."}

@app.post("/api/ejecutar-iw39-iw49")
def lanzar_iw39(ruta_destino: str, background_tasks: BackgroundTasks):
    background_tasks.add_task(ejecutar_bot, "Proceso_IW39_IW49N_IW47_IW13.py", ruta_destino)
    return {"status": "success", "mensaje": "✅ Bot de Órdenes iniciado."}

@app.post("/api/ejecutar-ia9-ia10")
def lanzar_ia9(ruta_destino: str, background_tasks: BackgroundTasks):
    background_tasks.add_task(ejecutar_bot, "ia9ia10.py", ruta_destino)
    return {"status": "success", "mensaje": "✅ Bot de Hojas de Ruta iniciado."}
