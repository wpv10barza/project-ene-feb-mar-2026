from __future__ import annotations
import os
import time
from dataclasses import dataclass
from typing import Optional, List

import pyperclip
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config 

@dataclass
class SeleniumResult:
    ok: bool
    text: str
    message: str

def _get_options(use_profile: bool = True) -> Options:
    options = Options()
    if use_profile:
        user_data = getattr(config, "EDGE_USER_DATA_DIR", None)
        profile = getattr(config, "EDGE_PROFILE_DIR", "Default")
        if user_data:
            options.add_argument(f"user-data-dir={user_data}")
            options.add_argument(f"--profile-directory={profile}")
    
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    return options

def _make_driver() -> webdriver.Edge:
    service = Service()
    try:
        return webdriver.Edge(options=_get_options(True), service=service)
    except Exception:
        print("[AVISO] Perfil de Edge ocupado. Abriendo sesión limpia...")
        return webdriver.Edge(options=_get_options(False), service=service)

def _extract_response_robust(driver: webdriver.Edge) -> str:
    """Intenta extraer el texto de 2 formas: Botón Copiar o Lectura Directa"""
    texto_final = ""
    
    # METODO 1: Botón de Copiar
    try:
        # Busca botones con el testid de ChatGPT o etiquetas aria de copiar
        copiers = driver.find_elements(By.CSS_SELECTOR, "button[data-testid$='copy-button']")
        if not copiers:
            copiers = driver.find_elements(By.XPATH, "//button[contains(@aria-label, 'Copy')]")
        
        if copiers:
            last_btn = copiers[-1]
            driver.execute_script("arguments[0].scrollIntoView(true);", last_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", last_btn)
            time.sleep(1)
            texto_final = (pyperclip.paste() or "").strip()
            if len(texto_final) > 50:
                print("[SELENIUM] Texto extraído exitosamente vía Portapapeles.")
                return texto_final
    except Exception as e:
        print(f"[WARN] Método de copiado fallido: {e}")

    # METODO 2: Lectura Directa del HTML (Markdown)
    try:
        print("[SELENIUM] Intentando lectura directa del HTML (Markdown)...")
        bloques = driver.find_elements(By.CSS_SELECTOR, ".markdown")
        if bloques:
            texto_final = bloques[-1].text
            print(f"[SELENIUM] Texto extraído vía HTML ({len(texto_final)} caracteres).")
            return texto_final
    except Exception as e:
        print(f"[ERROR] Falló la lectura directa: {e}")

    return texto_final

def run_chatgpt_via_edge(
    prompt_text: str, 
    file_paths: Optional[List[str]] = None, 
    wait_seconds: int = 240
) -> SeleniumResult:
    driver = None
    try:
        driver = _make_driver()
        driver.get(config.CHATGPT_URL)
        print("[SELENIUM] Cargando ChatGPT...")
        time.sleep(6)

        # 1. Subir Archivos (si existen)
        if file_paths:
            try:
                # El input file suele estar oculto, lo buscamos por tipo
                file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                for path in file_paths:
                    full_path = os.path.abspath(path)
                    if os.path.exists(full_path):
                        print(f"[SELENIUM] Subiendo: {os.path.basename(full_path)}")
                        file_input.send_keys(full_path)
                        time.sleep(4)
            except Exception as e:
                print(f"[SELENIUM] Error subiendo archivos: {e}")

        # 2. Localizar área de texto y pegar prompt
        textarea_id = getattr(config, "CHATGPT_TEXTAREA_ID", "prompt-textarea")
        try:
            box = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, textarea_id))
            )
        except:
            box = driver.find_element(By.TAG_NAME, "textarea")

        driver.execute_script("arguments[0].focus();", box)
        pyperclip.copy(prompt_text)
        time.sleep(1)
        box.send_keys(Keys.CONTROL, 'v')
        time.sleep(2)
        box.send_keys(Keys.ENTER)
        
        print(f"[SELENIUM] Procesando en la nube ({wait_seconds}s)...")
        time.sleep(wait_seconds)

        # 3. Extraer Respuesta
        texto_resultado = _extract_response_robust(driver)
        
        if not texto_resultado or len(texto_resultado) < 10:
            return SeleniumResult(False, "", "La extracción devolvió un texto vacío o muy corto.")

        return SeleniumResult(True, texto_resultado, "OK")

    except Exception as e:
        return SeleniumResult(False, "", f"Error crítico en Selenium: {e}")
    finally:
        print("[SELENIUM] Proceso finalizado en el navegador.")