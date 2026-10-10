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

def _get_options(use_profile: bool = True, headless: bool = False) -> Options:
    options = Options()
    
    # Configuración de Perfil (Solo si se pide y existe configuración)
    if use_profile:
        user_data = getattr(config, "EDGE_USER_DATA_DIR", None)
        profile = getattr(config, "EDGE_PROFILE_DIR", "Default")
        if user_data:
            # Importante: Se imprime para depuración
            print(f"[DEBUG] Intentando usar perfil: {profile}")
            options.add_argument(f"user-data-dir={user_data}")
            options.add_argument(f"--profile-directory={profile}")
    
    # Opciones estándar anti-detección y estabilidad
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--remote-allow-origins=*") # Crucial para nuevas versiones

    if headless:
        options.add_argument("--headless")
        
    return options

def _make_driver() -> webdriver.Edge:
    """
    Intenta crear el driver con perfil. Si falla, crea uno limpio.
    """
    service = Service() # Deja que Selenium Manager busque el driver
    
    # INTENTO 1: Con Perfil de Usuario
    try:
        print("[SELENIUM] Intentando abrir Edge con tu perfil actual...")
        return webdriver.Edge(options=_get_options(use_profile=True), service=service)
    except Exception as e:
        print(f"[AVISO] No se pudo abrir con el perfil (¿Quizás tienes Edge abierto?).")
        print(f"        Error técnico: {e}")
        print("[SELENIUM] -> Intentando abrir sesión LIMPIA (Sin perfil)...")
    
    # INTENTO 2: Sesión Limpia (Fallback)
    try:
        return webdriver.Edge(options=_get_options(use_profile=False), service=service)
    except Exception as e:
        raise Exception(f"No se pudo iniciar Edge de ninguna forma. Error: {e}")

def run_chatgpt_via_edge(
    prompt_text: str,
    file_paths: Optional[List[str]] = None,
    wait_seconds: int = 60
) -> SeleniumResult:
    
    driver = None
    try:
        driver = _make_driver()
        
        print("[SELENIUM] Cargando ChatGPT...")
        driver.get(config.CHATGPT_URL)
        
        # Espera inicial generosa
        time.sleep(5)

        # 1. SUBIR ARCHIVOS
        if file_paths:
            print(f"[SELENIUM] Subiendo {len(file_paths)} archivos...")
            try:
                file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                for path in file_paths:
                    if os.path.exists(path):
                        print(f"   -> {os.path.basename(path)}")
                        file_input.send_keys(path)
                        time.sleep(3) 
                    else:
                        print(f"   [ERROR] No existe: {path}")
                time.sleep(2) 
            except Exception as e:
                print(f"[SELENIUM] Error upload (no crítico): {e}")

        # 2. PEGAR PROMPT
        print("[SELENIUM] Pegando instrucciones...")
        try:
            box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, config.CHATGPT_TEXTAREA_ID))
            )
        except:
            # Fallback a buscar por tag si cambia el ID
            box = driver.find_element(By.TAG_NAME, "textarea")

        box.click()
        time.sleep(1)
        
        pyperclip.copy(prompt_text)
        box.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        # 3. ENVIAR
        print("[SELENIUM] Enviando...")
        box.send_keys(Keys.ENTER)
        
        print(f"[SELENIUM] Esperando {wait_seconds}s para respuesta...")
        time.sleep(wait_seconds)

        return SeleniumResult(True, "OK", "OK")

    except Exception as e:
        return SeleniumResult(False, "", f"Error Selenium: {e}")
    finally:
        print("[SELENIUM] Finalizado. (Ventana abierta para revisión)")
        # driver.quit() # Descomentar si quieres que se cierre solo