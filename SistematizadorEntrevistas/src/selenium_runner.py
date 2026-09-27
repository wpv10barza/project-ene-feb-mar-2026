from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Optional, List # <--- Asegúrate que List está importado

import pyperclip
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config 

@dataclass
class SeleniumResult:
    ok: bool
    tsv_text: str
    message: str

def _kill_edge_processes() -> None:
    os.system('taskkill /f /im msedge.exe /t >nul 2>&1')
    os.system('taskkill /f /im msedgedriver.exe /t >nul 2>&1')
    time.sleep(1)

def _make_driver(headless: bool = False, driver_path: Optional[str] = None) -> webdriver.Edge:
    user_data = getattr(config, "EDGE_USER_DATA_DIR", None)
    profile_dir = getattr(config, "EDGE_PROFILE_DIR", "Default")
    
    options = Options()
    if user_data:
        options.add_argument(f"user-data-dir={user_data}")
        options.add_argument(f"--profile-directory={profile_dir}")
        
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if headless:
        options.add_argument("--headless")

    service = None
    if driver_path and os.path.exists(driver_path):
        service = Service(executable_path=driver_path)
    else:
        print("[INFO] No se encontró driver local. Selenium intentará descargarlo automáticamente...")
        service = Service()
    
    try:
        return webdriver.Edge(options=options, service=service)
    except Exception:
        print("[WARN] Error iniciando Edge. Matando procesos antiguos y reintentando...")
        _kill_edge_processes()
        return webdriver.Edge(options=options, service=service)


def _try_click_copy_code(driver: webdriver.Edge) -> bool:
    try:
        copiers = driver.find_elements(By.XPATH, "//button[contains(., 'Copy code') or .//text()='Copy']")
        if not copiers:
            copiers = driver.find_elements(By.CSS_SELECTOR, "button[class*='text-token-text-secondary']")
        
        if copiers:
            driver.execute_script("arguments[0].scrollIntoView(true);", copiers[-1])
            time.sleep(0.5)
            copiers[-1].click()
            time.sleep(0.5)
            return True
    except Exception:
        pass
    return False

def _extract_last_codeblock_text(driver: webdriver.Edge) -> str:
    try:
        blocks = driver.find_elements(By.CSS_SELECTOR, "pre code")
        if blocks:
            return blocks[-1].text
    except:
        pass
    return ""

# --- FUNCIÓN PRINCIPAL CORREGIDA ---
def run_chatgpt_via_edge(
    prompt_text: str,
    wait_seconds: int = 60,
    sheets_url: Optional[str] = None,
    headless: bool = False,
    driver_path: Optional[str] = None,
    file_paths: Optional[List[str]] = None  # <--- ESTE ARGUMENTO FALTABA EN TU VERSIÓN
) -> SeleniumResult:
    
    driver = None
    try:
        path_check = driver_path or str(config.DEFAULT_DRIVER_PATH)
        driver = _make_driver(headless, path_check)
        
        driver.get(config.CHATGPT_URL)
        print(f"[INFO] Cargando ChatGPT (Esperando 5s)...")
        time.sleep(5) 

        # --- LÓGICA DE SUBIDA DE ARCHIVOS ---
        if file_paths:
            print(f"[INFO] Se detectaron {len(file_paths)} archivos para subir.")
            try:
                # Buscamos el input oculto de tipo file
                file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                
                for f_path in file_paths:
                    if os.path.exists(f_path):
                        print(f"    >> Subiendo: {os.path.basename(f_path)}")
                        file_input.send_keys(f_path)
                        time.sleep(3) # Espera técnica para que procese el upload
                    else:
                        print(f"    [WARN] No existe el archivo: {f_path}")
                
                print("[INFO] Archivos cargados. Esperando 3s adicionales...")
                time.sleep(3)
            except Exception as e:
                print(f"[WARN] No se pudieron subir los archivos. Error: {e}")

        # --- PEGAR TEXTO DEL PROMPT ---
        textarea_id = getattr(config, "CHATGPT_TEXTAREA_ID", "prompt-textarea")
        try:
            box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, textarea_id))
            )
        except:
            box = driver.find_element(By.TAG_NAME, "textarea")

        box.click()
        time.sleep(0.5)
        
        pyperclip.copy(prompt_text)
        box.send_keys(Keys.CONTROL, 'v')
        time.sleep(2)
        
        try:
            btn = driver.find_element(By.CSS_SELECTOR, getattr(config, "CHATGPT_SEND_BUTTON_SELECTOR", "[data-testid='send-button']"))
            btn.click()
        except:
            box.send_keys(Keys.ENTER)

        print(f"[INFO] Esperando respuesta ({wait_seconds}s)...")
        time.sleep(wait_seconds)

        # --- OBTENER RESPUESTA ---
        copied = _try_click_copy_code(driver)
        tsv = (pyperclip.paste() or "").strip() if copied else ""
        
        if not tsv:
            tsv = _extract_last_codeblock_text(driver)

        if not tsv:
            return SeleniumResult(False, "", "No se obtuvo TSV")

        if sheets_url:
            try:
                driver.get(sheets_url)
                time.sleep(5)
                ActionChains(driver).key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
                time.sleep(2)
            except: pass

        return SeleniumResult(True, tsv, "OK")

    except Exception as e:
        return SeleniumResult(False, "", f"Error Selenium: {e}")
    finally:
        # driver.quit() # Opcional: comentar si quieres ver qué pasó
        pass