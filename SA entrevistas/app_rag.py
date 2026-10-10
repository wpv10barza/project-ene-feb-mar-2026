import streamlit as st
import pandas as pd
import time
import re
from pathlib import Path
from io import StringIO
from google import genai

# --- CONFIGURACIÓN INICIAL ---
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="Auditor IA - Matriz Perfecta", layout="wide")

# --- FUNCIONES DE CONEXIÓN ---
def get_client():
    key_files = list(BASE_DIR.glob("AIza*.txt"))
    if not key_files:
        st.error("❌ No se encontró el archivo de API Key (AIza...)")
        return None
    try:
        key_text = key_files[0].read_text(encoding="utf-8").strip()
        return genai.Client(api_key=key_text)
    except Exception as e:
        st.error(f"❌ Error de conexión: {e}")
        return None

def save_uploaded_file(uploaded_file):
    try:
        file_path = UPLOAD_DIR / uploaded_file.name
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    except Exception as e:
        st.error(f"Error guardando archivo: {e}")
        return None

# --- LÓGICA DE CORRECCIÓN DE FORMATO ---
def clean_and_convert_to_df(text_response):
    """Limpia la respuesta y fuerza la estructura de columnas"""
    clean_text = text_response
    match = re.search(r'```(?:csv|text|markdown)?(.*?)```', text_response, re.DOTALL)
    if match:
        clean_text = match.group(1).strip()
    
    # Columnas esperadas
    expected_cols = [
        "PREGUNTA_CODIGO", "E.1", "I.1", "I.2", "I.3", "I.4", 
        "I.5", "I.6", "I.7", "I.8", "I.9", "I.10", 
        "PREGUNTA_DESCRIPCION", "PREGUNTA_SITUACION_DESEADA", "ANALISIS_DE_COBERTURA"
    ]

    try:
        # Usamos '|' como separador estricto
        df = pd.read_csv(
            StringIO(clean_text), 
            sep="|", 
            names=expected_cols, 
            header=0,            
            skipinitialspace=True,
            on_bad_lines='skip',
            dtype=str # Forzamos todo a texto para no perder ceros o IDs
        )
        
        # Limpieza
        df = df.dropna(how='all')
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
        
        # FILTRO DE SEGURIDAD:
        # Si la columna PREGUNTA_CODIGO está vacía o es un separador, borrar la fila
        df = df[df['PREGUNTA_CODIGO'].str.len() > 1]

        return df
        
    except Exception as e:
        print(f"Error parseando: {e}")
        return None

def process_response_to_excel(text_response):
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    df = clean_and_convert_to_df(text_response)
    
    if df is not None and not df.empty:
        filename = f"Matriz_Final_{timestamp}.xlsx"
        save_path = BASE_DIR / filename
        try:
            df.to_excel(save_path, index=False)
            return str(save_path), df
        except Exception as e:
            st.error(f"Error escribiendo Excel: {e}")
            return None, None
    else:
        filename_txt = f"ERROR_FORMATO_{timestamp}.txt"
        txt_path = BASE_DIR / filename_txt
        txt_path.write_text(text_response, encoding="utf-8")
        return str(txt_path), None

# --- INTERFAZ DE USUARIO ---
st.title("🚀 Auditor IA - Generador de Matrices")
st.info(f"📂 Carpeta de salida: `{BASE_DIR}`")

with st.sidebar:
    st.header("1. Archivos de Entrada")
    file_transcription = st.file_uploader("Subir Transcripción (.txt)", type=["txt"])
    file_questions = st.file_uploader("Subir Preguntas (.csv, .xlsx)", type=["csv", "xlsx"])

st.subheader("2. Configuración del Prompt")
# ==============================================================================
# INSTRUCCIONES: FILTRADO DE FILAS + EXTRACCIÓN DE NOMBRES
# ==============================================================================
instruction_suffix = """
====================================================================
INSTRUCCIONES OBLIGATORIAS DE FORMATO, FILTRADO Y LLENADO
====================================================================
1. FORMATO: Tu respuesta debe ser UNICAMENTE una tabla CSV usando la BARRA VERTICAL '|' como separador.

2. COLUMNAS: Debes entregar EXACTAMENTE estas 15 columnas en orden:
   PREGUNTA_CODIGO|E.1|I.1|I.2|I.3|I.4|I.5|I.6|I.7|I.8|I.9|I.10|PREGUNTA_DESCRIPCION|PREGUNTA_SITUACION_DESEADA|ANALISIS_DE_COBERTURA

3. REGLA DE FILTRADO (CRÍTICO - SOLO HALLAZGOS):
   - ANALIZA CUIDADOSAMENTE: Si la transcripción NO contiene información sobre una pregunta específica, NO GENERES ESA FILA.
   - OMITIR filas que resultarían en "NO CUBIERTA TEXTUALMENTE".
   - Tu tabla final debe contener SOLO las preguntas donde realmente hubo conversación o evidencia en el texto.

4. REGLA DE LLENADO DE METADATA (EXTRAER DEL TEXTO):
   - COLUMNA E.1 (ENTREVISTADOR): Busca en el inicio del texto el nombre de quien hace las preguntas (ej. "Speaker 1" o quien se presenta). Ej: "Marcel Carranza". Pon este nombre.
   - COLUMNA I.1 (INFORMANTE): Busca en el inicio el nombre del entrevistado principal (ej. "Speaker 2"). Ej: "José Arredondo". Pon este nombre.
   - COLUMNAS I.2-I.10: Rellena con '-' si no hay más hablantes.

5. CONTENIDO:
   - No uses saltos de línea (enter) DENTRO de las celdas.
   - En 'ANALISIS_DE_COBERTURA', detalla específicamente (a)_LO_QUE_SI_APARECE y (b)_LO_QUE_NO_APARECE basándote en lo hablado.
"""
prompt_path = BASE_DIR / "REGLA_DE_CARGA_COMPLETA_habilitando.txt"
if prompt_path.exists():
    base_prompt = prompt_path.read_text(encoding="utf-8", errors="ignore")
else:
    base_prompt = "Pega aquí tus instrucciones de auditoría..."

prompt_text = st.text_area("Instrucciones Base:", value=base_prompt, height=200)

if st.button("✨ EJECUTAR ANÁLISIS", type="primary"):
    if not file_transcription or not file_questions:
        st.warning("⚠️ Faltan archivos por subir.")
    else:
        client = get_client()
        if client:
            status = st.status("⚙️ Iniciando proceso...", expanded=True)
            try:
                # 1. Guardar y Leer archivos
                status.write("📥 Leyendo archivos...")
                path_txt = save_uploaded_file(file_transcription)
                path_quest = save_uploaded_file(file_questions)
                
                if path_quest.suffix == '.xlsx':
                    df_q = pd.read_excel(path_quest)
                else:
                    try:
                        df_q = pd.read_csv(path_quest, encoding='utf-8')
                    except:
                        df_q = pd.read_csv(path_quest, encoding='latin1')
                
                # 2. Armar Prompt
                final_prompt = f"{prompt_text}\n\n=== TABLA DE PREGUNTAS ===\n{df_q.to_string(index=False)}\n\n{instruction_suffix}"

                # 3. Subir
                status.write("☁️ Subiendo a Google AI...")
                file_ref = client.files.upload(file=path_txt)
                
                while True:
                    remote_file = client.files.get(name=file_ref.name)
                    if remote_file.state == "ACTIVE": break
                    if remote_file.state == "FAILED": raise Exception("Fallo en Google Cloud.")
                    time.sleep(1)

                # 4. Generar con Gemini 2.5
                status.write("🧠 Analizando con Gemini 2.5 Flash...")
                chat = client.chats.create(model="gemini-2.5-flash")
                
                response = chat.send_message(message=[remote_file, final_prompt])
                
                # 5. Guardar
                status.write("💾 Generando Excel...")
                saved_path, df_result = process_response_to_excel(response.text)
                
                if df_result is not None:
                    status.update(label="✅ Matriz Completa", state="complete", expanded=False)
                    st.success(f"Archivo guardado: {saved_path}")
                    st.dataframe(df_result)
                    with open(saved_path, "rb") as f:
                        st.download_button("⬇️ Descargar Excel", f, file_name=Path(saved_path).name)
                else:
                    status.update(label="⚠️ Revisar Salida", state="error")
                    st.error("Error de formato. Se guardó el texto original.")
                    st.text_area("Debug:", response.text)

            except Exception as e:
                status.update(label="❌ Error", state="error")
                st.error(str(e))