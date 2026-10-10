import os
import pandas as pd
from pathlib import Path
from flask import Flask, request, jsonify, render_template_string, send_file
import google.generativeai as genai
import marko
from datetime import datetime

# --- CONFIGURACIÓN ---
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
# Busca archivos de clave que empiecen con AIza
API_KEY_FILES = list(BASE_DIR.glob("AIza*.txt"))

# Crear carpetas si no existen
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app = Flask(__name__)

# --- FUNCIÓN INTELIGENTE: CONFIGURAR GEMINI ---
model = None

def configure_gemini():
    global model
    try:
        if not API_KEY_FILES:
            print("[ERROR] No se encontró archivo de API KEY (AIza...) en la carpeta.")
            return False
        
        # Leer la API Key
        api_key = API_KEY_FILES[0].read_text(encoding="utf-8").strip()
        genai.configure(api_key=api_key)
        
        generation_config = {
            "temperature": 0.1,
            "top_p": 0.95,
            "max_output_tokens": 8192,
        }
        
        # LISTA DE MODELOS A PROBAR (Del más nuevo al más antiguo)
        # Si uno falla, prueba el siguiente automáticamente.
        candidates = [
            "gemini-1.5-flash",       # Alias estándar
            "gemini-1.5-flash-001",   # Versión específica v1
            "gemini-1.5-flash-002",   # Versión específica v2
            "gemini-1.5-pro",         # Versión Pro
            "gemini-pro"              # Versión Legacy (si todo lo demás falla)
        ]

        print(">>> Iniciando detección de modelo compatible...")
        
        for model_name in candidates:
            try:
                print(f"   ... Probando modelo: '{model_name}'")
                temp_model = genai.GenerativeModel(model_name=model_name, generation_config=generation_config)
                # Prueba de fuego: Intentar generar un "Hola"
                temp_model.start_chat(history=[]).send_message("Hola")
                
                # Si llegamos aquí, funcionó
                model = temp_model
                print(f">>> ¡ÉXITO! Conectado usando el modelo: {model_name}")
                return True
            except Exception as e:
                # Si falla, imprimimos un aviso pequeño y seguimos al siguiente
                print(f"       [X] {model_name} no disponible. (Error 404/400)")
                continue
        
        print("[ERROR CRÍTICO] Ningún modelo funcionó. Verifica tu API Key o conexión.")
        return False

    except Exception as e:
        print(f"[ERROR GENÉRICO] {e}")
        return False

# --- FUNCIÓN: LEER ARCHIVO (CSV O EXCEL) ---
def read_input_file(path):
    """Detecta si es Excel o CSV y devuelve un DataFrame"""
    ext = path.suffix.lower()
    try:
        if ext in ['.xlsx', '.xls']:
            return pd.read_excel(path)
        else:
            try:
                return pd.read_csv(path, encoding='utf-8')
            except UnicodeDecodeError:
                return pd.read_csv(path, encoding='latin1')
    except Exception as e:
        raise ValueError(f"Error leyendo el archivo: {str(e)}")

# --- RUTAS ---

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def process():
    # Asegurar configuración si no se hizo al inicio
    if not model:
        if not configure_gemini():
            return jsonify({"error": "No se pudo conectar con ningún modelo de Gemini. Revisa la consola del servidor."}), 500

    if 'transcripcion' not in request.files or 'preguntas' not in request.files:
        return jsonify({"error": "Faltan archivos."}), 400
    
    file_trans = request.files['transcripcion']
    file_eval = request.files['preguntas']

    if file_trans.filename == '' or file_eval.filename == '':
        return jsonify({"error": "Nombres de archivo vacíos."}), 400

    path_trans = UPLOAD_FOLDER / file_trans.filename
    path_eval = UPLOAD_FOLDER / file_eval.filename
    file_trans.save(path_trans)
    file_eval.save(path_eval)

    try:
        txt_content = path_trans.read_text(encoding="utf-8", errors="ignore")
        df = read_input_file(path_eval)
        eval_text = df.to_string(index=False)

        prompt = (
            "Actúa como un Auditor Experto en Mantenimiento Minero. \n"
            "Tu tarea es analizar la ENTREVISTA y responder los puntos de la EVALUACIÓN.\n"
            "Cita evidencia textual exacta de la entrevista para justificar cada respuesta.\n"
            "Si la respuesta no está en el texto, indica 'No hay información suficiente'.\n\n"
            f"=== TABLA DE EVALUACIÓN (CRITERIOS) ===\n{eval_text}\n\n"
            f"=== TRANSCRIPCIÓN DE ENTREVISTA ===\n{txt_content}\n"
        )

        chat = model.start_chat(history=[])
        response = chat.send_message(prompt)
        ai_text = response.text
        ai_html = marko.convert(ai_text)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_out = f"Resultado_Auditoria_{timestamp}.xlsx"
        path_out = OUTPUT_FOLDER / filename_out

        df_out = pd.DataFrame({'Analisis_IA': [ai_text]})
        df_out.to_excel(path_out, index=False)

        if path_trans.exists(): os.remove(path_trans)
        if path_eval.exists(): os.remove(path_eval)

        return jsonify({
            "status": "success",
            "html": ai_html,
            "download_url": f"/download/{filename_out}"
        })

    except Exception as e:
        print(f"ERROR PROCESANDO: {e}")
        return jsonify({"error": f"Error interno: {str(e)}"}), 500

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(OUTPUT_FOLDER / filename, as_attachment=True)

# --- INTERFAZ HTML ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistematizador v2.2 (Auto-Fix)</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        body { background-color: #f0f2f5; padding-top: 40px; font-family: 'Segoe UI', sans-serif; }
        .card-custom { border: none; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.08); }
        .header-bg { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border-radius: 12px 12px 0 0; padding: 30px; }
        .step-badge { background: #fff; color: #1e3c72; width: 30px; height: 30px; display: inline-flex; align-items: center; justify-content: center; border-radius: 50%; font-weight: bold; margin-right: 10px; }
    </style>
</head>
<body>
<div class="container">
    <div class="card card-custom mx-auto" style="max-width: 800px;">
        <div class="header-bg text-center">
            <h2>🛠️ Sistematizador Inteligente</h2>
            <p class="opacity-75">Soporte para Excel y Autodetección de Modelo IA</p>
        </div>
        <div class="card-body p-4">
            <form id="main-form">
                <div class="row g-3">
                    <div class="col-md-6">
                        <label class="form-label fw-bold"><span class="step-badge">1</span>Evaluación</label>
                        <input type="file" class="form-control" name="preguntas" accept=".xlsx, .xls, .csv" required>
                        <div class="form-text">Sube tu Excel o CSV con las preguntas.</div>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-bold"><span class="step-badge">2</span>Entrevista</label>
                        <input type="file" class="form-control" name="transcripcion" accept=".txt" required>
                        <div class="form-text">Sube la transcripción (.txt).</div>
                    </div>
                </div>
                <hr class="my-4">
                <button type="submit" class="btn btn-primary w-100 btn-lg fw-bold" id="btn-process">⚡ ANALIZAR ENTREVISTA</button>
            </form>

            <div id="loading" class="text-center mt-4" style="display:none;">
                <div class="spinner-border text-primary" role="status"></div>
                <p class="mt-2 fw-bold text-muted">La IA está leyendo y cruzando datos...</p>
            </div>

            <div id="results-area" class="mt-4" style="display:none;">
                <div class="alert alert-success d-flex justify-content-between align-items-center">
                    <strong>✅ Análisis Completado</strong>
                    <a id="download-btn" href="#" class="btn btn-success btn-sm fw-bold">📥 Descargar Excel</a>
                </div>
                <div class="card bg-light border-0">
                    <div class="card-body p-3" style="max-height: 400px; overflow-y: auto;">
                        <div id="preview-content"></div>
                    </div>
                </div>
            </div>

            <div id="error-area" class="alert alert-danger mt-3" style="display:none;"></div>
        </div>
    </div>
</div>

<script>
$(document).ready(function() {
    $('#main-form').on('submit', function(e) {
        e.preventDefault();
        var formData = new FormData(this);
        
        $('#btn-process').prop('disabled', true).text('Procesando...');
        $('#loading').show();
        $('#results-area').hide();
        $('#error-area').hide();

        $.ajax({
            url: '/process',
            type: 'POST',
            data: formData,
            contentType: false,
            processData: false,
            success: function(response) {
                $('#loading').hide();
                $('#btn-process').prop('disabled', false).text('⚡ ANALIZAR ENTREVISTA');
                $('#preview-content').html(response.html);
                $('#download-btn').attr('href', response.download_url);
                $('#results-area').fadeIn();
            },
            error: function(xhr) {
                $('#loading').hide();
                $('#btn-process').prop('disabled', false).text('⚡ ANALIZAR ENTREVISTA');
                var msg = (xhr.responseJSON && xhr.responseJSON.error) ? xhr.responseJSON.error : 'Error de conexión';
                $('#error-area').text('❌ ' + msg).show();
            }
        });
    });
});
</script>
</body>
</html>
"""

if __name__ == '__main__':
    print("--- INICIANDO SERVIDOR ---")
    configure_gemini()
    app.run(debug=True, host="0.0.0.0", port=5000)