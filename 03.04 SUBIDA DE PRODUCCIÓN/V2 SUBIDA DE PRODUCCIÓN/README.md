# GACSAC - Cargador Evaluación (Excel -> SQL) V3

Cambios V3:
- preview NO se cae si hay PREGUNTA_CODIGO sin mapeo (solo WARN).
- load-* sigue siendo estricto.
- repo.py ya no usa pd.read_sql (elimina UserWarning).

Setup:
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r .\requirements.txt

Ejemplo preview:
python .\main.py preview --evaluacion-id 17 --informante-id 108 --excel "C:\...\MATRIZ_sistematización.xlsx"

