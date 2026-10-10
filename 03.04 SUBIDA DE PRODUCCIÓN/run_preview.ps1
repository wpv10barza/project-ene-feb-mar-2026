param(
  [int]$EvaluacionId = 17,
  [int]$InformanteId = 108,
  [string]$Excel = "C:\ruta\Matriz de Sistematización.xlsx"
)
.\.venv\Scripts\Activate.ps1
python .\main.py preview --evaluacion-id $EvaluacionId --informante-id $InformanteId --excel $Excel

