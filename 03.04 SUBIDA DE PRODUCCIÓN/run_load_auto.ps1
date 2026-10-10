param(
  [int]$EvaluacionId = 17,
  [int]$InformanteId = 108,
  [int]$UsuarioRegistroId = 1,
  [string]$Excel = "C:\ruta\Matriz de Sistematización.xlsx"
)
.\.venv\Scripts\Activate.ps1
python .\main.py load-auto --evaluacion-id $EvaluacionId --informante-id $InformanteId --usuario-registro-id $UsuarioRegistroId --excel $Excel

