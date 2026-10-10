# GACSAC - Cargador Evaluación (Excel -> SQL) con Preview / Insert / Merge

## 1) Requisitos
- Windows + PowerShell
- Python 3.10+ (recomendado)
- ODBC Driver 18 for SQL Server instalado

## 2) Crear entorno virtual (.venv)
```powershell
cd <esta_carpeta>
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r .\requirements.txt
```

## 3) Configurar conexión (db_config.json)
Crea un archivo **db_config.json** en la raíz del proyecto (al lado de main.py):

```json
{
  "server": "TU_SERVIDOR",
  "database": "gaeval_v2",
  "username": "TU_USUARIO",
  "password": "TU_PASSWORD"
}
```

## 4) Ejecutar Preview (no escribe en BD)
Ejemplo (prueba):
```powershell
python .\main.py preview --evaluacion-id 17 --informante-id 108 --excel "C:\ruta\Matriz de Sistematización.xlsx"
```

El preview reporta:
- Total de preguntas en VW_PREGUNTA_POR_EVALUACION (por EVALUACION_ID)
- Preguntas en Excel (por PREGUNTA_CODIGO)
- Match / faltantes / extras
- AplicacionId (EvaluacionId+InformanteId)
- Respuestas existentes en dbo.Respuesta (AplicacionId)
- Preguntas “subidas” vs “no subidas”

## 5) Cargar a BD (auto: si está vacío inserta, si no hace merge)
```powershell
python .\main.py load-auto --evaluacion-id 17 --informante-id 108 --usuario-registro-id 1 --excel "C:\ruta\Matriz de Sistematización.xlsx"
```

- Si dbo.Respuesta para esa AplicacionId está vacía → inserta (solo faltantes)
- Si ya hay respuestas → MERGE con concatenación de “elementos nuevos”

## 6) Modos explícitos
Insert (solo faltantes):
```powershell
python .\main.py load-insert --evaluacion-id 17 --informante-id 108 --usuario-registro-id 1 --excel "C:\ruta\Matriz de Sistematización.xlsx"
```

Merge (concatena):
```powershell
python .\main.py load-merge --evaluacion-id 17 --informante-id 108 --usuario-registro-id 1 --excel "C:\ruta\Matriz de Sistematización.xlsx"
```

## 7) Columnas esperadas en Excel
Sheet: RESULTADO  
Columnas:
- PREGUNTA_CODIGO
- CONVERSACION_ENTREVISTADOR_1
- CONVERSACION_ENTREVISTADOR_2
- CONVERSACION_INFORMANTE_1
- ANÁLISIS_DE_COBERTURA

Puedes editarlo en config.py si cambian.



## Inventario de Google Drive

- Nombre: 03.04 SUBIDA DE PRODUCCIÓN.
- Propósito documentado: GACSAC - Cargador Evaluación (Excel -> SQL) con Preview / Insert / Merge.
- Periodo de archivos: 2025-12-23 a 2026-01-02.
- Tecnologías y formatos: .md, .ps1, .py, .txt.
- Estructura: README.md, V2 SUBIDA DE PRODUCCIÓN, config.py, main.py, requirements.txt, run_load_auto.ps1, run_preview.ps1, setup_venv.ps1, src.
- Módulos principales: V2 SUBIDA DE PRODUCCIÓN/config.py, V2 SUBIDA DE PRODUCCIÓN/main.py, src/db.py, src/excel_reader.py, src/mapping.py, src/loaders.py, src/repo.py, src/sql.py, config.py, main.py.
- Origen: carpeta Drive `103fEMtiJl_GgEXICz9cuNITpiBFFJk7v`.
- Archivos originales seleccionados: 17.
- Observaciones: se conserva el original de Drive cuando existen archivos excluidos o insumos no copiados. Se conserva cada proyecto por separado.
