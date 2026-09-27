# P13 Sistematizador (Selenium + SQL) — v29

Este proyecto tiene 3 comandos:

- `preview`: lee **un solo TXT** (prompt) y lo ejecuta por Selenium en **chatgpt.com**.
- `run`: une el TSV generado por `preview` con las preguntas SQL de `VW_PREGUNTA_POR_EVALUACION` por `EvaluacionId`.
- `load`: sube una matriz (Excel/CSV) a `dbo.Respuesta` usando MERGE (3 columnas: SituacionActual/Notas/RespuestaLiteral).

## 1) Preparar venv (PowerShell)

```powershell
cd "C:\ruta\a\P13_sistematizador_selenium_v29"

python -m venv .venv
.\.venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
python -m pip install -r .\requirements.txt
```

## 2) Configurar DB

Crea `db_config.json` en la raíz del repo:

### Ejemplo Azure SQL (usuario + password)

```json
{
  "server": "<SQL_SERVER>",
  "database": "<DATABASE>",
  "username": "<USERNAME>",
  "password": "<PASSWORD>",
  "encrypt": "yes",
  "trust_server_certificate": "no"
}
```

### Ejemplo LocalDB (Windows Auth)

```json
{
  "server": "<LOCAL_SQL_SERVER>",
  "database": "<DATABASE>",
  "encrypt": "no",
  "trust_server_certificate": "yes"
}
```

## 3) preview (Selenium)

Por defecto lee este prompt:

- `G:\Mi unidad\TGA_PROYECTO\Nuevotxt\combined_output.txt`

y lo copia a:

- `01.02 Promp\prompt_trabajo.txt`

Luego ejecuta Selenium y guarda salida en `01.03 Resultado ChatGPT\`:

- `resultado.tsv`
- `resultado.csv`
- `resultado.xlsx`

Comando:

```powershell
python .\main.py preview
```

## 4) run (unir con SQL)

Lee `resultado.tsv` y une con preguntas por EvaluacionId:

```powershell
python .\main.py run --evaluacion-id 17 --db-config ".\db_config.json"
```

Salida:

- `01.03 Resultado ChatGPT\resultado_unido.tsv`
- `01.03 Resultado ChatGPT\resultado_unido.csv`
- `01.03 Resultado ChatGPT\resultado_unido.xlsx`

## 5) load (subir matriz a dbo.Respuesta)

```powershell
python .\main.py load --evaluacion-id 17 --informante-id 108 --db-config ".\db_config.json" --excel "C:\ruta\MATRIZ_sistematización.xlsx"
```

> Importante: la Matriz debe contener `PREGUNTA_CODIGO` y alguna(s) columnas de conversación/análisis.
El loader detecta dinámicamente columnas que contengan:
- `CONVERSACION_ENTREVISTADOR` (se agrupa en NOTAS)
- `CONVERSACION_INFORMANTE` (se agrupa en SITUACION)
- `ANÁLISIS_DE_COBERTURA` / `ANALISIS_DE_COBERTURA` (ANALISIS)


## Seguridad de la migración

Los archivos `db_config.json`, `db_config.json.json` y `db_config.json.txt` del origen no se versionan porque contienen configuración sensible. Use `db_config.example.json` como plantilla y mantenga las credenciales fuera del repositorio.
