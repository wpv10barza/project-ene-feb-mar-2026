# SAP DATALOG — Ejecución (modo “para niños”)

## 1) ¿Qué hace esto?
Estos scripts descargan/actualizan archivos Excel desde SAP (por transacción: IW29, IA09, IW39, IW47, IW49N, IP16, IP18, IP24, etc.)
y los guardan en la carpeta de la **semana**.

## 2) ¿Dónde se guardan los archivos?
Por defecto guardan en:

C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - P10. DX MANTO CON DATA SAP\05.00 CIERRE\05.01 Herramienta\AAQ\01. Bronce\2026\<SEMANA>\

Ejemplo semana 09:
...\2026\09\

### Cambiar semana (SIN tocar código)
Usa variable de entorno:
- `SAP_DATALOG_SEMANA=09`

### Cambiar carpeta raíz (si cambiaste de PC)
- `SAP_DATALOG_BASE_DIR=TU_RUTA_HASTA...\2026`

## 3) Nombres con fecha (ddmmaa)
Si un archivo tiene nombre con fecha tipo `_200226`, ahora se reemplaza automáticamente por **la fecha de hoy**:
- Ejemplo: `IW47_notificaciones_200226.xlsx` → `IW47_notificaciones_240226.xlsx` (si hoy es 24/02/26)

## 4) Ejecutar con entorno virtual (venv)

### A) Abrir PowerShell
1. Ve a la carpeta donde están los scripts (recomendado copiarlos a la carpeta semana 09 o a una carpeta `scripts`)
2. Shift + click derecho → “Abrir PowerShell aquí”

### B) Crear venv (solo la primera vez)
```powershell
py -m venv .venv
.\.venv\Scripts\activate
pip install --upgrade pip
pip install pywin32 pandas openpyxl
```

### C) Definir la semana 09 (cada vez que cambies semana)
```powershell
$env:SAP_DATALOG_SEMANA="09"
```

### D) Ejecutar un script (ejemplo)
```powershell
python IW39_CABECERA.py
```

## 5) Ejecutar varios scripts con main.py
`main.py` sirve para correr varios scripts uno tras otro.

### Ejemplo: correr un “proceso” completo
```powershell
$env:SAP_DATALOG_SEMANA="09"
python main.py --run Proceso_IW39_IW49N_IW47_IW13.py
```

### Ejemplo: correr una lista
```powershell
python main.py --run IW29_Avisos.py --run IW39_CABECERA.py --run IW47_IW13.py
```

## 6) ¿Qué cambié (explicación simple)?
- ✅ Agregué una **configuración automática** arriba de los scripts:
  - Semana destino (`SAP_DATALOG_SEMANA`)
  - Semana origen (`SAP_DATALOG_SEMANA_ORIGEN`, por defecto semana-1)
  - Carpeta raíz (`SAP_DATALOG_BASE_DIR`)
  - Fecha de hoy (`FECHA_HOY`)
- ✅ Eliminé dependencias de “semana fija” dentro de las rutas (07/08) y ahora usan la semana por variable.
- ✅ Cuando el nombre del archivo trae `_ddmmaa`, se reemplaza por la fecha de hoy.

Si algo falla, lo primero es:
1) confirmar que SAP GUI está abierto y logueado
2) confirmar que `$env:SAP_DATALOG_SEMANA` está bien
3) confirmar permisos de carpeta (OneDrive/SharePoint sincronizado)