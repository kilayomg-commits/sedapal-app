"""
config.py — Rutas y constantes globales de la aplicación ONCH-SEDAPAL.
"""
from pathlib import Path

# Directorio raíz de la aplicación
BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

# Asegurar que los directorios existan al importar config
DATA_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ZONAS = ["NORTE", "CENTRO"]
TIPOS = ["CR", "POZO"]

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# ─── Configuración por zona ────────────────────────────────────────────────────
# Define las diferencias estructurales entre los archivos fuente de cada zona.
ZONA_CONFIG = {
    "NORTE": {
        "header_row": 5,   # Fila donde están los encabezados CR-XXX  N°Y
        "tipo_row":   6,   # Fila con etiquetas BOMBA / TABLERO / ACCESORIOS
        "data_start": 8,   # Primera fila con materiales
        "ref_rows": {      # Filas de referencia de estilo en PLANTILLA MODELO
            "E_PM03":  10,
            "ACT":     11,
            "MAT":     12,
            "TA_PM03": 16,
            "AH_PM03": 21,
        },
        "num_cols": 21,    # Columnas de salida en la plantilla
    },
    "CENTRO": {
        "header_row": 4,
        "tipo_row":   5,
        "data_start": 7,
        "ref_rows": {
            "E_PM03":  10,
            "ACT":     11,
            "MAT":     12,
            "TA_PM03": 18,
            "AH_PM03": 24,
        },
        "num_cols": 11,
    },
}

# Fila de inicio de datos en la plantilla de salida
PLANTILLA_START_ROW = 10
