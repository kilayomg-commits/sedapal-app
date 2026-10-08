# ONCH · Automatizador SEDAPAL

Aplicación local para consolidar, generar y conciliar las plantillas de
mantenimiento preventivo (CR) para SEDAPAL.

---

## Requisitos previos

- Python 3.10 o superior
- pip

---

## Instalación (una sola vez)

```bash
# 1. Descomprime el ZIP y entra al directorio
cd sedapal_app

# 2. (Opcional pero recomendado) Crea un entorno virtual
python -m venv venv

# En Windows:
venv\Scripts\activate

# En Mac/Linux:
source venv/bin/activate

# 3. Instala las dependencias
pip install -r requirements.txt
```

---

## Ejecución

```bash
streamlit run app.py
```

Se abrirá automáticamente en tu navegador en `http://localhost:8501`.

---

## Flujo de uso

### 1. Gestión de Archivos
- Selecciona **Zona** (NORTE/CENTRO) y **Tipo** (CR/POZO) en la barra lateral.
- Sube los archivos `.xlsx` fuente del período.
- Sube la **plantilla vacía** y el **archivo de estilos** (`.xlsm`).
- Los archivos se guardan en `data/` y persisten entre sesiones.

### 2. Generación
- Verifica que los 3 requisitos estén marcados en verde.
- Pulsa **▶ Generar** y espera el log en tiempo real.
- Descarga el archivo `.xlsm` generado con el botón de descarga.

### 3. Conciliación
- Una vez generado el archivo, ve a la pestaña **Conciliación**.
- Ingresa el monto esperado por CR directamente en la tabla editable,
  **o** importa un CSV/Excel con columnas `CR` y `Monto Esperado`.
- Pulsa **📊 Calcular Conciliación** para ver la tabla con diferencias.
- Exporta el resultado en CSV con un clic.

---

## Estructura de carpetas

```
sedapal_app/
├── app.py                    ← Interfaz Streamlit
├── config.py                 ← Rutas y configuración de zonas
├── requirements.txt
├── modules/
│   ├── utils.py              ← Normalización, keywords SAP
│   ├── ingesta.py            ← Gestión de archivos
│   ├── generador.py          ← Motor de generación XLSM
│   └── conciliacion.py       ← Módulo de conciliación financiera
├── data/                     ← Archivos fuente y plantillas (persistentes)
│   ├── norte/cr/fuentes/
│   ├── norte/cr/plantillas/
│   ├── centro/cr/fuentes/
│   └── centro/cr/plantillas/
└── output/                   ← Archivos XLSM generados
```

---

## Agregar nuevas palabras clave SAP

Edita `modules/utils.py` → lista `KEYWORDS`.
Cada entrada es una tupla `("texto_a_buscar_normalizado", "codigo_sap")`.

---

## Notas técnicas

- La aplicación lee y escribe archivos `.xlsm` con macros VBA preservadas
  usando `openpyxl` con `keep_vba=True`.
- Los archivos en `data/` NO se borran al cerrar la aplicación.
  Elimínalos desde la pestaña "Gestión de Archivos" o manualmente.
- El puerto por defecto es **8501**. Para usar otro:
  `streamlit run app.py --server.port 8502`
