"""
asistente.py — Asistente sin IA: interpreta el log de la última generación y
responde preguntas sobre los datos del archivo generado.

Todo es local (sin servicios externos ni costo): las respuestas se arman con
reglas sobre el log y con pandas sobre la hoja 'PLANTILLA MODELO'.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
from openpyxl import load_workbook

from config import PLANTILLA_START_ROW
from modules.utils import normalize

COLUMNAS = ["Orden", "Operación", "Ubicación", "Línea", "Tipo", "SAP",
            "Descripción", "Cantidad", "Precio", "Subtotal", "Fecha"]
LINEAS = {"E": "Bomba", "TA": "Tablero", "AH": "Accesorios"}


# ─── Hallazgos del log ────────────────────────────────────────────────────────
@dataclass
class Hallazgo:
    nivel: str                     # "error" | "aviso" | "ok"
    titulo: str
    detalle: str
    que_hacer: str
    items: list[str] = field(default_factory=list)


def _explicar_sin_datos(codigo: str, motivos: list[str]) -> tuple[str, str]:
    """Devuelve (detalle, qué hacer) para un CR/Pozo que salió sin datos."""
    m = " ".join(motivos)
    if "no se encontró" in m:
        return (
            f"El archivo de {codigo} no tiene el código '{codigo.replace('-', ' ')}' "
            "en las cabeceras de ninguna hoja.",
            "Revisa que la Zona y el Tipo seleccionados arriba sean los correctos "
            "(un archivo de Pozo no se lee si el Tipo está en CR, y al revés). "
            "Si están bien, abre el archivo y confirma que la cabecera diga, por ejemplo, "
            f"'{codigo.replace('-', ' ')} N° 1 DE 30 HP' en la hoja activa.",
        )
    if "no hay columnas BOMBA" in m:
        return (
            f"Se encontró {codigo}, pero debajo no están las etiquetas BOMBA / TABLERO / ACCESORIOS.",
            "Abre el archivo y revisa la fila siguiente a la cabecera: debe decir BOMBA, "
            "TABLERO y ACCESORIOS justo debajo del código.",
        )
    if "vacías o en 0" in m:
        return (
            f"Se encontró {codigo} con sus columnas, pero todas las cantidades están vacías o en 0.",
            "Es normal si ese mantenimiento no consumió materiales. Si sí los consumió, "
            "revisa que las cantidades estén escritas como número (no como texto) en las "
            "columnas CANT. de ese equipo.",
        )
    return (
        f"No se pudieron leer datos de {codigo}.",
        "Vuelve a generar y revisa el motivo que aparece en el log debajo de 'Sin datos'.",
    )


def analizar_log(logs: list[str]) -> list[Hallazgo]:
    """Convierte las líneas del log en hallazgos con explicación y qué hacer."""
    hallazgos: list[Hallazgo] = []
    if not logs:
        return hallazgos

    ignorados: list[str] = []
    sin_datos: dict[str, list[str]] = {}
    sin_mapear: list[str] = []
    act_faltante: list[str] = []
    ultimo: Optional[str] = None
    en_mapeo = False

    for raw in logs:
        line = raw.rstrip()
        s = line.strip()
        m = re.search(r"IGNORADO \(sin código (\w+)\): (.+)$", s)
        if m:
            ignorados.append(m.group(2))
            continue
        m = re.match(r"⚠ Sin datos para (\S+)", s)
        if m:
            ultimo = m.group(1)
            sin_datos[ultimo] = []
            continue
        if s.startswith("↳") and ultimo in sin_datos:
            sin_datos[ultimo].append(s.lstrip("↳ ").strip())
            continue
        if "MATERIALES SIN MAPEAR" in s:
            en_mapeo = True
            continue
        if en_mapeo and s.startswith("- "):
            item = s[2:]
            if item.startswith("ACT "):
                act_faltante.append(item)
            else:
                sin_mapear.append(item)
            continue
        if s.startswith("✅") or s.startswith("✓"):
            en_mapeo = False

    if ignorados:
        hallazgos.append(Hallazgo(
            "error", f"{len(ignorados)} archivo(s) ignorado(s) por el nombre",
            "El nombre del archivo no tiene el código esperado (CR-XXX o P-XXX), "
            "así que no se procesó.",
            "Renombra el archivo para que incluya el código y la fecha, por ejemplo "
            "'P-642-01.10.2026 AL 01.10.2026.xlsx', y revisa que el Tipo seleccionado "
            "coincida (CR o POZO).",
            ignorados,
        ))

    for codigo, motivos in sin_datos.items():
        detalle, que_hacer = _explicar_sin_datos(codigo, motivos)
        hallazgos.append(Hallazgo("error", f"{codigo} salió sin datos", detalle,
                                  que_hacer, motivos))

    if act_faltante:
        hallazgos.append(Hallazgo(
            "error", "Código ACT sin precio",
            "El código de actividad no está en la hoja ACTIVIDAD del archivo de estilos, "
            "por eso la fila ACT salió con precio 0.",
            "Agrega el código (columna A), su descripción (columna C) y su precio "
            "(columna D) en la hoja ACTIVIDAD del archivo de estilos, súbelo de nuevo "
            "en 'Gestión de Archivos' y vuelve a generar.",
            act_faltante,
        ))

    if sin_mapear:
        hallazgos.append(Hallazgo(
            "aviso", f"{len(sin_mapear)} material(es) sin código SAP",
            "La descripción del material en el archivo fuente no coincide con ninguna "
            "palabra clave conocida, así que salió sin código SAP y con precio 0.",
            "Revisa si la descripción tiene un error de tipeo en el archivo fuente. "
            "Si es un material nuevo, hay que agregarlo al catálogo de palabras clave "
            "(modules/utils.py) y a la hoja MATERIALES del archivo de estilos.",
            sin_mapear,
        ))

    if not hallazgos:
        hallazgos.append(Hallazgo(
            "ok", "La última generación no tuvo problemas",
            "Todos los archivos se leyeron y todos los materiales tienen código SAP.",
            "Puedes descargar el archivo y pasar a Conciliación.",
        ))
    return hallazgos


# ─── Datos del archivo generado ───────────────────────────────────────────────
def cargar_datos(path: str) -> pd.DataFrame:
    """Lee las filas de 'PLANTILLA MODELO' (columnas A–K) a un DataFrame."""
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["PLANTILLA MODELO"]
    filas = []
    for row in ws.iter_rows(min_row=PLANTILLA_START_ROW, max_col=11, values_only=True):
        if not row or not row[2] or not row[4]:
            continue
        filas.append(list(row) + [None] * (11 - len(row)))
    wb.close()
    df = pd.DataFrame(filas, columns=COLUMNAS)
    for c in ("Cantidad", "Precio", "Subtotal"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Ubicación"] = df["Ubicación"].astype(str).str.strip()
    df["Tipo"] = df["Tipo"].astype(str).str.strip()
    df["Línea"] = df["Orden"].astype(str).str.extract(r"^(TA|AH|E)", expand=False)
    df["Equipo"] = pd.to_numeric(
        df["Orden"].astype(str).str.extract(r"(\d+)$", expand=False), errors="coerce"
    )
    return df


def totales_por_ubicacion(df: pd.DataFrame) -> pd.DataFrame:
    """Suma los subtotales PM03 (que ya incluyen ACT y materiales) por CR/Pozo."""
    pm = df[df["Tipo"] == "PM03"]
    t = pm.pivot_table(index="Ubicación", columns="Línea", values="Subtotal",
                       aggfunc="sum", fill_value=0.0)
    for k in ("E", "TA", "AH"):
        if k not in t.columns:
            t[k] = 0.0
    t = t[["E", "TA", "AH"]].rename(columns=LINEAS)
    t["Total"] = t.sum(axis=1)
    equipos = df.groupby("Ubicación")["Equipo"].nunique().rename("Equipos")
    t = t.join(equipos)
    return t.sort_values("Total", ascending=False).round(2)


def resumen(df: pd.DataFrame) -> dict:
    tot = totales_por_ubicacion(df)
    return {
        "ubicaciones": len(tot),
        "equipos": int(tot["Equipos"].sum()) if len(tot) else 0,
        "filas": len(df),
        "total": float(tot["Total"].sum()) if len(tot) else 0.0,
        "bomba": float(tot["Bomba"].sum()) if len(tot) else 0.0,
        "tablero": float(tot["Tablero"].sum()) if len(tot) else 0.0,
        "accesorios": float(tot["Accesorios"].sum()) if len(tot) else 0.0,
    }


def precio_cero(df: pd.DataFrame) -> pd.DataFrame:
    """Filas de material o actividad sin código SAP o con precio 0."""
    x = df[df["Tipo"].isin(["MAT", "ACT"])]
    # Los ítems especiales sin SAP pero con precio (ej. WD-40) son correctos: no se marcan
    x = x[x["Precio"].fillna(0) == 0]
    return x[["Ubicación", "Orden", "Tipo", "SAP", "Descripción", "Cantidad", "Precio"]]


def ranking_materiales(df: pd.DataFrame, por: str = "Subtotal") -> pd.DataFrame:
    mat = df[df["Tipo"] == "MAT"].copy()
    mat["SAP"] = mat["SAP"].fillna("—").astype(str)
    g = mat.groupby(["SAP", "Descripción"], as_index=False).agg(
        Cantidad=("Cantidad", "sum"),
        Subtotal=("Subtotal", "sum"),
        Usos=("Ubicación", "nunique"),
    )
    return g.sort_values(por, ascending=False).round(2).reset_index(drop=True)


def detalle_ubicacion(df: pd.DataFrame, ubicacion: str) -> pd.DataFrame:
    d = df[df["Ubicación"] == ubicacion]
    return d[["Orden", "Tipo", "SAP", "Descripción", "Cantidad", "Precio", "Subtotal", "Fecha"]]


def buscar_material(df: pd.DataFrame, texto: str) -> pd.DataFrame:
    """Dónde se usó un material: busca por descripción o código SAP."""
    q = normalize(texto)
    if not q:
        return df.iloc[0:0]
    mat = df[df["Tipo"] == "MAT"]
    mask = mat["Descripción"].astype(str).map(normalize).str.contains(q, regex=False) | \
        mat["SAP"].astype(str).str.contains(texto.strip(), regex=False)
    return mat[mask][["Ubicación", "Orden", "SAP", "Descripción", "Cantidad", "Precio", "Subtotal"]]


# ─── Preguntas frecuentes ─────────────────────────────────────────────────────
FAQ: list[tuple[str, str]] = [
    ("¿Por qué un CR o Pozo sale 'Sin datos'?",
     "Pasa cuando el programa no encuentra el código (por ejemplo 'P 642') en la "
     "cabecera del archivo, cuando debajo no están BOMBA / TABLERO / ACCESORIOS, o "
     "cuando todas las cantidades están vacías. Debajo de cada 'Sin datos' el log "
     "indica cuál de los tres fue. Lo más común es tener el Tipo (CR/POZO) o la Zona "
     "equivocados en la barra de arriba."),
    ("¿Qué código ACT se usa?",
     "Pozos: siempre 10003485 (MP en Pozo). CR: depende del número de equipos — "
     "1-2 → 10003486, 3-4 → 10003487, 5-6 → 10003488, 7-8 → 10003489, "
     "9 o más → 10003490. El precio sale de la hoja ACTIVIDAD del archivo de estilos."),
    ("¿Por qué la fila ACT sale con precio 0?",
     "El código ACT no está en la hoja ACTIVIDAD del archivo de estilos de esa zona "
     "y tipo. Agrégalo allí (código en A, descripción en C, precio en D), vuelve a "
     "subir el archivo de estilos y genera otra vez."),
    ("¿Por qué un material sale sin código SAP?",
     "La descripción del archivo fuente no coincide con ninguna palabra clave del "
     "catálogo. Suele ser un error de tipeo o un material nuevo. Si es nuevo, hay "
     "que agregarlo al catálogo de palabras clave y a la hoja MATERIALES."),
    ("¿Cómo se calcula el subtotal de la fila PM03?",
     "Bomba (E): suma de sus materiales; en el primer equipo de cada CR/Pozo también "
     "suma el precio del ACT. Tablero (TA) y Accesorios (AH): suma de sus materiales "
     "(sin cantidad ni precio en la fila PM03)."),
    ("¿Cómo deben llamarse los archivos fuente?",
     "Deben incluir el código y la fecha: 'CR-369-27.05.2026 AL 27.05.2026.xlsx' o "
     "'P-642-01.10.2026 AL 01.10.2026.xlsx'. Para CR de 200 HP o 500 HP, el nombre "
     "debe decir '200 HP' o '500 HP'."),
    ("¿Qué archivos necesito para generar?",
     "Los archivos fuente del mes, la plantilla vacía y el archivo de estilos de la "
     "zona y tipo seleccionados. Todo se carga en 'Gestión de Archivos'."),
]


def buscar_faq(texto: str) -> list[tuple[str, str]]:
    """Ordena las preguntas frecuentes por palabras en común con lo que se escribió."""
    palabras = {w for w in normalize(texto).split() if len(w) > 2}
    if not palabras:
        return []
    puntaje = []
    for preg, resp in FAQ:
        base = set(normalize(preg + " " + resp).split())
        p = len(palabras & base)
        if p:
            puntaje.append((p, preg, resp))
    puntaje.sort(key=lambda x: -x[0])
    return [(p, r) for _, p, r in puntaje[:3]]
