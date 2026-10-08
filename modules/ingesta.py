"""
ingesta.py — Gestión de archivos: guardar, listar y eliminar archivos fuente
y plantillas organizados por zona, tipo y categoría.
"""
import re
from pathlib import Path
from typing import Optional

from config import DATA_DIR


# ─── Rutas ────────────────────────────────────────────────────────────────────
def get_dir(zona: str, tipo: str, categoria: str) -> Path:
    """
    Devuelve (y crea si no existe) el directorio:
        data/<zona>/<tipo>/<categoria>/
    Ejemplo: data/norte/cr/fuentes/
    """
    d = DATA_DIR / zona.lower() / tipo.lower() / categoria
    d.mkdir(parents=True, exist_ok=True)
    return d


# ─── Guardar ──────────────────────────────────────────────────────────────────
def guardar_archivo(
    file_obj,
    zona: str,
    tipo: str,
    categoria: str,
    nombre_fijo: Optional[str] = None,
) -> Path:
    """
    Guarda file_obj (objeto con .name y .read()) en el directorio correspondiente.
    Si se pasa nombre_fijo, se usa ese nombre; si no, el nombre original del archivo.
    """
    dest_dir = get_dir(zona, tipo, categoria)
    fname = nombre_fijo or file_obj.name
    dest = dest_dir / fname
    content = file_obj.read()
    with open(dest, "wb") as f:
        f.write(content)
    return dest


# ─── Detectar CR desde nombre de archivo ──────────────────────────────────────
def detectar_cr(filepath: Path) -> str:
    m = re.search(r"CR-(\d+[A-Za-z]?)", filepath.name, re.IGNORECASE)
    return f"CR-{m.group(1)}" if m else "—"


def detectar_fecha(filepath: Path) -> str:
    m = re.search(r"(\d{2})[._](\d{2})[._](\d{4})", filepath.name)
    if m:
        return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    return "—"


# ─── Listar ───────────────────────────────────────────────────────────────────
def listar_archivos(zona: str, tipo: str, categoria: str) -> list[dict]:
    """
    Devuelve lista de dicts con información de cada archivo en el directorio.
    Cada dict: {nombre, tamano, cr, fecha_archivo, path}
    """
    d = get_dir(zona, tipo, categoria)
    archivos = []
    for f in sorted(d.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        size = f.stat().st_size
        if size < 1024 * 1024:
            size_str = f"{size / 1024:.1f} KB"
        else:
            size_str = f"{size / 1024 / 1024:.1f} MB"
        archivos.append({
            "nombre": f.name,
            "tamano": size_str,
            "cr": detectar_cr(f),
            "fecha_archivo": detectar_fecha(f),
            "path": str(f),
        })
    return archivos


# ─── Eliminar ─────────────────────────────────────────────────────────────────
def eliminar_archivo(zona: str, tipo: str, categoria: str, nombre: str) -> bool:
    d = get_dir(zona, tipo, categoria)
    f = d / nombre
    if f.exists():
        f.unlink()
        return True
    return False


# ─── Rutas de plantillas ──────────────────────────────────────────────────────
def get_plantilla_path(zona: str, tipo: str) -> Path:
    return get_dir(zona, tipo, "plantillas") / f"plantilla_{zona.lower()}_vacia.xlsm"


def get_estilos_path(zona: str, tipo: str) -> Path:
    return get_dir(zona, tipo, "plantillas") / f"estilos_{zona.lower()}_referencia.xlsm"
