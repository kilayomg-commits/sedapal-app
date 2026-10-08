"""
ingesta.py — Gestión de archivos: guardar, listar y eliminar archivos fuente
y plantillas organizados por zona, tipo y categoría.

Las plantillas (.xlsm) se guardan en Supabase Storage si está configurado,
con fallback al disco local para desarrollo.
"""
import re
from pathlib import Path
from typing import Optional

from config import DATA_DIR


# ─── Supabase (opcional) ──────────────────────────────────────────────────────
_BUCKET  = "plantillas"
_TMP_DIR = Path("/tmp/sedapal_plantillas")


def _supabase_client():
    """Devuelve cliente Supabase o None si no está configurado."""
    try:
        import streamlit as st
        url = st.secrets.get("SUPABASE_URL")
        key = st.secrets.get("SUPABASE_KEY")
        if not url or not key:
            return None
        from supabase import create_client
        return create_client(url, key)
    except Exception:
        return None


def _subir_supabase(nombre: str, data: bytes) -> bool:
    """Sube bytes al bucket de plantillas. Retorna True si OK."""
    sb = _supabase_client()
    if not sb:
        return False
    try:
        sb.storage.from_(_BUCKET).upload(
            nombre,
            data,
            file_options={
                "upsert": "true",
                "content-type": "application/octet-stream",
            },
        )
        return True
    except Exception as e:
        print(f"[Supabase] Error subiendo {nombre}: {e}")
        return False


def _descargar_supabase(nombre: str) -> Optional[bytes]:
    """Descarga un archivo del bucket. Retorna bytes o None."""
    sb = _supabase_client()
    if not sb:
        return None
    try:
        return sb.storage.from_(_BUCKET).download(nombre)
    except Exception:
        return None


def _tmp_path(nombre: str) -> Path:
    """Path en /tmp/ para cachear plantillas dentro de la sesión del servidor."""
    _TMP_DIR.mkdir(exist_ok=True)
    return _TMP_DIR / nombre


# ─── Rutas ────────────────────────────────────────────────────────────────────
def get_dir(zona: str, tipo: str, categoria: str) -> Path:
    """
    Devuelve (y crea si no existe) el directorio:
        data/<zona>/<tipo>/<categoria>/
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
    Guarda file_obj en el directorio correspondiente.
    Para plantillas: intenta subir a Supabase primero; si no hay Supabase,
    guarda en disco local.
    """
    content = file_obj.read() if hasattr(file_obj, "read") else bytes(file_obj)
    fname = nombre_fijo or file_obj.name

    if categoria == "plantillas":
        ok = _subir_supabase(fname, content)
        if ok:
            # Cachear localmente para uso inmediato en esta sesión
            tmp = _tmp_path(fname)
            tmp.write_bytes(content)
            return tmp

    # Fallback (o categoría no-plantilla): guardar en disco local
    dest_dir = get_dir(zona, tipo, categoria)
    dest = dest_dir / fname
    dest.write_bytes(content)
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
def listar_archivos(zona: str, tipo: str, categoria: str) -> list:
    """
    Devuelve lista de dicts con información de cada archivo en el directorio.
    """
    d = get_dir(zona, tipo, categoria)
    archivos = []
    for f in sorted(d.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        size = f.stat().st_size
        size_str = (
            f"{size / 1024:.1f} KB"
            if size < 1024 * 1024
            else f"{size / 1024 / 1024:.1f} MB"
        )
        archivos.append({
            "nombre": f.name,
            "tamano": size_str,
            "cr":     detectar_cr(f),
            "fecha_archivo": detectar_fecha(f),
            "path":   str(f),
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


# ─── Rutas de plantillas (Supabase → /tmp/ → disco local) ────────────────────
def get_plantilla_path(zona: str, tipo: str) -> Path:
    """
    Devuelve el path de la plantilla vacía.
    Orden: caché /tmp/ → Supabase → disco local.
    """
    nombre = f"plantilla_{zona.lower()}_vacia.xlsm"
    tmp = _tmp_path(nombre)
    if not tmp.exists():
        data = _descargar_supabase(nombre)
        if data:
            tmp.write_bytes(data)
            return tmp
        # Fallback: disco local
        return get_dir(zona, tipo, "plantillas") / nombre
    return tmp


def get_estilos_path(zona: str, tipo: str) -> Path:
    """
    Devuelve el path del archivo de estilos/referencia.
    Orden: caché /tmp/ → Supabase → disco local.
    """
    nombre = f"estilos_{zona.lower()}_referencia.xlsm"
    tmp = _tmp_path(nombre)
    if not tmp.exists():
        data = _descargar_supabase(nombre)
        if data:
            tmp.write_bytes(data)
            return tmp
        # Fallback: disco local
        return get_dir(zona, tipo, "plantillas") / nombre
    return tmp
