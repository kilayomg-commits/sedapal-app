"""
almacen.py — Guardado permanente de archivos (fuentes, plantillas, archivos
generados y montos esperados).

Streamlit Cloud borra el disco y la memoria cada vez que la app se reinicia,
se actualiza o se duerme. Por eso todo se guarda en Supabase Storage (bucket
"plantillas") y, si Supabase no está configurado, se cae a disco local, que
sirve para desarrollo pero NO es permanente en Streamlit Cloud.

Claves (rutas dentro del bucket):
    plantilla_<zona>_vacia.xlsm, estilos_<zona>_referencia.xlsm   (raíz, como antes)
    fuentes/<ZONA>_<TIPO>/<nombre en base64>   (los nombres traen espacios y símbolos
                                                que Supabase no acepta en una clave)
    salidas/<archivo generado>.xlsm
    esperados/<ZONA>_<TIPO>_<MES><AÑO>.json
"""
from __future__ import annotations

import base64
import os
import time
from pathlib import Path
from typing import Optional

from config import DATA_DIR

BUCKET = "plantillas"
_LOCAL = DATA_DIR / "_almacen"

_estado_cache: dict = {"t": 0.0, "valor": None}
_ultimo_error: Optional[str] = None


# ─── Conexión ─────────────────────────────────────────────────────────────────
def _credenciales() -> tuple[Optional[str], Optional[str]]:
    url = key = None
    try:
        import streamlit as st
        url = st.secrets.get("SUPABASE_URL")
        key = st.secrets.get("SUPABASE_KEY")
    except Exception:
        pass
    return url or os.environ.get("SUPABASE_URL"), key or os.environ.get("SUPABASE_KEY")


def _cliente():
    """Cliente de Supabase, o None si no está configurado o no se puede crear."""
    global _ultimo_error
    url, key = _credenciales()
    if not url or not key:
        return None
    try:
        from supabase import create_client
        return create_client(url, key)
    except Exception as exc:
        _ultimo_error = f"No se pudo conectar con Supabase: {exc}"
        return None


def estado(forzar: bool = False) -> dict:
    """
    Devuelve {"persistente": bool, "detalle": str}.
    persistente=True solo si Supabase responde y el bucket existe.
    Se cachea 60 s para no consultar en cada recarga.
    """
    ahora = time.time()
    if not forzar and _estado_cache["valor"] and ahora - _estado_cache["t"] < 60:
        return _estado_cache["valor"]

    url, key = _credenciales()
    if not url or not key:
        val = {"persistente": False,
               "detalle": "Faltan SUPABASE_URL y SUPABASE_KEY en los Secrets de Streamlit."}
    else:
        sb = _cliente()
        if sb is None:
            val = {"persistente": False, "detalle": _ultimo_error or "No se pudo conectar."}
        else:
            try:
                sb.storage.from_(BUCKET).list("", {"limit": 1})
                val = {"persistente": True, "detalle": "Supabase conectado."}
            except Exception as exc:
                val = {"persistente": False,
                       "detalle": f"Supabase respondió con error (¿existe el bucket "
                                  f"'{BUCKET}' y la clave es correcta?): {exc}"}
    _estado_cache.update(t=ahora, valor=val)
    return val


def ultimo_error() -> Optional[str]:
    return _ultimo_error


# ─── Operaciones ──────────────────────────────────────────────────────────────
def _local(clave: str) -> Path:
    p = _LOCAL / clave
    return p


def guardar(clave: str, data: bytes) -> bool:
    """Guarda `data`. Devuelve True si quedó en Supabase (permanente)."""
    global _ultimo_error
    sb = _cliente()
    if sb is not None:
        try:
            sb.storage.from_(BUCKET).upload(
                clave, data,
                file_options={"upsert": "true", "content-type": "application/octet-stream"},
            )
            return True
        except Exception as exc:
            _ultimo_error = f"No se pudo guardar '{clave}' en Supabase: {exc}"
    p = _local(clave)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)
    return False


def leer(clave: str) -> Optional[bytes]:
    sb = _cliente()
    if sb is not None:
        try:
            return sb.storage.from_(BUCKET).download(clave)
        except Exception:
            pass
    p = _local(clave)
    return p.read_bytes() if p.is_file() else None


def listar(prefijo: str) -> list[str]:
    """Nombres (último tramo de la clave) de los archivos dentro de `prefijo`."""
    nombres: set[str] = set()
    sb = _cliente()
    if sb is not None:
        try:
            for it in sb.storage.from_(BUCKET).list(prefijo, {"limit": 1000}):
                n = it.get("name")
                if n and not n.startswith("."):
                    nombres.add(n)
        except Exception:
            pass
    d = _local(prefijo)
    if d.is_dir():
        nombres.update(f.name for f in d.iterdir() if f.is_file())
    return sorted(nombres)


def borrar(clave: str) -> None:
    global _ultimo_error
    sb = _cliente()
    if sb is not None:
        try:
            sb.storage.from_(BUCKET).remove([clave])
        except Exception as exc:
            _ultimo_error = f"No se pudo borrar '{clave}' de Supabase: {exc}"
    p = _local(clave)
    if p.is_file():
        p.unlink()


# ─── Archivos fuente ──────────────────────────────────────────────────────────
def _codificar(nombre: str) -> str:
    return base64.urlsafe_b64encode(nombre.encode("utf-8")).decode("ascii").rstrip("=")


def _decodificar(codigo: str) -> Optional[str]:
    try:
        return base64.urlsafe_b64decode(codigo + "=" * (-len(codigo) % 4)).decode("utf-8")
    except Exception:
        return None


def _carpeta_fuentes(zona: str, tipo: str) -> str:
    return f"fuentes/{zona.upper()}_{tipo.upper()}"


def guardar_fuente(zona: str, tipo: str, nombre: str, data: bytes) -> bool:
    return guardar(f"{_carpeta_fuentes(zona, tipo)}/{_codificar(nombre)}", data)


def borrar_fuente(zona: str, tipo: str, nombre: str) -> None:
    borrar(f"{_carpeta_fuentes(zona, tipo)}/{_codificar(nombre)}")


def cargar_fuentes(zona: str, tipo: str) -> dict[str, bytes]:
    """Todos los archivos fuente guardados de una zona y tipo: {nombre: bytes}."""
    carpeta = _carpeta_fuentes(zona, tipo)
    out: dict[str, bytes] = {}
    for codigo in listar(carpeta):
        nombre = _decodificar(codigo)
        if not nombre:
            continue
        data = leer(f"{carpeta}/{codigo}")
        if data:
            out[nombre] = data
    return out


# ─── Archivo generado ─────────────────────────────────────────────────────────
def guardar_salida(nombre: str, data: bytes) -> bool:
    return guardar(f"salidas/{nombre}", data)


def leer_salida(nombre: str) -> Optional[bytes]:
    return leer(f"salidas/{nombre}")


# ─── Montos esperados ─────────────────────────────────────────────────────────
def guardar_esperados(clave: str, json_texto: str) -> bool:
    return guardar(f"esperados/{clave}.json", json_texto.encode("utf-8"))


def leer_esperados(clave: str) -> Optional[str]:
    data = leer(f"esperados/{clave}.json")
    return data.decode("utf-8") if data else None
