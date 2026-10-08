from typing import Optional, List, Dict, Tuple
"""
utils.py — Funciones compartidas: normalización de texto, catálogo de palabras
clave SAP, redondeo decimal compatible con Excel.
"""
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP


# ─── Redondeo ──────────────────────────────────────────────────────────────────
def excel_round(val) -> Decimal:
    return Decimal(str(val)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# ─── Normalización de texto ────────────────────────────────────────────────────
def normalize(text: str) -> str:
    """
    Convierte texto a minúsculas sin tildes, sin símbolo de grado (°/º),
    sin barras y sin comillas. Colapsa espacios múltiples.
    """
    if not text:
        return ""
    t = str(text).strip()
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = re.sub(r"[°º]", " ", t)          # símbolo grado no eliminado por NFD
    t = t.lower()
    t = re.sub(r"[/\\]", " ", t)
    t = re.sub(r'["“”‘’]', "", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


# ─── Catálogo de palabras clave → código SAP ──────────────────────────────────
# Orden: más específico primero (el primero que coincide gana).
KEYWORDS: list[tuple[str, str]] = [
    # Empaquetaduras
    ("empaquetadura jebe chesterton 3 16",          "900972"),
    ("empaquetadura jebe chesterton 1 16",          "900970"),
    ("empaquetadura jebe chesterton 1 8",           "900971"),
    ("empaquetadura fibra vegetal teflonada 1 2",   "900963"),
    ("empaquetadura fibra vegetal teflonada 3 8",   "900964"),
    ("empaquetadura fibra vegetal teflonad 1 2",    "900963"),
    ("empaquetadura fibra vegetal teflonad 3 8",    "900964"),
    ("empaquetadura grafitada 1 4",                 "900966"),
    ("empaquetadura grafitada 3 8",                 "900967"),
    # Cintas (más específicas primero)
    ("cinta aislante super scotch",                 "900919"),
    ("cinta ais super scotch",                      "900919"),   # abreviatura común
    ("cinta mastic",                                "900920"),
    ("cinta vulcanizante",                          "900922"),
    ("cinta teflon",                                "900921"),
    ("cinta algodon",                               "CINTA_ALGODON"),
    ("cinta algod",                                 "CINTA_ALGODON"),
    ("cinta aislante",                              "900918"),
    # Pinturas
    ("pintura negra anticorros",                    "901018"),
    ("pintura anticorros",                          "901018"),
    ("pintura martillada",                          "901017"),
    ("pintura esmalte",                             "901016"),
    # Silicona / lubricante
    ("spray tarjetas",                              "901068"),
    ("silicona en pasta",                           "901064"),
    ("silicona roj",                                "901065"),
    ("lubricante de silicona",                      "901065"),
    # Otros
    ("esparrago",                                   "900974"),
    ("soldadura punto azul",                        "901066"),
    ("aflojatodo",                                  "900883"),
    ("aceite",                                      "900882"),
    ("escobilla",                                   "900973"),
    ("grasa",                                       "900994"),
    ("lija de fierro n 40",                         "901004"),
    ("lija de fierro n 80",                         "901005"),
    ("lija de fierro n 120",                        "901006"),
    ("solvente",                                    "901067"),
    ("thinner",                                     "901078"),
    ("trapo industrial",                            "901079"),
    ("lubricante multiproposito wd",                "WD40"),
    ("wd 40",                                       "WD40"),
    ("wd40",                                        "WD40"),
    ("brocha",                                      "BROCHA"),
]

PERNO_KEYWORDS: list[tuple[str, str]] = [
    ("3 4", "901015"),
    ("1 2", "901014"),
    ("5 8", "901013"),
]

# Materiales especiales (sin SAP real o con precio fijo)
SPECIAL_ITEMS: Dict[str, dict] = {
    "WD40": {
        "sap":    None,
        "desc":   "Lubricante Multiproposito WD-40",
        "precio": Decimal("7.45"),
    },
    "BROCHA": {
        "sap":    "MATERIAL",
        "desc":   'Brocha de 2"',
        "precio": excel_round(9.00 * 1.242),
    },
    "CINTA_ALGODON": {
        "sap":    "MATERIAL",
        "desc":   'Cinta algodón 3/4"',
        "precio": excel_round(10.00 * 1.242),
    },
}


def match_keyword(desc: str) -> Optional[str]:
    """Retorna el token SAP si encuentra coincidencia exacta de keyword."""
    n = normalize(desc)
    if "perno" in n:
        for size_kw, sap in PERNO_KEYWORDS:
            if size_kw in n:
                return sap
        return None
    for kw, sap in KEYWORDS:
        if kw in n:
            return sap
    return None


def match_keyword_loose(desc: str) -> Optional[str]:
    """Coincidencia sin dígitos como fallback."""
    n2 = re.sub(r"\d", "", normalize(desc)).strip()
    for kw, sap in KEYWORDS:
        kw2 = re.sub(r"\d", "", kw).strip()
        if kw2 and kw2 in n2:
            return sap
    return None
