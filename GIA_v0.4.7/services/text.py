# services/text.py
import unicodedata
import re

def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.strip().lower()
    # descomponer y eliminar diacríticos (tildes)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    # reemplazar espacios por guion bajo y quitar caracteres no alfanuméricos
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_]", "", s)
    return s
