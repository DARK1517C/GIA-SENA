# services/utils.py
import unicodedata
import re
from datetime import datetime, date, timedelta
try:
    from flask import request, current_app
except Exception:
    request = None
    current_app = None


# ---------------------------
# Parsing helpers
# ---------------------------
def parse_form(form_or_fields, fields=None):
    """Parse request/form data accepting both old and new call styles."""
    source = request.form if fields is None and request is not None else form_or_fields
    field_defs = form_or_fields if fields is None else fields
    data = {}
    for item in field_defs:
        key = item[0] if isinstance(item, (tuple, list)) else item
        data[key] = (source.get(key, "") or "").strip()
    return data


# ---------------------------
# Date helpers and formatting
# ---------------------------
def parse_date_value(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = clean_cell(value)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def format_date_value(value):
    parsed = parse_date_value(value)
    return parsed.strftime("%d/%m/%Y") if parsed else clean_cell(value)


def html_date_value(value):
    parsed = parse_date_value(value)
    return parsed.strftime("%Y-%m-%d") if parsed else (str(value).strip() if value else "")


def add_months(value, months):
    source = parse_date_value(value)
    if not source:
        return None
    month = source.month - 1 + months
    year = source.year + month // 12
    month = month % 12 + 1
    days = [
        31,
        29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
        31,
        30,
        31,
        30,
        31,
        31,
        30,
        31,
        30,
        31,
    ]
    return date(year, month, min(source.day, days[month - 1]))


def calculate_followup_ranges(practice_start, practice_end):
    from services.date_rules import calculate_followup_ranges_from_ep
    return calculate_followup_ranges_from_ep(practice_start, practice_end)


def followup_range_label(start, end):
    start_text = format_date_value(start)
    end_text = format_date_value(end)
    if start_text and end_text:
        return f"{start_text} al {end_text}"
    return start_text or end_text or ""


def calculate_group_validity(training_end_date):
    from services.date_rules import calculate_group_validity_from_training_end
    return calculate_group_validity_from_training_end(training_end_date)


def validate_group_validity(ep_start_date, training_end_date, group_validity):
    from services.date_rules import audit_date_consistency
    return audit_date_consistency(
        ep_start=ep_start_date,
        training_end=training_end_date,
        group_validity=group_validity,
    )


# ---------------------------
# Cell / header helpers
# ---------------------------
def normalize_header(value):
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip().upper()
    text = text.replace("\u00c2\u00b0", "\u00b0").replace("\u00c2\u00ba", "\u00b0")
    text = text.replace("\u00ba", "\u00b0")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.replace("\u00b0", " ")
    text = re.sub(r"[^A-Z0-9/()&. -]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_cell(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    text = str(value).replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)


def build_alias_lookup(mapping):
    lookup = {}
    for key, aliases in mapping.items():
        for alias in aliases:
            lookup[normalize_header(alias)] = key
    return lookup


# ---------------------------
# Normalización de texto y modalidades EP
# ---------------------------
def normalize_text(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[\s\-_]+", " ", s)
    return s


def canonical_ep_modality(raw):
    """
    Devuelve la etiqueta canónica legible (en español) o None si no se reconoce.
    Etiquetas canónicas:
      - 'Contrato de aprendizaje'
      - 'Contrato de vinculo formativo'
      - 'Prácticas en la economía popular y/o campesina'
      - 'Proyecto productivo'
      - 'Vínculo laboral'
    """
    key = normalize_text(raw)
    if not key:
        return None
    if "contrato" in key and "aprendiz" in key:
        return "Contrato de aprendizaje"
    if ("contrato" in key and ("vincul" in key or "víncul" in key or "vinculo" in key)) or ("vincul" in key and "form" in key):
        return "Contrato de vinculo formativo"
    if "practic" in key or "economia popular" in key or "economía popular" in key:
        return "Prácticas en la economía popular y/o campesina"
    if "proyect" in key:
        return "Proyecto productivo"
    if "vincul" in key and "labor" in key:
        return "Vínculo laboral"
    return None


# ---------------------------
# Backwards-compatible exports (optional)
# ---------------------------
__all__ = [
    "parse_form",
    "parse_date_value",
    "format_date_value",
    "html_date_value",
    "add_months",
    "calculate_followup_ranges",
    "followup_range_label",
    "calculate_group_validity",
    "validate_group_validity",
    "normalize_header",
    "clean_cell",
    "build_alias_lookup",
    "normalize_text",
    "canonical_ep_modality",
]
