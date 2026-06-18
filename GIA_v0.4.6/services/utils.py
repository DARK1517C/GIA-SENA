import re
import unicodedata
from datetime import datetime, date, timedelta
try:
    from flask import request
except Exception:
    request = None


def parse_form(form_or_fields, fields=None):
    """Parse request/form data accepting both old and new call styles."""
    source = request.form if fields is None and request is not None else form_or_fields
    field_defs = form_or_fields if fields is None else fields
    data = {}
    for item in field_defs:
        key = item[0] if isinstance(item, (tuple, list)) else item
        data[key] = (source.get(key, "") or "").strip()
    return data


def html_date_value(value):
    parsed = parse_date_value(value)
    return parsed.strftime("%Y-%m-%d") if parsed else (str(value).strip() if value else "")


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


def add_months(value, months):
    source = parse_date_value(value)
    if not source:
        return None
    month = source.month - 1 + months
    year = source.year + month // 12
    month = month % 12 + 1
    days = [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    return date(year, month, min(source.day, days[month - 1]))


def calculate_followup_ranges(practice_start, practice_end):
    start = parse_date_value(practice_start)
    end = parse_date_value(practice_end)
    ranges = {}
    if start:
        ranges["followup_moment1_start"] = start
        ranges["followup_moment1_end"] = start + timedelta(days=15)
        moment2_start = add_months(start, 3)
        ranges["followup_moment2_start"] = moment2_start
        ranges["followup_moment2_end"] = moment2_start + timedelta(days=15) if moment2_start else None
    if end:
        ranges["followup_moment3_start"] = end - timedelta(days=15)
        ranges["followup_moment3_end"] = end
    ranges.setdefault("followup_moment4_start", None)
    ranges.setdefault("followup_moment4_end", None)
    return {key: (value.strftime("%d/%m/%Y") if value else "") for key, value in ranges.items()}


def followup_range_label(start, end):
    start_text = format_date_value(start)
    end_text = format_date_value(end)
    if start_text and end_text:
        return f"{start_text} al {end_text}"
    return start_text or end_text or ""


def calculate_group_validity(training_end_date):
    end = parse_date_value(training_end_date)
    validity = add_months(end, 6) if end else None
    return validity.strftime("%d/%m/%Y") if validity else ""


def validate_group_validity(ep_start_date, training_end_date, group_validity):
    ep_start = parse_date_value(ep_start_date)
    validity = parse_date_value(group_validity)
    calculated = parse_date_value(calculate_group_validity(training_end_date))
    notes = []
    if calculated and validity and abs((validity - calculated).days) > 7:
        notes.append("La vigencia difiere de la regla de 6 meses despues del fin de formacion.")
    if ep_start and validity and abs((validity - add_months(ep_start, 12)).days) > 45:
        notes.append("La vigencia no coincide aproximadamente con 1 ano despues del inicio de practicas.")
    return notes
