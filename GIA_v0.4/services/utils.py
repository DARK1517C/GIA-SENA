import re
import unicodedata
from datetime import datetime
from flask import request

def parse_form(fields):
    data = {}
    for key, _label in fields:
        data[key] = request.form.get(key, "").strip()
    return data

def html_date_value(value):
    text_value = (value or "").strip()
    if not text_value:
        return ""
    for date_format in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text_value, date_format).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text_value

def normalize_header(value):
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text)
    return text

def clean_cell(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    text = str(value).replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)

def build_alias_lookup(mapping):
    lookup = {}
    for key, aliases in mapping.items():
        for alias in aliases:
            lookup[normalize_header(alias)] = key
    return lookup
