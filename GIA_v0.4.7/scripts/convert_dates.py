#!/usr/bin/env python3
"""
scripts/convert_dates.py

Normalize date-like text columns in SQLite DB to DD-MM-YYYY (or DD-MM-YYYY HH:MM:SS).
Handles single dates and ranges like "01/01/2024 - 31/12/2024".

Usage:
  python scripts/convert_dates.py --db path/to/gia.db --report convert_dates_report.txt --backup
"""

import argparse
import sqlite3
import shutil
import sys
import os
import re
from datetime import datetime

# Try to import dateutil for robust parsing; optional
try:
    from dateutil import parser as dateutil_parser  # type: ignore
    HAVE_DATEUTIL = True
except Exception:
    HAVE_DATEUTIL = False

# Columns to attempt to normalize
DEFAULT_COLUMNS = {
    "apprentice": [
        # exact dates
        "practice_start_date",
        "practice_end_date",
        "created_at",
        "group_validity",
        # ranges (may contain "A - B" or similar)
        "evaluation_date",
        "followup_moment1_start",
        "followup_moment1_end",
        "followup_moment2_start",
        "followup_moment2_end",
        "followup_moment3_start",
        "followup_moment3_end",
        "followup_moment4_start",
        "followup_moment4_end",
    ],
}

# Fallback formats
FALLBACK_FORMATS = [
    "%Y-%m-%d",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S.%f",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%Y/%m/%d",
    "%d %b %Y",
    "%d %B %Y",
    "%m/%d/%Y",
    "%d/%m/%Y %H:%M",
    "%d-%m-%Y %H:%M",
    "%Y/%m/%d %H:%M:%S",
]

# Range separators
RANGE_SEPARATORS = [r"\s*-\s*", r"\s*–\s*", r"\s*—\s*", r"\s*to\s*", r"\s*/\s*"]


def split_range(value: str):
    if value is None:
        return [value]
    s = str(value).strip()
    if s == "":
        return [s]
    for sep in RANGE_SEPARATORS:
        parts = re.split(sep, s, maxsplit=1, flags=re.IGNORECASE)
        if len(parts) == 2:
            return [parts[0].strip(), parts[1].strip()]
    return [s]


def _clean_string(s: str):
    s = s.strip()
    if (s.startswith("'") and s.endswith("'")) or (s.startswith('"') and s.endswith('"')):
        s = s[1:-1].strip()
    s = s.replace("\u00A0", " ").replace("\u200b", "").strip()
    return s


def _try_fromiso(s: str):
    # Normalize microseconds to max 6 digits and remove unexpected chars
    s_clean = re.sub(r"[^\dT:\-\. +]", "", s)
    if "." in s_clean:
        main, frac = s_clean.split(".", 1)
        frac_digits = re.sub(r"\D", "", frac)
        if len(frac_digits) > 6:
            frac_digits = frac_digits[:6]
        s_clean = f"{main}.{frac_digits}"
    try:
        return datetime.fromisoformat(s_clean)
    except Exception:
        return None


def parse_date(value: str):
    if value is None:
        return None
    s = _clean_string(str(value))
    if s == "":
        return None

    # Try ISO-like first (handles YYYY-MM-DD and microseconds)
    dt = _try_fromiso(s)
    if dt:
        return dt

    # If dateutil available, prefer it
    if HAVE_DATEUTIL:
        try:
            return dateutil_parser.parse(s, dayfirst=False, yearfirst=False)
        except Exception:
            try:
                return dateutil_parser.parse(s, dayfirst=True)
            except Exception:
                pass

    # Fallback formats
    for fmt in FALLBACK_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue

    # Last resort: try to extract digits groups like DD MM YYYY
    m = re.search(r"(\d{1,2})\D+(\d{1,2})\D+(\d{4})", s)
    if m:
        d, mth, y = m.group(1), m.group(2), m.group(3)
        try:
            return datetime(int(y), int(mth), int(d))
        except Exception:
            pass

    return None


def normalize_to_ddmmyyyy(dt: datetime):
    if dt is None:
        return None
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0 and dt.microsecond == 0:
        return dt.strftime("%d-%m-%Y")
    return dt.strftime("%d-%m-%Y %H:%M:%S")


def choose_part_for_column(col_name: str, parts: list):
    if len(parts) == 1:
        return parts[0]
    lower = col_name.lower()
    if lower.endswith("_start"):
        return parts[0]
    if lower.endswith("_end"):
        return parts[1]
    if lower == "evaluation_date":
        return parts[1]
    return parts[0]


def main():
    parser = argparse.ArgumentParser(description="Normalize date-like text columns in SQLite DB to DD-MM-YYYY format.")
    parser.add_argument("--db", default="gia.db", help="Path to SQLite DB file (default: gia.db)")
    parser.add_argument("--report", default="convert_dates_report.txt", help="Report output file")
    parser.add_argument("--backup", action="store_true", help="Create a .bak copy of the DB before modifying")
    parser.add_argument("--tables", nargs="*", help="Optional: table names to process (default: all in DEFAULT_COLUMNS)")
    args = parser.parse_args()

    db_path = args.db
    report_path = args.report

    if not os.path.exists(db_path):
        print(f"ERROR: DB file not found: {db_path}", file=sys.stderr)
        sys.exit(2)

    if args.backup:
        bak_path = db_path + ".bak"
        shutil.copy2(db_path, bak_path)
        print(f"Backup created: {bak_path}")

    tables_to_process = args.tables if args.tables else list(DEFAULT_COLUMNS.keys())

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    report_lines = []
    total_updates = 0
    total_attempts = 0
    total_failures = 0

    for table in tables_to_process:
        cols = DEFAULT_COLUMNS.get(table)
        if not cols:
            report_lines.append(f"Skipping table {table}: no configured columns.")
            continue

        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (table,))
        if not cur.fetchone():
            report_lines.append(f"Table not found: {table} (skipping).")
            continue

        for col in cols:
            cur.execute(f"PRAGMA table_info({table});")
            cols_info = [r["name"] for r in cur.fetchall()]
            if col not in cols_info:
                report_lines.append(f"Column {col} not in table {table} (skipping).")
                continue

            cur.execute(f"SELECT rowid AS _rowid, {col} FROM {table} WHERE {col} IS NOT NULL AND TRIM({col}) != ''")
            rows = cur.fetchall()
            if not rows:
                report_lines.append(f"No values to process for {table}.{col}")
                continue

            report_lines.append(f"Processing {len(rows)} rows for {table}.{col}")
            updates = 0
            failures = 0
            attempts = 0

            for r in rows:
                attempts += 1
                try:
                    rowid = r[0]
                    raw = r[1]
                except Exception:
                    try:
                        rowid = r["_rowid"]
                    except Exception:
                        failures += 1
                        report_lines.append(f"  [ERROR] cannot determine rowid for row: {r}")
                        continue
                    try:
                        raw = r[col]
                    except Exception:
                        failures += 1
                        report_lines.append(f"  [ERROR] cannot determine value for {table}.{col} rowid={rowid}")
                        continue

                parts = split_range(raw)
                chosen = choose_part_for_column(col, parts)
                dt = parse_date(chosen)
                if dt is None:
                    # fallback: try other part(s)
                    for p in parts:
                        if p == chosen:
                            continue
                        dt = parse_date(p)
                        if dt:
                            break
                if dt is None:
                    failures += 1
                    report_lines.append(f"  [FAIL] {table}.{col} rowid={rowid} value={raw!r}")
                    continue
                newval = normalize_to_ddmmyyyy(dt)
                if str(raw).strip() != newval:
                    try:
                        cur.execute(f"UPDATE {table} SET {col} = ? WHERE rowid = ?", (newval, rowid))
                        updates += 1
                    except Exception as e:
                        failures += 1
                        report_lines.append(f"  [ERROR] updating rowid={rowid} col={col} error={e}")
            conn.commit()
            report_lines.append(f"  {table}.{col}: attempts={attempts} updated={updates} failures={failures}")
            total_updates += updates
            total_attempts += attempts
            total_failures += failures

    conn.close()

    summary = [
        f"DB: {db_path}",
        f"Processed tables: {', '.join(tables_to_process)}",
        f"Total attempts: {total_attempts}",
        f"Total updates: {total_updates}",
        f"Total failures: {total_failures}",
        f"dateutil available: {HAVE_DATEUTIL}",
        "Output date format: DD-MM-YYYY or DD-MM-YYYY HH:MM:SS when time present",
        "Range handling: *_start -> first date; *_end -> second date; evaluation_date -> second date"
    ]
    report = "\n".join(["=== convert_dates report ==="] + summary + [""] + report_lines)
    with open(report_path, "w", encoding="utf8") as f:
        f.write(report)
    print(report)
    print(f"\nReport written to: {report_path}")


if __name__ == "__main__":
    main()
