"""Static audit for the canonical evidence architecture."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
FORBIDDEN = re.compile(r"(?<!audit_evidence_domain\.py)\b(?:EVIDENCE_TYPES|DEFAULT_EVIDENCES)\b")
LEGACY_READS = re.compile(r"(?:activity\.)evidence_type|EvidenceActivity\.evidence_type|\bevidence_type\s*[:=]")

hits = []
for path in ROOT.rglob("*"):
    if not path.is_file() or "__pycache__" in path.parts:
        continue
    if path.name == "audit_evidence_domain.py" or "migrations" in path.parts:
        continue
    if path.suffix not in {".py", ".html", ".js", ".jinja2"}:
        continue
    text = path.read_text(encoding="utf-8", errors="ignore")
    for lineno, line in enumerate(text.splitlines(), 1):
        if FORBIDDEN.search(line):
            hits.append((path.relative_to(ROOT), lineno, line.strip()))
        elif LEGACY_READS.search(line):
            # The migration bridge may write the field; reads are forbidden.
            hits.append((path.relative_to(ROOT), lineno, line.strip()))

if hits:
    print("Legacy evidence architecture references found:")
    for item in hits:
        print(f"{item[0]}:{item[1]}: {item[2]}")
    raise SystemExit(1)

print("OK: canonical evidence domain has no active legacy catalog/type dependency.")
print("Canonical domain: EvidenceCategory -> EvidenceTemplate -> EvidenceActivity -> EvidenceSubmission -> EvidenceSubmissionAttempt")
