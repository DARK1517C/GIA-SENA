# GIA — Mega refactor v0.4.8

## Canonical domain

`EvidenceCategory -> EvidenceTemplate -> EvidenceActivity -> EvidenceSubmission -> EvidenceSubmissionAttempt`

`EvidenceComment` stores review/feedback history as structured records. `observations TEXT` is no longer part of the ORM domain.

## File policy

`allowed_extensions` and `max_file_size_mb` are enforced at upload. `requires_signature` is enforced at approval: signature-required evidence cannot be approved until a signed PDF exists.

## Roles

- Aprendiz: own submissions only; upload/resubmit own evidence.
- Instructor: manage assigned groups/submissions; review/approve within scope; sign.
- Instructor líder: global management/review/approval; sign.
- Administrativo: global read-only data/evidence visibility.
- Certificador: global evidence review/approval; no user administration or PDF signing.
- Soporte: global technical/user/catalog administration; evidence management/approval; signing remains intentionally restricted.

## Migration

Migration `a5d8e7f4c2b1` creates attempt history, backfills current submissions as attempt 1, converts legacy `observations` into structured comments, and removes the text history column.
