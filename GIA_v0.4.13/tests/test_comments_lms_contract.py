from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_comment_model_tracks_attempt():
    text = (ROOT / "models" / "evidence.py").read_text(encoding="utf-8")
    assert 'ForeignKey("evidence_submission_attempt.id")' in text
    assert 'is_correction_request' in text
    assert 'attempt = relationship(' in text


def test_comment_route_restricts_authors_to_apprentice_and_followup_instructors():
    text = (ROOT / "routes" / "evidences.py").read_text(encoding="utf-8")
    assert 'is_instructor = current_user.role in {' in text
    assert '"FOLLOW_UP_INSTRUCTOR"' in text
    assert '"LEAD_FOLLOW_UP_INSTRUCTOR"' in text
    assert 'if not (is_apprentice or is_instructor):' in text


def test_comment_ui_is_conversation_and_preserves_version_reference():
    text = (ROOT / "templates" / "evidences" / "detail.html").read_text(encoding="utf-8")
    assert "Conversación de esta actividad" in text
    assert "Los comentarios permanecen aunque exista una reentrega" in text
    assert "Versión {{ comment.attempt.version_number }}" in text
    assert "Solicitud de corrección" in text


def test_add_observation_attaches_comment_to_submission_relationship():
    text = (ROOT / "models" / "evidence.py").read_text(encoding="utf-8")
    assert "self.comments.append(comment)" in text


def test_detail_route_exposes_explicit_review_ui_flags():
    text = (ROOT / "routes" / "evidences.py").read_text(encoding="utf-8")
    assert "can_request_correction=" in text
    assert "can_approve_evidence=" in text
    assert "can_sign_pdf=" in text


def test_detail_template_uses_review_ui_flags():
    text = (ROOT / "templates" / "evidences" / "detail.html").read_text(encoding="utf-8")
    assert "{% if can_request_correction %}" in text
    assert "{% if can_approve_evidence %}" in text
    assert "{% if can_sign_pdf %}" in text
    assert "request-correction-checkbox" in text
