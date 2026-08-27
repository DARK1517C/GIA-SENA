from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ROUTES = ROOT / "routes" / "evidences.py"


def test_normal_comment_is_explicitly_added_to_session():
    text = ROUTES.read_text(encoding="utf-8")
    assert 'comment = submission.add_observation(' in text
    assert 'db.session.add(comment)' in text


def test_correction_comment_is_explicitly_added_to_session():
    text = ROUTES.read_text(encoding="utf-8")
    blocks = text.split('def add_comment', 1)[1]
    assert 'comment = submission.add_observation(' in blocks
    assert 'db.session.add(comment)' in blocks
