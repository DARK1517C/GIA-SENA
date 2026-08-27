from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_pdf_viewer_has_render_endpoint_and_helpers():
    routes = (ROOT / "routes" / "evidences.py").read_text(encoding="utf-8")
    service = (ROOT / "services" / "pdf_service.py").read_text(encoding="utf-8")
    assert 'def pdf_page(' in routes
    assert 'render_pdf_page(' in routes
    assert 'def get_pdf_page_count(' in service
    assert 'def render_pdf_page(' in service


def test_detail_uses_integrated_pdf_viewer():
    html = (ROOT / "templates" / "evidences" / "detail.html").read_text(encoding="utf-8")
    assert 'gia-pdf-viewer' in html
    assert 'data-open-pdf-viewer' in html
    assert 'data-pdf-page' in html
    assert 'data-pdf-next' in html
    assert 'data-pdf-zoom-in' in html
    assert 'evidences.pdf_page' in html


def test_pymupdf_is_explicit_dependency():
    req = (ROOT / "requirements.txt").read_text(encoding="utf-8")
    assert 'PyMuPDF' in req or 'pymupdf' in req.lower()
