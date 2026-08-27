
from pathlib import Path
import ast
import re

ROOT = Path(__file__).resolve().parents[1]


def _route_endpoints():
    result = set()
    for path in (ROOT / "routes").glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        bp_names = {}
        for node in tree.body:
            if isinstance(node, ast.Assign) and isinstance(node.value, ast.Call):
                if getattr(node.value.func, "id", None) == "Blueprint":
                    if node.targets and isinstance(node.targets[0], ast.Name):
                        name = node.targets[0].id
                        if node.value.args and isinstance(node.value.args[0], ast.Constant):
                            bp_names[name] = node.value.args[0].value
        for node in tree.body:
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            for deco in node.decorator_list:
                if not (isinstance(deco, ast.Call) and isinstance(deco.func, ast.Attribute)):
                    continue
                if deco.func.attr not in {"route", "get", "post", "put", "delete", "patch"}:
                    continue
                base = deco.func.value
                if isinstance(base, ast.Name) and base.id in bp_names:
                    result.add(f"{bp_names[base.id]}.{node.name}")
    return result


def test_active_migration_graph_has_one_root_and_one_head():
    nodes = {}
    for path in (ROOT / "migrations" / "versions").glob("*.py"):
        text = path.read_text(encoding="utf-8")
        revision = re.search(r'^revision\s*=\s*["\']([^"\']+)', text, re.M)
        down = re.search(r'^down_revision\s*=\s*["\']([^"\']+)', text, re.M)
        assert revision, f"Missing revision in {path}"
        nodes[revision.group(1)] = down.group(1) if down else None
    roots = [rev for rev, down in nodes.items() if down is None]
    children = {down for down in nodes.values() if down}
    heads = [rev for rev in nodes if rev not in children]
    assert roots == ["b7e2c1a4f901"]
    assert heads == ["e5f60718293a"]


def test_template_url_for_endpoints_exist():
    endpoints = _route_endpoints()
    missing = []
    for path in (ROOT / "templates").rglob("*.html"):
        text = path.read_text(encoding="utf-8")
        for endpoint in re.findall(r"url_for\(\s*['\"]([^'\"]+)", text):
            if endpoint == "static":
                continue
            if endpoint not in endpoints:
                missing.append((str(path.relative_to(ROOT)), endpoint))
    assert not missing, f"Undefined template endpoints: {missing}"


def test_import_templates_exist():
    assert (ROOT / "templates/apprentices/import.html").is_file()
    assert (ROOT / "templates/groups/import.html").is_file()


def test_no_sensitive_db_uri_logging():
    text = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'logger.debug("DB URI:' not in text


def test_password_change_route_requires_current_password():
    text = (ROOT / "routes/users.py").read_text(encoding="utf-8")
    assert "check_password(current_password)" in text
    assert "new_password != confirm_password" in text
