# diag_import.py
import importlib, traceback, sys

print("Python:", sys.version)
print("Intentando importar routes.evidences ...")
try:
    m = importlib.import_module("routes.evidences")
    print("OK: routes.evidences importado correctamente")
    print("Tiene evidences_bp?:", hasattr(m, "evidences_bp"))
except Exception:
    print("FALLO al importar routes.evidences:")
    traceback.print_exc()
