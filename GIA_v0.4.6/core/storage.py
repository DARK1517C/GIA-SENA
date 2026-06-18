import os
import uuid
import shutil
from typing import Tuple, Optional
from werkzeug.utils import secure_filename
from flask import current_app

# Extensiones permitidas (ajusta según necesidades)
ALLOWED_EXTENSIONS = {"pdf", "docx", "xlsx", "jpg", "png"}

def allowed_file(filename: str) -> bool:
    """Return True if filename has an allowed extension."""
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXTENSIONS

def make_upload_dir(subdir: str) -> str:
    """Ensure upload subdirectory exists and return its absolute path."""
    base = current_app.config.get("UPLOAD_DIR", os.path.join(os.getcwd(), "uploads"))
    upload_dir = os.path.join(base, subdir) if subdir else base
    os.makedirs(upload_dir, exist_ok=True)
    return upload_dir

def unique_filename(original: str) -> str:
    """Generate a filesystem-safe unique filename preserving original name."""
    safe = secure_filename(original)
    uid = uuid.uuid4().hex
    return f"{uid}_{safe}"

def save_file(file, subdir: str = "") -> Tuple[str, str]:
    """
    Save an uploaded file to UPLOAD_DIR/subdir.
    Returns (absolute_path, stored_filename).
    Raises ValueError on invalid file.
    """
    if file is None or not getattr(file, "filename", None):
        raise ValueError("No se proporcionó archivo.")
    filename = file.filename
    if not allowed_file(filename):
        raise ValueError("Tipo de archivo no permitido.")
    upload_dir = make_upload_dir(subdir)
    stored_name = unique_filename(filename)
    path = os.path.join(upload_dir, stored_name)
    file.save(path)
    return path, stored_name

def remove_file(path: str) -> bool:
    """Remove a file if it exists. Returns True if removed, False if not found."""
    try:
        if os.path.isfile(path):
            os.remove(path)
            return True
        return False
    except Exception:
        return False

def move_file(src_path: str, dest_subdir: str) -> Optional[str]:
    """
    Move an existing file into the configured UPLOAD_DIR/dest_subdir.
    Returns new absolute path or None on failure.
    """
    if not os.path.isfile(src_path):
        return None
    dest_dir = make_upload_dir(dest_subdir)
    filename = os.path.basename(src_path)
    dest_path = os.path.join(dest_dir, filename)
    try:
        shutil.move(src_path, dest_path)
        return dest_path
    except Exception:
        return None
