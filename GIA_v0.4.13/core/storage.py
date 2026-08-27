import os
import uuid
import shutil
from typing import Tuple, Optional
from werkzeug.utils import secure_filename
from flask import current_app
from werkzeug.datastructures import FileStorage

# Extensiones permitidas
ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx",
    "xls",
    "xlsx",
    "jpg",
    "jpeg",
    "png",
    "webp",
}

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



def _validate_file_signature(file: FileStorage, extension: str) -> None:
    """Valida firmas binarias básicas para evitar depender solo de la extensión/MIME del cliente."""
    signatures = {
        "pdf": lambda data: data.startswith(b"%PDF-"),
        "png": lambda data: data.startswith(b"\x89PNG\r\n\x1a\n"),
        "jpg": lambda data: data.startswith(b"\xff\xd8\xff"),
        "jpeg": lambda data: data.startswith(b"\xff\xd8\xff"),
        "webp": lambda data: len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP",
        # OOXML (docx/xlsx) son contenedores ZIP.
        "docx": lambda data: data.startswith(b"PK\x03\x04"),
        "xlsx": lambda data: data.startswith(b"PK\x03\x04"),
        # Office binario antiguo (.doc/.xls) usa OLE Compound File.
        "doc": lambda data: data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"),
        "xls": lambda data: data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"),
    }
    check = signatures.get(extension)
    if check is None:
        return

    position = file.stream.tell()
    try:
        file.stream.seek(0)
        header = file.stream.read(16)
    finally:
        file.stream.seek(position)

    if not check(header):
        raise ValueError(f"El contenido del archivo no coincide con el tipo .{extension}.")

def save_file(
    file: FileStorage,
    subdir: str = "",
    *,
    allowed_extensions=None,
    max_size_mb: int | None = None,
) -> Tuple[str, str]:
    """Guarda un archivo aplicando la política efectiva de la actividad."""
    if file is None or not getattr(file, "filename", None):
        raise ValueError("No se proporcionó archivo.")
    filename = file.filename
    ext = filename.rsplit(".", 1)[1].lower() if "." in filename else ""
    allowed = None
    if allowed_extensions:
        allowed = {str(x).lower().lstrip(".") for x in allowed_extensions}
    elif ALLOWED_EXTENSIONS:
        allowed = ALLOWED_EXTENSIONS
    if not ext or ext not in allowed:
        raise ValueError(f"Extensión no permitida: .{ext or 'sin_extensión'}.")

    _validate_file_signature(file, ext)

    if max_size_mb is not None:
        try:
            max_bytes = int(max_size_mb) * 1024 * 1024
        except (TypeError, ValueError):
            raise ValueError("La política de tamaño máximo no es válida.")
        current = file.stream.tell()
        file.stream.seek(0, os.SEEK_END)
        size = file.stream.tell()
        file.stream.seek(current)
        if size > max_bytes:
            raise ValueError(f"El archivo supera el máximo permitido de {max_size_mb} MB.")
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
