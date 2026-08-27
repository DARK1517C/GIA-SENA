"""
SERVICIO DE DOCUMENTOS PDF

Responsabilidades

1. Abrir documentos PDF.
2. Obtener metadatos.
3. Descargar documentos.
4. Firmar documentos PDF.
5. Guardar documentos firmados.

Este servicio NO contiene lógica de negocio relacionada con
evidencias, aprendices o certificaciones.

Toda la lógica de aprobación de evidencias pertenece a
services/evidence_service.py.

Este servicio únicamente manipula documentos PDF.
"""

from __future__ import annotations

from pathlib import Path

from services.permissions import has_permission
from datetime import datetime
from typing import Optional

import pymupdf  # PyMuPDF
from flask import (
    current_app,
    send_file,
)

# =============================================================================
# CONFIGURACIÓN
# =============================================================================

PDF_EXTENSION = ".pdf"

DEFAULT_SIGNATURE_TEXT = "Documento firmado electrónicamente"

DEFAULT_SIGNATURE_FONT = "helv"

DEFAULT_SIGNATURE_FONT_SIZE = 9

DEFAULT_SIGNATURE_COLOR = (0, 0, 0)

DEFAULT_SIGNATURE_MARGIN = 20


# =============================================================================
# UTILIDADES PDF
# =============================================================================

def pdf_exists(file_path: str | Path) -> bool:
    """
    Verifica si un archivo PDF existe.
    """

    return Path(file_path).is_file()


def get_pdf(file_path: str | Path) -> pymupdf.Document:
    """
    Abre un documento PDF.

    Raises
    ------
    FileNotFoundError
        Si el archivo no existe.
    """

    file_path = Path(file_path)

    if not file_path.exists():

        raise FileNotFoundError(
            f"No existe el archivo PDF: {file_path}"
        )

    return pymupdf.open(file_path)


def get_pdf_metadata(file_path: str | Path) -> dict:
    """
    Obtiene información básica del documento PDF.
    """

    pdf = get_pdf(file_path)

    try:

        return {

            "filename": Path(file_path).name,

            "pages": pdf.page_count,

            "size": Path(file_path).stat().st_size,

            "modified_at": datetime.fromtimestamp(
                Path(file_path).stat().st_mtime
            ),

            "is_encrypted": pdf.is_encrypted,

        }

    finally:

        pdf.close()


def get_pdf_page_count(file_path: str | Path) -> int:
    """Devuelve el número de páginas de un PDF sin exponer el documento."""
    pdf = get_pdf(file_path)
    try:
        return pdf.page_count
    finally:
        pdf.close()


def render_pdf_page(
    file_path: str | Path,
    page_number: int,
    *,
    zoom: float = 1.0,
) -> tuple[bytes, int, int]:
    """Renderiza una página PDF a PNG para el visor integrado de GIA.

    page_number es 1-indexado para coincidir con la interfaz.
    Devuelve (png_bytes, page_width, page_height) en píxeles renderizados.
    """
    if page_number < 1:
        raise IndexError("El número de página debe ser mayor o igual a 1.")
    if zoom <= 0:
        raise ValueError("El zoom debe ser positivo.")

    pdf = get_pdf(file_path)
    try:
        index = page_number - 1
        if index >= pdf.page_count:
            raise IndexError("La página solicitada no existe.")

        page = pdf.load_page(index)
        matrix = pymupdf.Matrix(zoom, zoom)
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        return pixmap.tobytes("png"), pixmap.width, pixmap.height
    finally:
        pdf.close()

# =============================================================================
# DESCARGA DE DOCUMENTOS
# =============================================================================

def download_pdf(file_path: str | Path):
    """
    Devuelve un documento PDF para descarga.
    """

    file_path = Path(file_path)

    if not pdf_exists(file_path):

        raise FileNotFoundError(
            f"No existe el archivo PDF: {file_path}"
        )

    return send_file(

        file_path,

        mimetype="application/pdf",

        as_attachment=True,

        download_name=file_path.name,

    )


# =============================================================================
# UTILIDADES DE MANIPULACIÓN
# =============================================================================

def close_pdf(document: pymupdf.Document) -> None:
    """
    Cierra un documento PDF de forma segura.
    """

    if document is not None:

        document.close()


def save_pdf(
    document: pymupdf.Document,
    output_path: str | Path,
) -> Path:
    """
    Guarda un documento PDF.

    Si la carpeta no existe, será creada automáticamente.
    """

    output_path = Path(output_path)

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    document.save(
        output_path,
        garbage=4,
        deflate=True,
    )

    current_app.logger.info(
        "PDF guardado correctamente: %s",
        output_path,
    )

    return output_path


def duplicate_pdf(
    source_path: str | Path,
    destination_path: str | Path,
) -> Path:
    """
    Duplica un documento PDF.

    Esta función es útil cuando se desea firmar una copia del
    documento original sin modificarlo.
    """

    pdf = get_pdf(source_path)

    try:

        return save_pdf(
            pdf,
            destination_path,
        )

    finally:

        close_pdf(pdf)


# =============================================================================
# VALIDACIONES
# =============================================================================

def is_pdf_signed(
    file_path: str | Path,
) -> bool:
    """
    Determina si el PDF ya fue firmado por GIA.

    Actualmente la implementación es básica y siempre devuelve
    False.

    En futuras versiones verificará la existencia del sello de
    firma generado por GIA.
    """

    return False

# =============================================================================
# FIRMA DE DOCUMENTOS
# =============================================================================

def can_sign_pdf(user) -> bool:
    """Determina si el usuario puede firmar documentos PDF."""
    role = getattr(user, "role", None)
    return role in {
        "FOLLOW_UP_INSTRUCTOR",
        "LEAD_FOLLOW_UP_INSTRUCTOR",
    } and has_permission("evidences.sign")


def get_signature_image(user):
    """
    Obtiene la imagen PNG de la firma del usuario.

    Returns
    -------
    pathlib.Path | None

        Devuelve la ruta absoluta de la firma si existe.

        Si el usuario aún no tiene firma registrada,
        devuelve None.
    """

    signature_path = getattr(
        user,
        "signature_file_path",
        None,
    )

    if not signature_path:
        return None

    signature_path = Path(signature_path)

    if not signature_path.exists():
        return None

    return signature_path


def build_signature_metadata(user):
    """
    Construye la información que acompaña la firma.
    """

    now = datetime.now()

    full_name = (
        f"{user.first_names} {user.last_names}"
        if hasattr(user, "first_names")
        else getattr(user, "full_name", "")
    )

    return {
        "full_name": full_name.strip(),
        "role": getattr(user, "role", ""),
        "date": now.strftime("%d/%m/%Y"),
        "time": now.strftime("%H:%M"),
    }


def sign_pdf(
    input_path,
    output_path,
    user,
    page_number,
    x,
    y,
    width,
    height,
    signature_mode="required",
):
    """
    Firma un documento PDF.

    Parameters
    ----------
    signature_mode

        required
            La firma es obligatoria.

        optional
            Si existe firma se inserta.

        none
            No inserta firma.
    """

    if not can_sign_pdf(user):
        raise PermissionError(
            "El usuario no tiene permisos para firmar documentos."
        )

    document = get_pdf(input_path)

    page = document[page_number]

    metadata = build_signature_metadata(user)

    signature = None

    if signature_mode != "none":
        signature = get_signature_image(user)

    if signature_mode == "required" and signature is None:
        raise FileNotFoundError(
            "El usuario no tiene una firma registrada."
        )

    image_rect = pymupdf.Rect(
    x,
    y,
    x + width,
    y + height,
    )

    text_x = x
    text_y = y + height + 10

    if signature is not None:

        page.insert_image(
            image_rect,
            filename=str(signature),
            keep_proportion=True,
            overlay=True,
        )

    page.insert_text(
        pymupdf.Point(text_x, text_y),
        metadata["full_name"],
        fontsize=9,
    )

    page.insert_text(
        pymupdf.Point(text_x, text_y + 12),
        metadata["role"],
        fontsize=8,
    )

    page.insert_text(
        pymupdf.Point(text_x, text_y + 24),
        f"Fecha: {metadata['date']}",
        fontsize=8,
    )

    page.insert_text(
        pymupdf.Point(text_x, text_y + 36),
        f"Hora: {metadata['time']}",
        fontsize=8,
    )

    save_pdf(
        document,
        output_path,
    )

    close_pdf(document)

    return output_path