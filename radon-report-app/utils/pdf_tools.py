"""
Conversión a PDF y fusión de documentos para generar un único PDF completo
con el informe y todos los anexos adjuntos.

Requiere que el binario `soffice` (LibreOffice) esté disponible en el
sistema para convertir .docx/.doc a PDF. En Streamlit Community Cloud esto
se consigue añadiendo `libreoffice` al fichero `packages.txt` del repositorio
(paquete apt). La conversión de imágenes (jpg/png) a PDF se hace con
Pillow, sin dependencias externas.
"""

from __future__ import annotations

import io
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter


class PdfToolsError(RuntimeError):
    """Error al convertir o fusionar documentos a PDF."""


def libreoffice_disponible() -> bool:
    """True si el binario `soffice`/`libreoffice` está instalado."""
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def _soffice_bin() -> str:
    return shutil.which("soffice") or shutil.which("libreoffice") or "soffice"


def convertir_office_a_pdf(contenido: bytes, extension: str) -> bytes:
    """Convierte un .docx/.doc (u otro formato soportado por LibreOffice) a
    PDF usando `soffice --headless --convert-to pdf`."""
    if not libreoffice_disponible():
        raise PdfToolsError(
            "LibreOffice no está instalado en este entorno, así que no se puede "
            "convertir el documento a PDF. Añade 'libreoffice' al fichero "
            "packages.txt del repositorio para habilitar esta función en Streamlit "
            "Community Cloud."
        )
    extension = extension.lstrip(".").lower() or "docx"
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        origen = tmp_path / f"documento.{extension}"
        origen.write_bytes(contenido)
        try:
            subprocess.run(
                [
                    _soffice_bin(),
                    "--headless",
                    "--norestore",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(tmp_path),
                    str(origen),
                ],
                check=True,
                capture_output=True,
                timeout=120,
            )
        except subprocess.CalledProcessError as e:
            raise PdfToolsError(
                f"Fallo al convertir el documento a PDF con LibreOffice: "
                f"{e.stderr.decode(errors='ignore')[:500]}"
            ) from e
        except subprocess.TimeoutExpired as e:
            raise PdfToolsError("La conversión a PDF con LibreOffice tardó demasiado.") from e

        pdf_path = tmp_path / "documento.pdf"
        if not pdf_path.exists():
            raise PdfToolsError("LibreOffice no generó el PDF esperado.")
        return pdf_path.read_bytes()


def imagen_a_pdf(contenido: bytes) -> bytes:
    """Convierte una imagen (jpg/png/...) en un PDF de una página."""
    img = Image.open(io.BytesIO(contenido))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    buffer = io.BytesIO()
    img.save(buffer, format="PDF")
    return buffer.getvalue()


def archivo_a_pdf(nombre: str, contenido: bytes) -> bytes:
    """Convierte cualquier anexo soportado (pdf/doc/docx/jpg/png) a PDF.
    Si ya es un PDF, se devuelve tal cual."""
    ext = nombre.rsplit(".", 1)[-1].lower() if "." in nombre else ""
    if ext == "pdf":
        return contenido
    if ext in ("doc", "docx"):
        return convertir_office_a_pdf(contenido, ext)
    if ext in ("jpg", "jpeg", "png"):
        return imagen_a_pdf(contenido)
    raise PdfToolsError(f"Formato de anexo no soportado para PDF: .{ext}")


def fusionar_pdfs(pdfs: list[bytes]) -> bytes:
    """Fusiona una lista de PDFs (bytes) en un único PDF, en el orden dado."""
    writer = PdfWriter()
    for contenido in pdfs:
        reader = PdfReader(io.BytesIO(contenido))
        for page in reader.pages:
            writer.add_page(page)
    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def construir_pdf_completo(
    informe_docx: bytes,
    anexos: list[tuple[str, bytes]],
) -> bytes:
    """Convierte el informe (.docx) a PDF y lo fusiona, en orden, con los
    anexos indicados como lista de (nombre_archivo, contenido).

    Lanza `PdfToolsError` si algún paso de conversión falla (p.ej. si
    LibreOffice no está disponible, o un anexo tiene un formato no
    soportado)."""
    partes: list[bytes] = [convertir_office_a_pdf(informe_docx, "docx")]
    for nombre, contenido in anexos:
        partes.append(archivo_a_pdf(nombre, contenido))
    return fusionar_pdfs(partes)
