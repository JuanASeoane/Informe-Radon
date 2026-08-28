"""Logotipo institucional por defecto (UPRL / Servizo Galego de Saúde /
Área Sanitaria), usado en la cabecera del informe y en el Anexo II cuando el
usuario no sube su propio logotipo. También los anexos III (informe de
ensayo del laboratorio acreditado) y IV (certificado ENAC) por defecto, que
apenas cambian de una vez para otra."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

_CARPETA_ASSETS = Path(__file__).resolve().parent.parent / "assets_informe"
_LOGO_PATH = _CARPETA_ASSETS / "logo_uprl.png"
_ANEXO3_PATH = _CARPETA_ASSETS / "anexo3_default.pdf"
_ANEXO4_PATH = _CARPETA_ASSETS / "anexo4_default.pdf"


@lru_cache(maxsize=1)
def logo_por_defecto() -> bytes | None:
    """Devuelve los bytes del logotipo por defecto, o None si el fichero no
    está presente en el repositorio (para que el resto de la app siga
    funcionando igualmente, mostrando el texto de repuesto habitual)."""
    try:
        return _LOGO_PATH.read_bytes()
    except OSError:
        return None


@lru_cache(maxsize=1)
def anexo3_por_defecto() -> tuple[str, bytes] | None:
    """Devuelve (nombre_archivo, bytes) del Anexo III (informe de ensayo del
    laboratorio acreditado) por defecto, o None si no está presente."""
    try:
        return (_ANEXO3_PATH.name, _ANEXO3_PATH.read_bytes())
    except OSError:
        return None


@lru_cache(maxsize=1)
def anexo4_por_defecto() -> tuple[str, bytes] | None:
    """Devuelve (nombre_archivo, bytes) del Anexo IV (certificado ENAC) por
    defecto, o None si no está presente."""
    try:
        return (_ANEXO4_PATH.name, _ANEXO4_PATH.read_bytes())
    except OSError:
        return None

