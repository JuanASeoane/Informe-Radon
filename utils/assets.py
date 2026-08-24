"""Logotipo institucional por defecto (UPRL / Servizo Galego de Saúde /
Área Sanitaria), usado en la cabecera del informe y en el Anexo II cuando el
usuario no sube su propio logotipo."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

_LOGO_PATH = Path(__file__).resolve().parent.parent / "assets" / "logo_uprl.png"


@lru_cache(maxsize=1)
def logo_por_defecto() -> bytes | None:
    """Devuelve los bytes del logotipo por defecto, o None si el fichero no
    está presente en el repositorio (para que el resto de la app siga
    funcionando igualmente, mostrando el texto de repuesto habitual)."""
    try:
        return _LOGO_PATH.read_bytes()
    except OSError:
        return None
