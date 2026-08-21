"""
Lectura y normalización del Excel de detectores de radón.

Columnas esperadas en la hoja "Detectores" (o la primera hoja del libro):

    Centro, Área, ID, Planta, Sala, Código Sala,
    Personas trabajando en la sala, Puestos en la sala, Código,
    Fecha de colocación, Fecha de retirada real,
    Resultado medición Bq/m³, Plano, Foto situación, Foto detector
"""

from __future__ import annotations

import pandas as pd

# Nombre de columna tal y como aparece en el Excel de origen.
# La columna de resultado tiene un salto de línea real dentro del propio
# encabezado ("Resultado medición\nBq/m³"), así que se referencia con cuidado.
COL_CENTRO = "Centro"
COL_AREA = "Área"
COL_ID = "ID"
COL_PLANTA = "Planta"
COL_SALA = "Sala"
COL_CODIGO_SALA = "Código Sala"
COL_PERSONAS = "Personas trabajando en la sala"
COL_PUESTOS = "Puestos en la sala"
COL_CODIGO_DETECTOR = "Código"
COL_FECHA_COLOCACION = "Fecha de colocación"
COL_FECHA_RETIRADA = "Fecha de retirada real"
COL_PLANO = "Plano"
COL_FOTO_SITUACION = "Foto situación"
COL_FOTO_DETECTOR = "Foto detector"

REQUIRED_COLUMNS = [
    COL_CENTRO,
    COL_AREA,
    COL_SALA,
    COL_CODIGO_DETECTOR,
    COL_FECHA_COLOCACION,
    COL_FECHA_RETIRADA,
]


class ExcelFormatError(ValueError):
    """El Excel no tiene el formato de columnas esperado."""


def _find_result_column(columns: list[str]) -> str:
    """La columna de resultado puede traer el salto de línea como \\n real
    o como espacio, según cómo se guardó el Excel. Se busca por prefijo."""
    for c in columns:
        normalized = " ".join(str(c).split())
        if normalized.startswith("Resultado medición"):
            return c
    raise ExcelFormatError(
        "No se ha encontrado la columna de resultado de medición "
        "('Resultado medición Bq/m³') en el Excel."
    )


def load_excel(file) -> pd.DataFrame:
    """Carga el Excel subido por el usuario y devuelve un DataFrame normalizado.

    `file` puede ser una ruta o un objeto tipo fichero (p.ej. el que entrega
    st.file_uploader).
    """
    xls = pd.ExcelFile(file)
    sheet_name = "Detectores" if "Detectores" in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet_name)

    result_col = _find_result_column(list(df.columns))
    df = df.rename(columns={result_col: "Resultado Bq/m3"})

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ExcelFormatError(
            "Faltan columnas obligatorias en el Excel: " + ", ".join(missing)
        )

    # Descarta filas totalmente vacías (huecos entre datos)
    df = df.dropna(how="all").reset_index(drop=True)

    # Normaliza fechas a texto dd/mm/yyyy para el informe
    for col in (COL_FECHA_COLOCACION, COL_FECHA_RETIRADA):
        df[col + " fmt"] = df[col].apply(_format_date)

    # Asegura que el resultado es numérico
    df["Resultado Bq/m3"] = pd.to_numeric(df["Resultado Bq/m3"], errors="coerce")

    return df


def _format_date(value) -> str:
    if pd.isna(value):
        return ""
    ts = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(ts):
        return str(value)
    return ts.strftime("%d/%m/%Y")


def group_options(df: pd.DataFrame, group_col: str) -> list[str]:
    """Valores únicos (no vacíos) de la columna de agrupación, para que el
    usuario elija sobre qué centro de trabajo generar el informe."""
    values = df[group_col].dropna().astype(str).unique().tolist()
    return sorted(values)


def filter_group(df: pd.DataFrame, group_col: str, value: str) -> pd.DataFrame:
    return df[df[group_col].astype(str) == str(value)].reset_index(drop=True)
