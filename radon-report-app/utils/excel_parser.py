"""
Lectura y normalización del Excel de detectores de radón (formato actualizado,
con 3 hojas: "Detectores", "Planos" y "Categorías profesionales").

Estructura de la hoja "Detectores":
  - Fila 1 (índice 0): metadatos sueltos (Centro, Área, "Técnico", nombre del
    técnico, "Fecha", fecha, "Dirección", dirección), colocados en columnas
    no siempre alineadas con las columnas de la tabla de datos.
  - Fila 2 (índice 1): cabecera real de la tabla de datos.
  - Fila 3 en adelante: una fila por detector, con columnas como:
      Centro, Área, ID, Planta, Nivel, Sala, Código de la sala,
      Profesionales en la sala, Turno de trabajo, Código,
      Resultado (Bq/m³/h), Incertidumbre,
      Fecha de colocación, Hora de colocación,
      Fecha de retirada óptima, Fecha de retirada real, Hora de retirada real,
      Nombre del plano, Punto X, Punto Y, Plano, Foto situación, Foto detector

Estructura de la hoja "Planos":
  - Fila 1: "Empresa" | <nombre de la empresa/xerencia>
  - Fila 2: "CIF" | <CIF>
  - Filas siguientes (tras una fila en blanco): "Nombre" | "Imagen" (cabecera
    de una tabla de planos/fotos), y debajo el nombre de cada planta con su
    imagen asociada.

Estructura de la hoja "Categorías profesionales":
  - Columnas: "Categoría profesional", "Nº personas expuestas"
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

# --- Columnas de la tabla de detectores -----------------------------------
COL_CENTRO = "Centro"
COL_AREA = "Área"
COL_SALA = "Sala"
COL_CODIGO_SALA = "Código de la sala"
COL_PROFESIONALES_SALA = "Profesionales en la sala"
COL_CODIGO_DETECTOR = "Código"
COL_FECHA_COLOCACION = "Fecha de colocación"
COL_FECHA_RETIRADA_REAL = "Fecha de retirada real"
COL_FECHA_RETIRADA_OPTIMA = "Fecha de retirada óptima"

REQUIRED_COLUMNS = [
    COL_CENTRO,
    COL_AREA,
    COL_SALA,
    COL_CODIGO_DETECTOR,
    COL_FECHA_COLOCACION,
]

METADATA_LABELS = ["Técnico", "Fecha", "Dirección", "Xerencia", "Empresa", "CIF"]


class ExcelFormatError(ValueError):
    """El Excel no tiene el formato esperado."""


def _find_column(columns: list[str], prefix: str) -> str | None:
    for c in columns:
        normalized = " ".join(str(c).split())
        if normalized.lower().startswith(prefix.lower()):
            return c
    return None


def _format_date(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    ts = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(ts):
        return str(value)
    return ts.strftime("%d/%m/%Y")


def _extract_row0_metadata(raw_row0: list) -> dict:
    """Busca en la primera fila (metadatos sueltos) etiquetas conocidas y
    toma como valor la siguiente celda no vacía a su derecha."""
    meta = {}
    n = len(raw_row0)
    for i, val in enumerate(raw_row0):
        if not isinstance(val, str):
            continue
        label = val.strip()
        for known in METADATA_LABELS:
            if label.lower() == known.lower():
                for j in range(i + 1, n):
                    v = raw_row0[j]
                    if v is not None and not (isinstance(v, float) and pd.isna(v)):
                        meta[known] = v
                        break
                break
    return meta


def load_detectores(file) -> tuple[pd.DataFrame, dict]:
    """Carga la hoja 'Detectores'. Devuelve (dataframe_normalizado, metadatos)."""
    xls = pd.ExcelFile(file)
    if "Detectores" not in xls.sheet_names:
        raise ExcelFormatError("El Excel no tiene una hoja llamada 'Detectores'.")

    raw = xls.parse("Detectores", header=None)
    if len(raw) < 2:
        raise ExcelFormatError("La hoja 'Detectores' no tiene suficientes filas.")

    metadata = _extract_row0_metadata(raw.iloc[0].tolist())
    if "Fecha" in metadata:
        metadata["Fecha"] = _format_date(metadata["Fecha"])

    header_row = raw.iloc[1].tolist()
    data = raw.iloc[2:].reset_index(drop=True)
    data.columns = header_row

    missing = [c for c in REQUIRED_COLUMNS if c not in data.columns]
    if missing:
        raise ExcelFormatError(
            "Faltan columnas obligatorias en la hoja 'Detectores': " + ", ".join(missing)
        )

    data = data.dropna(how="all").reset_index(drop=True)

    result_col = _find_column(list(data.columns), "Resultado")
    incert_col = _find_column(list(data.columns), "Incertidumbre")
    data["Resultado Bq/m3"] = (
        pd.to_numeric(data[result_col], errors="coerce") if result_col else pd.NA
    )
    data["Incerteza expandida e K"] = data[incert_col] if incert_col else ""
    data["Incerteza expandida e K"] = data["Incerteza expandida e K"].fillna("")

    for col in (COL_FECHA_COLOCACION, COL_FECHA_RETIRADA_REAL, COL_FECHA_RETIRADA_OPTIMA):
        if col in data.columns:
            data[col + " fmt"] = data[col].apply(_format_date)
        else:
            data[col + " fmt"] = ""

    if COL_CODIGO_SALA not in data.columns:
        data[COL_CODIGO_SALA] = data[COL_SALA]

    return data, metadata


def load_planos_metadata(file) -> dict:
    """Extrae Empresa (Xerencia) y CIF de la hoja 'Planos', si existe."""
    xls = pd.ExcelFile(file)
    if "Planos" not in xls.sheet_names:
        return {}
    raw = xls.parse("Planos", header=None)
    meta = {}
    for _, row in raw.iterrows():
        label = row[0]
        if not isinstance(label, str):
            continue
        label_norm = label.strip().lower()
        value = row[1] if len(row) > 1 else None
        if pd.isna(value):
            continue
        if label_norm == "empresa":
            meta["Empresa"] = value
        elif label_norm == "cif":
            meta["CIF"] = value
    return meta


def load_categorias_profesionales(file) -> pd.DataFrame | None:
    """Carga la hoja 'Categorías profesionales', si existe."""
    xls = pd.ExcelFile(file)
    if "Categorías profesionales" not in xls.sheet_names:
        return None
    df = xls.parse("Categorías profesionales")
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)


def load_workbook(file) -> dict:
    """Carga el libro completo y devuelve un dict con todo lo necesario para
    rellenar el formulario y generar el informe."""
    df, det_meta = load_detectores(file)
    planos_meta = load_planos_metadata(file)
    categorias = load_categorias_profesionales(file)
    return {
        "detectores": df,
        "detectores_meta": det_meta,
        "planos_meta": planos_meta,
        "categorias": categorias,
    }


def group_options(df: pd.DataFrame, group_col: str = COL_CENTRO) -> list[str]:
    values = df[group_col].dropna().astype(str).unique().tolist()
    return sorted(values)


def filter_group(df: pd.DataFrame, value: str, group_col: str = COL_CENTRO) -> pd.DataFrame:
    return df[df[group_col].astype(str) == str(value)].reset_index(drop=True)


def salas_medidas(df: pd.DataFrame) -> list[str]:
    """Lista de todas las salas medidas (para el punto 3 del informe)."""
    if COL_SALA not in df.columns:
        return []
    return sorted({str(s).strip() for s in df[COL_SALA].dropna() if str(s).strip()})


def areas_muestreadas(df: pd.DataFrame) -> list[str]:
    if COL_AREA not in df.columns:
        return []
    return sorted({str(a).strip() for a in df[COL_AREA].dropna() if str(a).strip()})


def categorias_resumen(categorias_df: pd.DataFrame | None) -> tuple[int, str]:
    """Devuelve (total_personas, texto_resumen) a partir de la hoja de
    categorías profesionales, p.ej. (5, 'Médico (2), Enfermería (2), PSG (1)')."""
    if categorias_df is None or categorias_df.empty:
        return 0, ""
    cat_col = _find_column(list(categorias_df.columns), "Categoría")
    num_col = _find_column(list(categorias_df.columns), "Nº") or _find_column(
        list(categorias_df.columns), "Num"
    )
    if not cat_col or not num_col:
        return 0, ""
    total = int(pd.to_numeric(categorias_df[num_col], errors="coerce").fillna(0).sum())
    partes = [
        f"{row[cat_col]} ({int(row[num_col])})"
        for _, row in categorias_df.iterrows()
        if pd.notna(row[cat_col]) and pd.notna(row[num_col])
    ]
    return total, ", ".join(partes)


def merge_resultados(df: pd.DataFrame, resultados_df: pd.DataFrame) -> pd.DataFrame:
    """Combina un Excel de resultados externo (con columna 'Código' y
    columnas de resultado/incertidumbre) dentro de la tabla de detectores,
    haciendo match por el código del detector."""
    resultados_df = resultados_df.copy()
    resultados_df.columns = [str(c).strip() for c in resultados_df.columns]

    codigo_col = _find_column(list(resultados_df.columns), "Código") or _find_column(
        list(resultados_df.columns), "Code"
    )
    resultado_col = _find_column(list(resultados_df.columns), "Resultado")
    incert_col = _find_column(list(resultados_df.columns), "Incertidumbre")

    if not codigo_col:
        raise ExcelFormatError(
            "El Excel de resultados debe tener una columna 'Código' con el código del detector."
        )

    merged = df.copy()
    lookup_resultado = {}
    lookup_incert = {}
    for _, row in resultados_df.iterrows():
        code = str(row[codigo_col]).strip()
        if not code or code.lower() == "nan":
            continue
        if resultado_col and pd.notna(row.get(resultado_col)):
            lookup_resultado[code] = pd.to_numeric(row[resultado_col], errors="coerce")
        if incert_col and pd.notna(row.get(incert_col)):
            lookup_incert[code] = row[incert_col]

    def apply_resultado(r):
        code = str(r[COL_CODIGO_DETECTOR]).strip()
        if code in lookup_resultado and pd.notna(lookup_resultado[code]):
            return lookup_resultado[code]
        return r["Resultado Bq/m3"]

    def apply_incert(r):
        code = str(r[COL_CODIGO_DETECTOR]).strip()
        if code in lookup_incert:
            return lookup_incert[code]
        return r["Incerteza expandida e K"]

    merged["Resultado Bq/m3"] = merged.apply(apply_resultado, axis=1)
    merged["Incerteza expandida e K"] = merged.apply(apply_incert, axis=1)
    return merged


# ---------------------------------------------------------------------------
# Punto 3 del informe: postos de traballo por sala y quendas por categoría
# ---------------------------------------------------------------------------

COL_TURNO = "Turno de trabajo"

TURNO_TRADUCCIONES = {
    "mañana": "mañá",
    "manana": "mañá",
    "tarde": "tarde",
    "noche": "noite",
    "noite": "noite",
    "partido": "partido",
    "continuo": "continuo",
}

# Vocabulario habitual en nomes de salas/postos de traballo de centros
# sanitarios, para traducir automaticamente do castelán ao galego. As claves
# van sen acentos e en minúscula (ver _strip_accents); o valor mantén xa a
# grafía galega correcta.
ES_GL_VOCABULARIO = {
    "consulta": "consulta",
    "consultas": "consultas",
    "enfermeria": "enfermaría",
    "enfermera": "enfermeira",
    "enfermero": "enfermeiro",
    "enfermeras": "enfermeiras",
    "enfermeros": "enfermeiros",
    "medico": "médico",
    "medica": "médica",
    "medicos": "médicos",
    "medicas": "médicas",
    "polivalente": "polivalente",
    "planta": "planta",
    "plantas": "plantas",
    "baja": "baixa",
    "bajo": "baixo",
    "alta": "alta",
    "alto": "alto",
    "sotano": "sótano",
    "rasante": "rasante",
    "primera": "primeira",
    "primero": "primeiro",
    "segunda": "segunda",
    "segundo": "segundo",
    "tercera": "terceira",
    "tercero": "terceiro",
    "cuarta": "cuarta",
    "cuarto": "cuarto",
    "despacho": "despacho",
    "sala": "sala",
    "salas": "salas",
    "espera": "agarda",
    "recepcion": "recepción",
    "urgencias": "urxencias",
    "quirofano": "quirófano",
    "laboratorio": "laboratorio",
    "farmacia": "farmacia",
    "archivo": "arquivo",
    "almacen": "almacén",
    "vacunacion": "vacinación",
    "extracciones": "extraccións",
    "radiologia": "radioloxía",
    "pediatria": "pediatría",
    "direccion": "dirección",
    "administracion": "administración",
    "reunion": "reunión",
    "reuniones": "reunións",
    "vestuario": "vestiario",
    "vestuarios": "vestiarios",
    "pasillo": "corredor",
    "aseo": "aseo",
    "aseos": "aseos",
}


def _capitalizar_como(original: str, traducido: str) -> str:
    if original.isupper():
        return traducido.upper()
    if original[:1].isupper():
        return traducido[:1].upper() + traducido[1:]
    return traducido


def traducir_es_gl(texto: str | None) -> str:
    """Traduce automáticamente al gallego el vocabulario habitual de salas y
    puestos de trabajo (consulta, enfermería, planta baja...), dejando
    intactos números, códigos y palabras no reconocidas (p.ej. acrónimos
    como 'PSG' o nombres propios)."""
    if not texto:
        return texto or ""

    def _repl(match: re.Match) -> str:
        palabra = match.group(0)
        clave = _strip_accents(palabra).lower()
        if clave in ES_GL_VOCABULARIO:
            return _capitalizar_como(palabra, ES_GL_VOCABULARIO[clave])
        return palabra

    return re.sub(r"[A-Za-zÀ-ÿ]+", _repl, texto)


def _strip_accents(text) -> str:
    text = str(text)
    return "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")


def _profession_root(text) -> str:
    """Raíz normalizada (sin acentos, minúscula, 5 caracteres) usada para
    emparejar variantes de género/número: 'Enfermería' y 'Enfermera' dan
    ambas 'enfer'."""
    return _strip_accents(text).lower().strip()[:5]


def _traducir_turno(turno) -> str:
    if turno is None or (isinstance(turno, float) and pd.isna(turno)):
        return ""
    clave = _strip_accents(turno).lower().strip()
    return TURNO_TRADUCCIONES.get(clave, str(turno).strip().lower())


def parse_profesionales_en_sala(value) -> tuple[str | None, int]:
    """'Médico (1)' -> ('Médico', 1); 'PSG' (sin paréntesis) -> ('PSG', 1)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, 0
    text = str(value).strip()
    m = re.match(r"^(.*?)\s*\((\d+)\)\s*$", text)
    if m:
        return m.group(1).strip(), int(m.group(2))
    return text, 1 if text else 0


def postos_traballo_bullets(df: pd.DataFrame, traducir_galego: bool = True) -> list[str]:
    """Viñetas del punto 3: '<Sala>, (<n> <profesión>)', una por sala, a
    partir de las columnas 'Sala' y 'Profesionales en la sala'.

    Si `traducir_galego` es False, se mantiene el texto tal como viene en el
    Excel (castellano), sin traducir al gallego; se usa para generar la
    versión en castellano del informe."""
    if COL_SALA not in df.columns or COL_PROFESIONALES_SALA not in df.columns:
        return []
    bullets: list[str] = []
    seen = set()
    for _, row in df.iterrows():
        sala = row.get(COL_SALA)
        if pd.isna(sala):
            continue
        sala = str(sala).strip()
        if traducir_galego:
            sala = traducir_es_gl(sala)
        name, count = parse_profesionales_en_sala(row.get(COL_PROFESIONALES_SALA))
        if name:
            if traducir_galego:
                name = traducir_es_gl(name)
            display_name = name if name.isupper() else name.lower()
            texto = f"{sala}, ({count} {display_name})"
        else:
            texto = sala
        if texto not in seen:
            seen.add(texto)
            bullets.append(texto)
    return bullets


def categorias_turnos_bullets(
    df: pd.DataFrame, categorias_df: pd.DataFrame | None, traducir_galego: bool = True
) -> tuple[int, list[str]]:
    """Devuelve (total_traballadores, viñetas) del tipo
    'Médico: 2 (turno de mañá)', cruzando la hoja 'Categorías profesionais'
    con la columna 'Turno de trabajo' de Detectores (emparejando por el
    nombre de la profesión que aparece en 'Profesionales en la sala').

    Si `traducir_galego` es False, categoría y turno se muestran tal como
    vienen en el Excel (castellano); se usa para la versión en castellano
    del informe."""
    if categorias_df is None or categorias_df.empty:
        return 0, []

    cat_col = _find_column(list(categorias_df.columns), "Categoría")
    num_col = _find_column(list(categorias_df.columns), "Nº") or _find_column(
        list(categorias_df.columns), "Num"
    )
    if not cat_col or not num_col:
        return 0, []

    turno_col = _find_column(list(df.columns), "Turno")
    room_profesion_turno: list[tuple[str, str]] = []
    if turno_col and COL_PROFESIONALES_SALA in df.columns:
        for _, row in df.iterrows():
            name, _ = parse_profesionales_en_sala(row.get(COL_PROFESIONALES_SALA))
            if name:
                room_profesion_turno.append((name, row.get(turno_col)))

    total = 0
    bullets: list[str] = []
    for _, row in categorias_df.iterrows():
        categoria = row.get(cat_col)
        n = row.get(num_col)
        if pd.isna(categoria) or pd.isna(n):
            continue
        n = int(n)
        total += n
        root = _profession_root(categoria)
        turnos_raw = sorted(
            {
                str(t).strip()
                for (name, t) in room_profesion_turno
                if _profession_root(name) == root and pd.notna(t) and str(t).strip()
            }
        )
        if traducir_galego:
            turnos = sorted({_traducir_turno(t) for t in turnos_raw if _traducir_turno(t)})
        else:
            turnos = [t.lower() for t in turnos_raw]
        turno_text = " / ".join(turnos)
        categoria_txt = (
            traducir_es_gl(str(categoria).strip()) if traducir_galego else str(categoria).strip()
        )
        if turno_text:
            bullets.append(f"{categoria_txt}: {n} (turno de {turno_text})")
        else:
            bullets.append(f"{categoria_txt}: {n}")
    return total, bullets
