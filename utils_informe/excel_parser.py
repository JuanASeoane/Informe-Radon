"""
Lectura y normalización del Excel de detectores de radón (formato actualizado,
con 3 hojas: "Detectores", "Planos" y "Categorías profesionales").

Estructura de la hoja "Detectores":
  - Fila 1 (índice 0): metadatos sueltos (Centro, Área / Zona, "Técnico", nombre del
    técnico, "Fecha", fecha, "Dirección", dirección), colocados en columnas
    no siempre alineadas con las columnas de la tabla de datos.
  - Fila 2 (índice 1): cabecera real de la tabla de datos.
  - Fila 3 en adelante: una fila por detector, con columnas como:
      Centro, Área / Zona, ID, Planta, Nivel, Sala, Código de la sala,
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
COL_AREA = "Área / Zona"
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
    toma como valor la celda INMEDIATAMENTE siguiente (aunque esté vacía).
    No sigue buscando más a la derecha: si lo hiciera, un campo vacío
    (p.ej. "Técnico" sin nombre) haría que se tomara por error el texto
    de la siguiente etiqueta ("Fecha") como si fuera su valor."""
    meta = {}
    n = len(raw_row0)
    for i, val in enumerate(raw_row0):
        if not isinstance(val, str):
            continue
        label = val.strip()
        for known in METADATA_LABELS:
            if label.lower() == known.lower():
                if i + 1 < n:
                    v = raw_row0[i + 1]
                    if v is not None and not (isinstance(v, float) and pd.isna(v)):
                        meta[known] = v
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
    """Extrae Empresa (Xerencia), CIF y los datos del informe final
    (superficie, plantas, comunicación a los trabajadores) de la hoja
    'Planos', si existen."""
    xls = pd.ExcelFile(file)
    if "Planos" not in xls.sheet_names:
        return {}
    raw = xls.parse("Planos", header=None)
    meta = {}
    etiquetas = {
        "empresa": "Empresa",
        "cif": "CIF",
        "superficie construida": "Superficie construida",
        "superficie útil": "Superficie útil",
        "nº de plantas": "Nº de plantas",
        "fecha comunicación trabajadores": "Fecha comunicación trabajadores",
        "medio de comunicación": "Medio de comunicación",
    }
    for _, row in raw.iterrows():
        label = row[0]
        if not isinstance(label, str):
            continue
        label_norm = label.strip().lower()
        value = row[1] if len(row) > 1 else None
        if pd.isna(value):
            continue
        if label_norm in etiquetas:
            meta[etiquetas[label_norm]] = value
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


def extraer_resultados_pdf_laboratorio(pdf_bytes) -> pd.DataFrame:
    """Lee un PDF de resultados de laboratorio (como los que genera CYE
    Control y Estudios: una tabla por centro, con filas "Nº Código
    Ubicación... Exposición ±incert Concentración ±incert") y devuelve
    un DataFrame con columnas 'Código', 'Resultado' e 'Incertidumbre',
    listo para pasar a merge_resultados junto con los detectores del
    centro actual (solo se usarán los códigos que coincidan).

    Se apoya solo en pypdf (ya es una dependencia de la app) para leer
    el texto, sin depender de programas externos del sistema como
    poppler, que no tiene por qué estar disponible en todos los
    sitios donde se instale la app (p.ej. Termux)."""
    from pypdf import PdfReader

    reader = PdfReader(pdf_bytes)
    texto = ""
    for pagina in reader.pages:
        try:
            texto += (pagina.extract_text() or "") + "\n"
        except Exception:
            continue

    # Cada fila real de la tabla trae el nº de orden y el código del
    # detector juntos al principio de línea (p.ej. "1 GJ4133 Planta 0.
    # CS/0/01"); cada resultado son 4 números seguidos, dos de ellos
    # con el símbolo ± delante (exposición ±incert., concentración
    # ±incert.) — se cogen el 3º y 4º de cada grupo (la concentración
    # y su incertidumbre), no los dos primeros (que son de exposición,
    # una magnitud distinta).
    pares_codigo = re.findall(r"(?:^|\n)\s*\d{1,2}\s+([A-Z]{2}\d{3,5}\*?)\s", texto)
    filas_resultado = re.findall(r"\d+\s*±\s*\d+\s+(\d+)\s*±\s*(\d+)", texto)

    filas = []
    for codigo, (concentracion, incertidumbre) in zip(pares_codigo, filas_resultado):
        filas.append({
            "Código": codigo.rstrip("*").strip(),
            "Resultado": concentracion,
            "Incertidumbre": incertidumbre,
        })
    return pd.DataFrame(filas, columns=["Código", "Resultado", "Incertidumbre"])


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
    "pac": "PAC",
    "rotatorio": "rotatorio",
    "rotatorio complejo": "rotatorio complexo",
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


def _pluralizar_palabra(palabra: str) -> str:
    """Pluralización razonable en galego/castelán para nomes de
    categorías profesionais (regras xerais, sen pretender cubrir
    todas as excepcións do idioma)."""
    if not palabra:
        return palabra
    ultima = palabra[-1].lower()
    if ultima == "s":
        return palabra
    if ultima == "z":
        return palabra[:-1] + "ces"
    if ultima in "aeiouáéíóúàèìòù":
        return palabra + "s"
    return palabra + "es"


def _pluralizar_categoria(categoria: str) -> str:
    """Pluraliza só a primeira palabra do nome da categoría (p.ex.
    'Auxiliar de enfermaría' -> 'Auxiliares de enfermaría'), que adoita
    ser o núcleo do sintagma nestes nomes de profesións compostos.
    Se esa primeira palabra é unha sigla en maiúsculas (p.ex. "TER",
    "UCI"), non se pluraliza nunca."""
    categoria = categoria.strip()
    if not categoria:
        return categoria
    partes = categoria.split(" ", 1)
    primera_palabra = partes[0]
    resto = partes[1] if len(partes) > 1 else ""
    if primera_palabra.isupper():
        return categoria
    primeira = _pluralizar_palabra(primera_palabra)
    return f"{primeira} {resto}".strip() if resto else primeira


def _traducir_turno(turno) -> str:
    """Traduce un turno (posiblemente composto, p.ex. 'Mañana + tarde +
    noche') ao galego, unindo as partes con ' - ' en vez de ' + '."""
    if turno is None or (isinstance(turno, float) and pd.isna(turno)):
        return ""
    partes = [p.strip() for p in str(turno).split("+") if p.strip()]
    if not partes:
        return ""
    traducidas = []
    for p in partes:
        clave = _strip_accents(p).lower().strip()
        traducidas.append(TURNO_TRADUCCIONES.get(clave, p.lower()))
    return " - ".join(traducidas)


def _normalizar_turno_es(turno) -> str:
    """Igual que _traducir_turno pero sin traducir, solo normalizando el
    separador '+' a ' - ' (para la versión en castellano)."""
    if turno is None or (isinstance(turno, float) and pd.isna(turno)):
        return ""
    partes = [p.strip().lower() for p in str(turno).split("+") if p.strip()]
    return " - ".join(partes)


def parse_profesionales_en_sala(value) -> tuple[str | None, int]:
    """'Médico (1)' -> ('Médico', 1); 'PSG' (sin paréntesis) -> ('PSG', 1)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, 0
    text = str(value).strip()
    m = re.match(r"^(.*?)\s*\((\d+)\)\s*$", text)
    if m:
        return m.group(1).strip(), int(m.group(2))
    return text, 1 if text else 0


def parse_profesionales_multiples(value) -> list[tuple[str, int, str]]:
    """Igual que parse_profesionales_en_sala, pero admite varias
    categorías en la misma sala, cada una con su propio turno
    (p.ej. 'Enfermería (2) - Mañana, Celador (1) - Mañana + tarde'
    -> [('Enfermería', 2, 'Mañana'), ('Celador', 1, 'Mañana + tarde')]).
    Si una línea no lleva turno (formato antiguo, sin turno por
    categoría), el turno devuelto es "". Se admiten tanto comas como
    saltos de línea como separador entre categorías, porque la app
    guarda el valor con comas, pero al pasar por el Excel se
    convierten en saltos de línea para que la celda se vea bien
    ajustada (ver _incrustar_en_hoja / generar_excel)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    partes = [p.strip() for p in re.split(r"[,\n]", str(value)) if p.strip()]
    resultado = []
    for parte in partes:
        turno_parte = ""
        if " - " in parte:
            parte, turno_parte = parte.split(" - ", 1)
            parte = parte.strip()
            turno_parte = turno_parte.strip()
        nombre, cuenta = parse_profesionales_en_sala(parte)
        if nombre:
            resultado.append((nombre, cuenta, turno_parte))
    return resultado


def postos_traballo_bullets(
    df: pd.DataFrame, traducir_galego: bool = True, es_atencion_primaria: bool = False
) -> list[str]:
    """Viñeta del punto 3 por sala: '<Sala>: <n> <profesión> (<horario
    del turno>) +<n> <profesión> (<horario del turno>)...', agrupando
    todas las categorías profesionales de una misma sala (aunque
    vengan de detectores distintos) en una única línea, cada una con
    su propio turno si lo tiene. Sigue el mismo criterio que la
    segunda parte del punto 3: para mañana/tarde/noche se pone
    directamente el horario (sin decir "turno de..."); "Horario PAC"
    se sustituye por "Turnos PAC"; rotatorio y rotatorio complejo se
    quedan tal cual.

    `es_atencion_primaria` cambia el horario de "Tarde": de 15 a 22 h
    normalmente, o de 14 a 21 h si el centro es de Atención Primaria.

    Si `traducir_galego` es False, se mantiene el texto tal como viene en el
    Excel (castellano), sin traducir al gallego; se usa para generar la
    versión en castellano del informe."""
    if COL_SALA not in df.columns or COL_PROFESIONALES_SALA not in df.columns:
        return []
    agrupado: dict[str, list[tuple[int, str, str]]] = {}
    orden_salas: list[str] = []
    for _, row in df.iterrows():
        sala = row.get(COL_SALA)
        if pd.isna(sala):
            continue
        sala = str(sala).strip()
        if traducir_galego:
            sala = traducir_es_gl(sala)
        if sala not in agrupado:
            agrupado[sala] = []
            orden_salas.append(sala)
        entradas = parse_profesionales_multiples(row.get(COL_PROFESIONALES_SALA))
        for name, count, turno in entradas:
            if traducir_galego:
                name = traducir_es_gl(name)
            display_name = name if name.isupper() else name.lower()
            if count and count > 1:
                display_name = _pluralizar_categoria(display_name)
            turno_fmt = _formatear_turno_categoria(turno, traducir_galego, es_atencion_primaria)
            agrupado[sala].append((count, display_name, turno_fmt))

    bullets: list[str] = []
    for sala in orden_salas:
        entradas_sala = agrupado[sala]
        if not entradas_sala:
            bullets.append(sala)
            continue
        partes = []
        for count, display_name, turno_fmt in entradas_sala:
            texto_parte = f"{count} {display_name}"
            if turno_fmt:
                texto_parte += f" ({turno_fmt})"
            partes.append(texto_parte)
        bullets.append(f"{sala}: " + " +".join(partes))
    return bullets


_HORARIO_POR_TURNO = {
    "manana": "8 a 15 h",
    "noche": "22 a 8 h",
}


def _formatear_turno_categoria(turno_raw, traducir_galego: bool, es_atencion_primaria: bool = False) -> str:
    """Frase completa lista para meter entre paréntesis en el informe.
    Para mañana, tarde y noche se deja solo el horario, sin decir
    "turno de..." (p.ej. '(2 médicos 8 a 15 h)'): mañana -> "8 a 15
    h", tarde -> "15 a 22 horas" (o "14 a 21 h" si el centro es
    exactamente de Atención Primaria, no de PAC), noche -> "22 a 8
    h". "Horario PAC" se sustituye entero por "Turnos PAC". Rotatorio
    y rotatorio complejo se quedan tal cual, con el "turno de" delante
    (no tienen horario fijo que poner en su lugar). Si hay varios
    turnos combinados con "/" (p.ej. "Mañana/Tarde"), se ponen los dos
    horarios seguidos, separados por "/"."""
    if turno_raw is None or (isinstance(turno_raw, float) and pd.isna(turno_raw)) or not str(turno_raw).strip():
        return ""
    texto = str(turno_raw).strip()
    if _strip_accents(texto).lower() == "horario pac":
        return "Turnos PAC"

    horario_tarde = "14 a 21 h" if es_atencion_primaria else "15 a 22 horas"

    partes = [p.strip() for p in texto.split("/") if p.strip()]
    partes_fmt = []
    todas_con_horario = True
    for parte in partes:
        clave_parte = _strip_accents(parte).lower()
        horario = horario_tarde if clave_parte == "tarde" else _HORARIO_POR_TURNO.get(clave_parte)
        if horario:
            partes_fmt.append(horario)
            continue
        todas_con_horario = False
        palabras_fmt = []
        for palabra in parte.split():
            if palabra.isupper():
                palabras_fmt.append(palabra)
            elif traducir_galego:
                clave_palabra = _strip_accents(palabra).lower().strip()
                palabras_fmt.append(TURNO_TRADUCCIONES.get(clave_palabra, palabra.lower()))
            else:
                palabras_fmt.append(palabra.lower())
        partes_fmt.append(" ".join(palabras_fmt))

    cuerpo = " / ".join(partes_fmt)
    if todas_con_horario:
        # Mañana/tarde/noche: solo el horario, sin "turno de" delante.
        return cuerpo
    conjuncion = "turnos de" if len(partes) > 1 else "turno de"
    return f"{conjuncion} {cuerpo}"


def categorias_turnos_bullets(
    categorias_df: pd.DataFrame | None, traducir_galego: bool = True, es_atencion_primaria: bool = False
) -> tuple[int, list[str]]:
    """Devuelve (total_traballadores, viñetas) del tipo
    '2 médicos (turnos de mañana / tarde)' o '1 celador (turno de
    mañana)', a partir del campo "Turno" propio de cada categoría
    profesional (hoja "Categorías profesionales"). Este turno se
    introduce directamente en "Categorías profesionales" y NO tiene
    ninguna relación con los turnos que se puedan poner por sala en
    cada detector: es un dato aparte, pensado solo para esta parte
    del informe.

    `es_atencion_primaria` cambia el horario que se muestra para el
    turno de "Tarde": de 15 a 22 h normalmente, o de 14 a 21 h si el
    centro es exactamente de Atención Primaria (no cuenta si es solo
    PAC).

    Si `traducir_galego` es False, categoría y turno se muestran tal como
    vienen en el Excel (castellano); se usa para la versión en castellano
    del informe."""
    if categorias_df is None or categorias_df.empty:
        return 0, []

    cat_col = _find_column(list(categorias_df.columns), "Categoría")
    num_col = _find_column(list(categorias_df.columns), "Nº") or _find_column(
        list(categorias_df.columns), "Num"
    )
    turno_col = _find_column(list(categorias_df.columns), "Turno")
    if not cat_col or not num_col:
        return 0, []

    total = 0
    bullets: list[str] = []
    for _, row in categorias_df.iterrows():
        categoria = row.get(cat_col)
        n = row.get(num_col)
        if pd.isna(categoria) or pd.isna(n):
            continue
        n = int(n)
        total += n
        categoria_txt_base = (
            traducir_es_gl(str(categoria).strip()) if traducir_galego else str(categoria).strip()
        )
        categoria_txt_base = categoria_txt_base if categoria_txt_base.isupper() else categoria_txt_base.lower()
        categoria_txt = _pluralizar_categoria(categoria_txt_base) if n > 1 else categoria_txt_base

        turno_raw = row.get(turno_col) if turno_col else None
        turno_fmt = _formatear_turno_categoria(turno_raw, traducir_galego, es_atencion_primaria)

        if not turno_fmt:
            bullets.append(f"{n} {categoria_txt}")
        else:
            bullets.append(f"{n} {categoria_txt} ({turno_fmt})")
    return total, bullets
