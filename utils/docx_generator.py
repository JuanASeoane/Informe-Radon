"""
Generación del "Informe de resultados de medicións de radón no centro de
traballo" siguiendo la plantilla del Servizo Galego de Saúde / UPRL, con la
documentación de referencia do Consello de Seguridade Nuclear (CSN).

Genera un .docx en memoria (BytesIO) a partir de:
  - los datos de cabecera introducidos en el formulario (dict `context`)
  - la tabla de resultados de mediciones (DataFrame `df`, ya filtrado a un
    único centro de traballo)
  - opcionalmente, imágenes de logotipo para la cabecera institucional
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from utils.excel_parser import traducir_es_gl

NIVEL_REFERENCIA_BQ_M3 = 300

REFERENCIAS = [
    "Lei 31/1995, de 8 de novembro, de Prevención de Riscos Laborais (BOE nº 269, de 10 de novembro).",
    "Real Decreto 39/1997, de 17 de xaneiro, polo que se aproba o Regulamento dos servizos de "
    "prevención (BOE nº 27, de 31 de xaneiro).",
    "Real Decreto 732/2019, de 20 de decembro, polo que se modifica o Código Técnico da "
    "Edificación, aprobado polo Real Decreto 314/2006, de 17 de marzo (BOE nº 311, de 27 de "
    "decembro, páxinas 140488 a 140674).",
    "Instrución IS-47, de 9 de abril de 2025, do Consello de Seguridade Nuclear pola que se "
    "aproba o listado de términos municipais de actuación prioritaria contra o radón e establéce.",
    "Guía de Seguridade do CSN 11.4 Metodoloxía para a avaliación da exposición ao radon nos "
    "lugares de traballo.",
    "Estratexia para reducir a exposición ao radon en Galicia. Estratexia Reduce Radon 2025-2030. "
    "Consellería de Sanidade, Dirección Xeral de Saúde Pública, ano 2024. "
    "https://airesaude.sergas.gal/Contidos/Documents/135/Estratexia%20reduce%20radon%202025-2030.pdf",
]

ANEXOS = [
    "ANEXO I: FORMULARIOS TOMA DE DATOS",
    "ANEXO II: ESQUEMA GRÁFICO DO EDIFICIO E PLANOS DE CADA PLANTA",
    "ANEXO III: INFORME DE ENSAIO DO LABORATORIO ACREDITADO",
    "ANEXO IV: CERTIFICADO ENAC DO LABORATORIO ACREDITADO",
]

TABLE_HEADERS = [
    "Código zona de mostraxe",
    "Código detector",
    "Data de inicio exposición",
    "Data fin de exposición",
    "Concentración radón (Bq/m³) (*)",
    "Incerteza expandida e K",
    "Sala / Posto de traballo asociado",
]


@dataclass
class ReportContext:
    xerencia: str = ""
    cif: str = ""
    centro: str = ""
    enderezo: str = ""
    superficie_construida: str = ""
    superficie_util: str = ""
    num_plantas: str = ""
    postos_bullets: list = field(default_factory=list)
    num_traballadores_total: str = ""
    categorias_bullets: list = field(default_factory=list)
    notas_adicionais: str = ""
    data_informacion_traballadores: str = ""
    medio_informacion_traballadores: str = "correo electrónico"
    data_informe: str = ""
    tecnico_nome: str = ""
    servizo_unidade: str = ""
    incertezas_por_defecto: str = ""
    conclusion_manual: str = ""  # si el usuario quiere sobreescribir el texto automático
    logo: bytes | None = None  # imagen única de cabecera (ancho completo)


# ---------------------------------------------------------------------------
# Helpers de bajo nivel
# ---------------------------------------------------------------------------

def _set_cell_shading(cell, color_hex: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shd)


def _set_cell_borders(cell, sz=4, color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def _set_cell_vertical_margins(cell, top_twips: int = 100, bottom_twips: int = 100):
    """Añade espacio interno arriba/abajo dentro de la celda para que el
    texto no quede pegado a los bordes."""
    tcPr = cell._tc.get_or_add_tcPr()
    mar = tcPr.find(qn("w:tcMar"))
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcPr.append(mar)
    for edge, value in (("top", top_twips), ("bottom", bottom_twips)):
        el = mar.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            mar.append(el)
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")


def _set_table_fixed_layout(table, total_width_cm: float):
    """Fuerza layout de tabla fijo para que Word respete los anchos de celda
    en vez de autoajustarlos al contenido."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    dxa = int(total_width_cm * 566.929)  # 1 cm ≈ 566.929 twips (dxa)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(dxa))


def _fit_image_dims(image_bytes: bytes, max_width_cm: float, max_height_cm: float):
    """Calcula (width, height) en EMU para que la imagen quepa dentro del
    recuadro máximo indicado, conservando la proporción original."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        w_px, h_px = img.size
    except Exception:
        return Cm(max_width_cm), Cm(max_height_cm)

    aspect = w_px / h_px if h_px else 1
    max_width = Cm(max_width_cm)
    max_height = Cm(max_height_cm)

    height = max_height
    width = int(height * aspect)
    if width > max_width:
        width = max_width
        height = int(width / aspect)
    return width, height


def _add_field(paragraph, field_code: str, cached_text: str = "1", font_size_pt: float = 9):
    """Inserta un campo de Word (p.ej. PAGE, NUMPAGES) en un párrafo.

    Cada pieza del campo (begin / instrText / separate / valor en caché / end)
    debe ir en su propio <w:r>, si no Word no lo reconoce como un campo válido
    y no lo recalcula (aparece en blanco). Se incluye además un valor en
    caché para que se vea algo razonable antes del primer recálculo."""

    def _new_run():
        return paragraph.add_run()

    def _set_size(run):
        run.font.size = Pt(font_size_pt)

    r_begin = _new_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r_begin._r.append(fld_begin)

    r_instr = _new_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_code} "
    r_instr._r.append(instr)

    r_sep = _new_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r_sep._r.append(fld_sep)

    r_cached = _new_run()
    r_cached.text = cached_text
    _set_size(r_cached)

    r_end = _new_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end._r.append(fld_end)


def _heading(doc, number: str, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"{number}. {text}")
    run.bold = True
    run.font.size = Pt(11)
    return p


def _body(doc, text: str):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(6)
    for run in p.runs:
        run.font.size = Pt(10.5)
    return p


def _set_font(doc):
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10.5)


# ---------------------------------------------------------------------------
# Cabecera institucional
# ---------------------------------------------------------------------------

def _build_header(section, ctx: ReportContext):
    header = section.header
    header.is_linked_to_previous = False
    # limpia el párrafo por defecto
    for p in list(header.paragraphs):
        p.text = ""

    table = header.add_table(rows=1, cols=2, width=Cm(17))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_fixed_layout(table, 17)
    LOGO_CELL_WIDTH_CM = 13.2
    DATE_CELL_WIDTH_CM = 3.8
    widths = [Cm(LOGO_CELL_WIDTH_CM), Cm(DATE_CELL_WIDTH_CM)]
    for cell, w in zip(table.rows[0].cells, widths):
        cell.width = w
        _set_cell_borders(cell)

    logo_cell = table.rows[0].cells[0]
    logo_cell.vertical_alignment = 1
    para = logo_cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if ctx.logo:
        try:
            run = para.add_run()
            # ajusta el logotipo al espacio disponible en la celda,
            # conservando su proporción original
            width, height = _fit_image_dims(
                ctx.logo, max_width_cm=LOGO_CELL_WIDTH_CM - 0.6, max_height_cm=1.8
            )
            run.add_picture(io.BytesIO(ctx.logo), width=width, height=height)
        except Exception:
            para.add_run("[logotipo]").italic = True
    else:
        r = para.add_run("UPRL · SERVIZO GALEGO DE SAÚDE · ÁREA SANITARIA")
        r.font.size = Pt(8)
        r.italic = True

    data_cell = table.rows[0].cells[1]
    data_cell.vertical_alignment = 1
    dp = data_cell.paragraphs[0]
    dp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = dp.add_run(f"Data: {ctx.data_informe}\n")
    r1.font.size = Pt(9)
    dp2 = data_cell.add_paragraph()
    r2 = dp2.add_run("Páxina ")
    r2.font.size = Pt(9)
    _add_field(dp2, "PAGE", cached_text="1", font_size_pt=9)
    r3 = dp2.add_run(" de ")
    r3.font.size = Pt(9)
    _add_field(dp2, "NUMPAGES", cached_text="1", font_size_pt=9)

    # Título del informe, en caja con borde
    title_table = header.add_table(rows=1, cols=1, width=Cm(17))
    _set_table_fixed_layout(title_table, 17)
    cell = title_table.rows[0].cells[0]
    _set_cell_borders(cell, sz=8)
    _set_cell_shading(cell, "F2F2F2")
    tp = cell.paragraphs[0]
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(
        f"INFORME DE RESULTADOS MEDICIÓNS DE RADON NO CENTRO DE TRABALLO DE "
        f"{ctx.centro.upper() or '...................'}"
    )
    tr.bold = True
    tr.font.size = Pt(11)

    # espaciador extra al final de la cabecera para separarla del cuerpo
    spacer = header.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    spacer_run = spacer.add_run("")
    spacer_run.font.size = Pt(6)


# ---------------------------------------------------------------------------
# Cuerpo del informe
# ---------------------------------------------------------------------------

def _section_1(doc, ctx: ReportContext):
    _heading(doc, "1", "IDENTIFICACIÓN DO CENTRO DE TRABALLO")
    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"

    # fila XERENCIA + CIF en la misma línea, como en la plantilla
    cells0 = table.rows[0].cells
    cells0[0].text = "XERENCIA"
    cells0[0].paragraphs[0].runs[0].bold = True
    merged01 = cells0[1].merge(cells0[2])
    merged01.text = ctx.xerencia
    cells0[3].text = f"CIF: {ctx.cif}"

    cells1 = table.rows[1].cells
    cells1[0].text = "CENTRO"
    cells1[0].paragraphs[0].runs[0].bold = True
    merged1 = cells1[1].merge(cells1[2]).merge(cells1[3])
    merged1.text = ctx.centro

    cells2 = table.rows[2].cells
    cells2[0].text = "SERVIZO / UNIDADE MOSTREXADA"
    cells2[0].paragraphs[0].runs[0].bold = True
    merged2b = cells2[1].merge(cells2[2]).merge(cells2[3])
    merged2b.text = ctx.servizo_unidade

    cells3 = table.rows[3].cells
    cells3[0].text = "ENDEREZO"
    cells3[0].paragraphs[0].runs[0].bold = True
    merged2 = cells3[1].merge(cells3[2]).merge(cells3[3])
    merged2.text = ctx.enderezo

    cells4 = table.rows[4].cells
    cells4[0].text = "DESCRICIÓN DO CENTRO"
    cells4[0].paragraphs[0].runs[0].bold = True
    cells4[1].text = f"Superficie construída: {ctx.superficie_construida}"
    cells4[2].text = f"Superficie útil: {ctx.superficie_util}"
    cells4[3].text = f"N.º plantas: {ctx.num_plantas}"

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9.5)

    # primera columna (etiquetas) con fondo gris claro
    for row_cells in (cells0, cells1, cells2, cells3, cells4):
        _set_cell_shading(row_cells[0], "F2F2F2")

    # algo más de aire vertical dentro de cada celda y una altura mínima
    # de fila, para que el texto no quede pegado a las líneas de la tabla
    for row in table.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.height = Cm(0.9)
        for cell in row.cells:
            _set_cell_vertical_margins(cell, top_twips=120, bottom_twips=120)

    doc.add_paragraph()


def _section_2(doc, ctx: ReportContext):
    _heading(doc, "2", "OBXECTO DO INFORME")
    _body(
        doc,
        f"O presente informe ten como obxecto recoller os resultados das medicións de radon "
        f"no centro de traballo de {ctx.centro or '.................'}, para, no caso de que a "
        f"exposición a radon sexa superior ao nivel de referencia establecido "
        f"({NIVEL_REFERENCIA_BQ_M3} Bq/m³), establecer as medidas de remediación necesarias, "
        f"para diminuír ou eliminar estes niveis.",
    )


def _bullet_list(doc, items: list[str]):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(item)
        r.font.size = Pt(10)


def _section_3(doc, ctx: ReportContext):
    _heading(doc, "3", "INFORMACIÓN SOBRE OS/AS TRABALLADORES/AS")
    _body(
        doc,
        f"O hospital/centro de saúde de {ctx.centro or '.....'}, consta dos seguintes postos de "
        f"traballo e a seguinte ocupación por espazo de traballo, susceptibles de exposición a "
        f"radón pola súa situación en planta ou baixo cuberta:",
    )
    if ctx.postos_bullets:
        _bullet_list(doc, ctx.postos_bullets)
    else:
        _body(doc, ".....")

    _body(
        doc,
        f"O número de traballadores adscritos a este centro é de "
        f"{ctx.num_traballadores_total or '.....'}, no tipo de quendas que figuran a continuación:",
    )
    if ctx.categorias_bullets:
        _bullet_list(doc, ctx.categorias_bullets)
    else:
        _body(doc, ".....")

    if ctx.notas_adicionais.strip():
        _body(doc, ctx.notas_adicionais.strip())

    _body(
        doc,
        f"Os traballadores/as do centro de traballo {ctx.centro or '.....'} foron informados da "
        f"realización das medicións e da súa finalidade mediante "
        f"{ctx.medio_informacion_traballadores or 'correo electrónico'} "
        f"o día {ctx.data_informacion_traballadores or '.....'}.",
    )


def _section_4(doc):
    _heading(doc, "4", "CONDICIÓNS DA EXPOSICIÓN")
    _body(
        doc,
        "No/s Formulario/s de toma de datos do Anexo I recóllense: o tempo de exposición de "
        "cada detector, o responsable da súa colocación e retirada, a comprobación do estado "
        "dos detectores, as condicións de temperatura ou humidade, e observacións sobre o "
        "sistema de ventilación.",
    )


def _section_5(doc):
    _heading(doc, "5", "PLANOS")
    _body(
        doc,
        "O esquema gráfico do edificio e planos de cada planta móstranse no Anexo II, onde se "
        "indican as zonas de mostraxe e as localizacións dos detectores; xunto co código de "
        "cada zona de mostraxe, márcase o código do detector correspondente.",
    )


def _row_value(row, *keys, default=""):
    for k in keys:
        if k in row and pd.notna(row[k]):
            return row[k]
    return default


def _section_7(doc, df: pd.DataFrame, ctx: ReportContext):
    _heading(doc, "7", "RESULTADOS DAS MEDICIÓNS REALIZADAS")
    _body(doc, "Os resultados das medicións de radon realizadas poden verse na seguinte táboa:")

    table = doc.add_table(rows=1, cols=len(TABLE_HEADERS))
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(TABLE_HEADERS):
        hdr_cells[i].text = h
        _set_cell_shading(hdr_cells[i], "D9D9D9")
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(8.5)

    any_exceeds = False
    for _, row in df.iterrows():
        codigo_zona = _row_value(row, "Código de la sala", "Código Sala", "Sala")
        codigo_zona = "" if pd.isna(codigo_zona) else str(codigo_zona).rstrip(".0")
        codigo_det = _row_value(row, "Código")
        fecha_ini = _row_value(row, "Fecha de colocación fmt")
        fecha_fin = _row_value(row, "Fecha de retirada real fmt")
        concentracion = row.get("Resultado Bq/m3")
        # última columna: sala + posto de traballo, en dúas liñas na mesma cela.
        # O número entre parénteses do posto (p.ex. "Médico (1)") indica o
        # número de persoas desa categoría nesa sala.
        sala_nome = _row_value(row, "Sala", default="")
        sala_gl = traducir_es_gl(str(sala_nome).strip()) if pd.notna(sala_nome) and str(sala_nome).strip() else ""
        puestos = _row_value(row, "Profesionales en la sala", "Puestos en la sala", default="")
        if pd.notna(puestos) and not isinstance(puestos, str):
            puestos = str(int(puestos)) if float(puestos).is_integer() else str(puestos)
        puestos_gl = traducir_es_gl(str(puestos).strip()) if puestos else ""

        incerteza = row.get("Incerteza expandida e K", ctx.incertezas_por_defecto)
        if pd.isna(incerteza) or incerteza is None:
            incerteza = ctx.incertezas_por_defecto or ""

        cells = table.add_row().cells
        values = [
            str(codigo_zona) if codigo_zona else "",
            str(codigo_det) if pd.notna(codigo_det) else "",
            str(fecha_ini),
            str(fecha_fin),
            f"{concentracion:g}" if pd.notna(concentracion) else "",
            str(incerteza),
        ]
        exceeds = pd.notna(concentracion) and concentracion > NIVEL_REFERENCIA_BQ_M3
        any_exceeds = any_exceeds or exceeds

        for i, val in enumerate(values):
            cells[i].text = val
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.size = Pt(9)
                if exceeds and i == 4:  # columna de concentración
                    r.bold = True
                    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

        # última columna: sala (liña 1) e posto de traballo (liña 2)
        last_cell = cells[6]
        last_cell.text = ""
        p1 = last_cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(sala_gl)
        r1.font.size = Pt(9)
        r1.bold = True
        if puestos_gl:
            p2 = last_cell.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(puestos_gl)
            r2.font.size = Pt(9)

    note = doc.add_paragraph()
    note.paragraph_format.space_before = Pt(4)
    nr = note.add_run(
        f"* Resáltanse en negriña e letra vermella os valores superiores a "
        f"{NIVEL_REFERENCIA_BQ_M3} Bq/m³."
    )
    nr.italic = True
    nr.font.size = Pt(8.5)
    return any_exceeds


def _section_8(doc, ctx: ReportContext, any_exceeds: bool, exceeded_rooms: list[str]):
    _heading(doc, "8", "CONCLUSIÓNS")
    if ctx.conclusion_manual.strip():
        _body(doc, ctx.conclusion_manual.strip())
        return
    if not any_exceeds:
        _body(
            doc,
            "A vista dos datos do apartado 7, non hai ningún posto de traballo nos que se "
            "supere o nivel de referencia.",
        )
    else:
        salas = ", ".join(exceeded_rooms) if exceeded_rooms else "indicados no apartado 7"
        _body(
            doc,
            f"As medicións de radon nos seguintes postos de traballo: {salas}, dan valores "
            f"superiores ao nivel de referencia. Para estes casos estudiaranse medidas de "
            f"remediación, completaranse as medicións nas plantas superiores e/ou "
            f"efectuaranse medicións en continuo, segundo proceda.",
        )


def _section_9(doc):
    _heading(doc, "9", "DOCUMENTACIÓN DE REFERENCIA")
    for ref in REFERENCIAS:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(ref)
        run.font.size = Pt(9.5)


def _section_10(doc, ctx: ReportContext):
    _heading(doc, "10", "DATA E FIRMA DO TÉCNICO/A")
    _body(doc, f"Data: {ctx.data_informe or '.....'}")
    doc.add_paragraph()
    doc.add_paragraph()
    _body(doc, f"Asdo.: {ctx.tecnico_nome or '.....'}")


def _anexos(doc):
    doc.add_paragraph()
    for a in ANEXOS:
        p = doc.add_paragraph()
        r = p.add_run(a)
        r.bold = True
        r.font.size = Pt(10)


# ---------------------------------------------------------------------------
# Entrada principal
# ---------------------------------------------------------------------------

def _force_field_update_on_open(doc):
    """Añade <w:updateFields w:val="true"/> a settings.xml para que Word
    recalcule los campos PAGE/NUMPAGES automáticamente al abrir el
    documento (si no, algunos visores muestran el valor sin calcular o el
    último valor en caché)."""
    settings = doc.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)


def generate_report(ctx: ReportContext, df: pd.DataFrame) -> io.BytesIO:
    """Genera el informe .docx y lo devuelve como BytesIO listo para descargar."""
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(4.0)
    section.bottom_margin = Cm(2)
    section.header_distance = Cm(0.8)
    _set_font(doc)
    _build_header(section, ctx)
    _force_field_update_on_open(doc)

    _section_1(doc, ctx)
    _section_2(doc, ctx)
    _section_3(doc, ctx)
    _section_4(doc)
    _section_5(doc)
    any_exceeds = False
    exceeded_rooms: list[str] = []
    if not df.empty:
        any_exceeds = _section_7(doc, df, ctx)
        if any_exceeds:
            mask = df["Resultado Bq/m3"] > NIVEL_REFERENCIA_BQ_M3
            for _, row in df[mask].iterrows():
                sala = _row_value(row, "Sala", "Código de la sala", "Código Sala", default="")
                if sala:
                    exceeded_rooms.append(str(sala))
    else:
        _heading(doc, "7", "RESULTADOS DAS MEDICIÓNS REALIZADAS")
        _body(doc, "Non hai datos de mediciones cargados para este centro de traballo.")

    _section_8(doc, ctx, any_exceeds, exceeded_rooms)
    _section_9(doc)
    _section_10(doc, ctx)
    _anexos(doc)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
