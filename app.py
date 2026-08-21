import datetime as dt

import streamlit as st

from utils.docx_generator import ReportContext, generate_report
from utils.excel_parser import ExcelFormatError, filter_group, group_options, load_excel

st.set_page_config(
    page_title="Informe de mediciones de radón",
    page_icon="☢️",
    layout="wide",
)

st.title("☢️ Generador de informes de mediciones de radón")
st.caption(
    "A partir del Excel de detectores, genera el informe de resultados según el modelo "
    "oficial (UPRL / Servizo Galego de Saúde) y la documentación de referencia del "
    "Consello de Seguridade Nuclear (CSN)."
)

# ---------------------------------------------------------------------------
# 1. Carga del Excel
# ---------------------------------------------------------------------------
st.header("1. Cargar Excel de detectores")
uploaded = st.file_uploader("Excel con la hoja 'Detectores'", type=["xlsx", "xlsm"])

if not uploaded:
    st.info("Sube el Excel con los datos de los detectores para continuar.")
    st.stop()

try:
    df = load_excel(uploaded)
except ExcelFormatError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"No se ha podido leer el Excel: {e}")
    st.stop()

st.success(f"Excel cargado correctamente: {len(df)} fila(s) de detectores.")
with st.expander("Ver datos cargados"):
    st.dataframe(df, use_container_width=True)

# ---------------------------------------------------------------------------
# 2. Selección del centro de trabajo
# ---------------------------------------------------------------------------
st.header("2. Selecciona el centro de trabajo del informe")

group_col = st.selectbox(
    "Columna que identifica el centro de trabajo",
    options=["Área", "Centro"],
    index=0,
    help=(
        "En el Excel de ejemplo, 'Área' recoge el hospital/centro de salud "
        "(p.ej. HUAC) y 'Centro' la entidad que realiza la medición. Cambia la "
        "selección si tu Excel usa las columnas al revés."
    ),
)

options = group_options(df, group_col)
if not options:
    st.error(f"No se han encontrado valores en la columna '{group_col}'.")
    st.stop()

selected_value = st.selectbox("Centro de trabajo", options=options)
df_center = filter_group(df, group_col, selected_value)
st.caption(f"{len(df_center)} detector(es) para «{selected_value}».")

# ---------------------------------------------------------------------------
# 3. Datos del centro (no vienen en el Excel)
# ---------------------------------------------------------------------------
st.header("3. Datos del centro de trabajo")

col1, col2 = st.columns(2)
with col1:
    xerencia = st.text_input("Xerencia / Área Sanitaria", value="")
    cif = st.text_input("CIF", value="")
    centro_nombre = st.text_input("Nombre completo del centro", value=selected_value)
    enderezo = st.text_input("Dirección (Enderezo)", value="")
with col2:
    superficie_construida = st.text_input("Superficie construida (m²)", value="")
    superficie_util = st.text_input("Superficie útil (m²)", value="")
    num_plantas = st.text_input("N.º de plantas", value="")
    data_informe = st.date_input("Fecha del informe", value=dt.date.today())

st.subheader("Información sobre los/as trabajadores/as")
col3, col4 = st.columns(2)
with col3:
    postos_traballo_desc = st.text_area(
        "Descripción de los puestos de trabajo",
        placeholder="p.ej. consulta de enfermería, despacho de admisión, sala de espera...",
    )
    num_traballadores = st.text_input("N.º de trabajadores adscritos", value="")
    horario_quendas = st.text_input("Horario de trabajo / turnos", value="")
with col4:
    ocupacion_espazos = st.text_area(
        "Ocupación de los espacios de trabajo",
        placeholder="p.ej. consulta de enfermería: 2 personas, jornada completa...",
    )
    data_informacion_traballadores = st.text_input(
        "Fecha de comunicación a los trabajadores", value=""
    )
    medio_informacion_traballadores = st.text_input(
        "Medio de comunicación", value="correo electrónico"
    )

st.subheader("Firma")
tecnico_nome = st.text_input("Nombre del/de la técnico/a que firma el informe", value="")

st.subheader("Logotipos de cabecera (opcional)")
st.caption("Sube hasta 3 imágenes para la cabecera institucional (p.ej. UPRL, SERGAS, Área Sanitaria).")
logo_cols = st.columns(3)
logos = []
for i, c in enumerate(logo_cols):
    with c:
        f = st.file_uploader(f"Logo {i + 1}", type=["png", "jpg", "jpeg"], key=f"logo_{i}")
        logos.append(f.read() if f else None)

# ---------------------------------------------------------------------------
# 4. Tabla de resultados (editable: permite añadir incertidumbre, revisar)
# ---------------------------------------------------------------------------
st.header("4. Revisar la tabla de resultados del informe")
st.caption(
    "Puedes editar los valores antes de generar el informe (p.ej. añadir la incertidumbre "
    "expandida y K de cada detector, que no viene en el Excel de origen)."
)

if "Incerteza expandida e K" not in df_center.columns:
    df_center["Incerteza expandida e K"] = ""

incerteza_defecto = st.text_input(
    "Incertidumbre expandida y K por defecto (se usará si dejas la celda vacía)",
    value="",
    placeholder="p.ej. ±15% (k=2)",
)

editable_cols = [
    "Código Sala",
    "Sala",
    "Código",
    "Fecha de colocación fmt",
    "Fecha de retirada real fmt",
    "Resultado Bq/m3",
    "Incerteza expandida e K",
    "Puestos en la sala",
]
editable_cols = [c for c in editable_cols if c in df_center.columns]

edited = st.data_editor(
    df_center[editable_cols],
    use_container_width=True,
    num_rows="fixed",
    key="results_editor",
)

# vuelca las ediciones sobre el dataframe completo (mantiene el resto de columnas)
df_final = df_center.copy()
for c in editable_cols:
    df_final[c] = edited[c]

exceeded = int((df_final["Resultado Bq/m3"] > 300).sum()) if "Resultado Bq/m3" in df_final else 0
if exceeded:
    st.warning(f"⚠️ {exceeded} medición(es) superan el nivel de referencia de 300 Bq/m³.")
else:
    st.success("Ninguna medición supera el nivel de referencia de 300 Bq/m³.")

conclusion_manual = st.text_area(
    "Texto de conclusiones (opcional, sobrescribe el texto automático)",
    placeholder="Déjalo vacío para que se genere automáticamente según los resultados.",
)

# ---------------------------------------------------------------------------
# 5. Generar informe
# ---------------------------------------------------------------------------
st.header("5. Generar informe")

if st.button("📄 Generar informe Word", type="primary"):
    ctx = ReportContext(
        xerencia=xerencia,
        cif=cif,
        centro=centro_nombre,
        enderezo=enderezo,
        superficie_construida=superficie_construida,
        superficie_util=superficie_util,
        num_plantas=num_plantas,
        postos_traballo_desc=postos_traballo_desc,
        num_traballadores=num_traballadores,
        horario_quendas=horario_quendas,
        ocupacion_espazos=ocupacion_espazos,
        data_informacion_traballadores=data_informacion_traballadores,
        medio_informacion_traballadores=medio_informacion_traballadores,
        data_informe=data_informe.strftime("%d/%m/%Y"),
        tecnico_nome=tecnico_nome,
        incertezas_por_defecto=incerteza_defecto,
        conclusion_manual=conclusion_manual,
        logos=logos,
    )
    buffer = generate_report(ctx, df_final)
    st.session_state["report_buffer"] = buffer.getvalue()
    st.session_state["report_name"] = f"Informe_radon_{selected_value}_{data_informe.isoformat()}.docx"

if "report_buffer" in st.session_state:
    st.download_button(
        "⬇️ Descargar informe (.docx)",
        data=st.session_state["report_buffer"],
        file_name=st.session_state["report_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
