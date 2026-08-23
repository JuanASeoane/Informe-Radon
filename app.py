import datetime as dt
import io
import zipfile

import pandas as pd
import streamlit as st

from utils.docx_generator import ReportContext, generar_conclusion_automatica, generate_report
from utils.excel_parser import (
    ExcelFormatError,
    areas_muestreadas,
    categorias_resumen,
    categorias_turnos_bullets,
    filter_group,
    group_options,
    load_workbook,
    merge_resultados,
    postos_traballo_bullets,
    salas_medidas,
)

st.set_page_config(
    page_title="Informe de mediciones de radón",
    page_icon="☢️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Utilidades de validación visual (borde rojo en campos vacíos / sin archivo)
# ---------------------------------------------------------------------------
_campos_vacios: list[str] = []


def marca(key: str):
    """Marcador invisible que se coloca justo antes de un campo para poder
    localizarlo luego por CSS (ver `pintar_campos_vacios`)."""
    st.markdown(f'<div class="reqmark reqmark-{key}"></div>', unsafe_allow_html=True)


def registrar_si_vacio(key: str, valor) -> None:
    vacio = valor is None or (isinstance(valor, str) and not valor.strip())
    if vacio:
        _campos_vacios.append(key)


def pintar_campos_vacios(keys=None):
    """Inyecta el CSS que dibuja el borde rojo para cada campo vacío
    registrado hasta el momento (o para la lista `keys` indicada)."""
    claves = keys if keys is not None else _campos_vacios
    if not claves:
        return
    reglas = []
    for k in claves:
        reglas.append(
            f'.element-container:has(.reqmark-{k}) + div.element-container input,\n'
            f'.element-container:has(.reqmark-{k}) + div.element-container textarea,\n'
            f'.element-container:has(.reqmark-{k}) + div.element-container '
            f'div[data-testid="stFileUploaderDropzone"] {{\n'
            f'    border: 2px solid #d9534f !important;\n'
            f'    border-radius: 6px;\n'
            f'}}'
        )
    st.markdown(f"<style>{chr(10).join(reglas)}</style>", unsafe_allow_html=True)


st.title("☢️ Generador de informes de mediciones de radón")
st.caption(
    "A partir del Excel de detectores, genera el informe de resultados según el modelo "
    "oficial (UPRL / Servizo Galego de Saúde) y la documentación de referencia del "
    "Consello de Seguridade Nuclear (CSN). Los campos con borde rojo están vacíos o sin "
    "archivo subido."
)

# ---------------------------------------------------------------------------
# 1. Carga del Excel principal (Detectores / Planos / Categorías profesionales)
# ---------------------------------------------------------------------------
st.header("1. Cargar Excel de detectores")
marca("excel_principal")
uploaded = st.file_uploader(
    "Excel con las hojas 'Detectores', 'Planos' y 'Categorías profesionales'",
    type=["xlsx", "xlsm"],
)

if not uploaded:
    pintar_campos_vacios(["excel_principal"])
    st.info("Sube el Excel con los datos de los detectores para continuar.")
    st.stop()

try:
    wb = load_workbook(uploaded)
except ExcelFormatError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"No se ha podido leer el Excel: {e}")
    st.stop()

df = wb["detectores"]
det_meta = wb["detectores_meta"]
planos_meta = wb["planos_meta"]
categorias_df = wb["categorias"]

st.success(f"Excel cargado correctamente: {len(df)} fila(s) de detectores.")
with st.expander("Ver datos cargados"):
    st.dataframe(df, use_container_width=True)
    if planos_meta:
        st.caption(f"Datos de la hoja 'Planos': {planos_meta}")
    if det_meta:
        st.caption(f"Metadatos de la hoja 'Detectores': {det_meta}")

# ---------------------------------------------------------------------------
# 2. Selección del centro de trabajo
# ---------------------------------------------------------------------------
st.header("2. Selecciona el centro de trabajo del informe")

options = group_options(df, "Centro")
if not options:
    st.error("No se han encontrado valores en la columna 'Centro'.")
    st.stop()

selected_value = st.selectbox("Centro de trabajo", options=options)
df_center = filter_group(df, selected_value, "Centro")
st.caption(f"{len(df_center)} detector(es) para «{selected_value}».")

areas = areas_muestreadas(df_center)
salas = salas_medidas(df_center)
total_personas, categorias_texto = categorias_resumen(categorias_df)
postos_bullets_default = postos_traballo_bullets(df_center)
total_traballadores, categorias_bullets_default = categorias_turnos_bullets(df_center, categorias_df)

sin_turno = [b for b in categorias_bullets_default if "(turno de" not in b]
if sin_turno:
    st.warning(
        "⚠️ No se ha podido emparejar el turno para: "
        + ", ".join(sin_turno)
        + ". Revisa que la categoría profesional coincida con el texto de "
        "'Profesionales en la sala' en el Excel, o edítalo manualmente abajo."
    )

# ---------------------------------------------------------------------------
# 3. Datos del centro (autorrellenados desde el Excel, editables)
# ---------------------------------------------------------------------------
st.header("3. Datos del centro de trabajo")
st.caption(
    "Los campos se autorrellenan a partir del Excel (hojas 'Planos' y metadatos de "
    "'Detectores'). Revísalos y complétalos antes de generar el informe."
)

col1, col2 = st.columns(2)
with col1:
    marca("xerencia")
    xerencia = st.text_input("Xerencia (Empresa)", value=str(planos_meta.get("Empresa", "")))
    registrar_si_vacio("xerencia", xerencia)

    marca("cif")
    cif = st.text_input("CIF", value=str(planos_meta.get("CIF", "")))
    registrar_si_vacio("cif", cif)

    centro_nombre = st.text_input("Nombre completo del centro", value=selected_value)

    marca("servizo_unidade")
    servizo_unidade = st.text_input(
        "Servizo / Unidade mostrexada (Área)", value=", ".join(areas)
    )
    registrar_si_vacio("servizo_unidade", servizo_unidade)

    marca("enderezo")
    enderezo = st.text_input("Dirección (Enderezo)", value=str(det_meta.get("Dirección", "")))
    registrar_si_vacio("enderezo", enderezo)
with col2:
    marca("superficie_construida")
    superficie_construida = st.text_input("Superficie construida (m²)", value="")
    registrar_si_vacio("superficie_construida", superficie_construida)

    marca("superficie_util")
    superficie_util = st.text_input("Superficie útil (m²)", value="")
    registrar_si_vacio("superficie_util", superficie_util)

    marca("num_plantas")
    num_plantas = st.text_input("N.º de plantas", value="")
    registrar_si_vacio("num_plantas", num_plantas)

    fecha_excel = det_meta.get("Fecha", "")
    try:
        fecha_default = (
            dt.datetime.strptime(fecha_excel, "%d/%m/%Y").date()
            if fecha_excel
            else dt.date.today()
        )
    except ValueError:
        fecha_default = dt.date.today()
    data_informe = st.date_input("Fecha del informe", value=fecha_default)

st.subheader("Información sobre los/as trabajadores/as")
st.caption(
    "Se generan automáticamente a partir del Excel: los puestos por sala (columna "
    "'Profesionales en la sala') y las quendas por categoría (hoja 'Categorías "
    "profesionales' cruzada con 'Turno de trabajo'). Puedes editar cada línea."
)
col3, col4 = st.columns(2)
with col3:
    st.markdown("**Puestos de trabajo por sala** (una línea por puesto)")
    postos_text = st.text_area(
        "postos_text",
        value="\n".join(postos_bullets_default),
        label_visibility="collapsed",
        height=120,
    )
    postos_bullets = [line.strip() for line in postos_text.splitlines() if line.strip()]

    marca("num_traballadores_total")
    num_traballadores_total = st.text_input(
        "N.º total de trabajadores adscritos", value=str(total_traballadores or "")
    )
    registrar_si_vacio("num_traballadores_total", num_traballadores_total)
with col4:
    st.markdown("**Categorías profesionales y quendas** (una línea por categoría)")
    categorias_text = st.text_area(
        "categorias_text",
        value="\n".join(categorias_bullets_default),
        label_visibility="collapsed",
        height=120,
    )
    categorias_bullets = [line.strip() for line in categorias_text.splitlines() if line.strip()]

    marca("data_informacion_traballadores")
    data_informacion_traballadores = st.text_input(
        "Fecha de comunicación a los trabajadores", value=""
    )
    registrar_si_vacio("data_informacion_traballadores", data_informacion_traballadores)

    marca("medio_informacion_traballadores")
    medio_informacion_traballadores = st.text_input(
        "Medio de comunicación", value="correo electrónico"
    )
    registrar_si_vacio("medio_informacion_traballadores", medio_informacion_traballadores)

notas_adicionais = st.text_area(
    "Notas adicionales sobre ocupación de espacios (opcional)",
    placeholder="Texto libre que se añadirá como párrafo adicional en el punto 3, si lo rellenas.",
)

st.subheader("Firma")
marca("tecnico_nome")
tecnico_nome = st.text_input(
    "Nombre del/de la técnico/a que firma el informe", value=str(det_meta.get("Técnico", ""))
)
registrar_si_vacio("tecnico_nome", tecnico_nome)

st.subheader("Logotipo de cabecera (opcional)")
st.caption("Un único logotipo, a todo el ancho de la cabecera (p.ej. logo combinado UPRL/SERGAS/Área Sanitaria).")
logo_file = st.file_uploader("Logo", type=["png", "jpg", "jpeg"], key="logo_unico")
logo_bytes = logo_file.read() if logo_file else None

# ---------------------------------------------------------------------------
# 4. Tabla de resultados (punto 7): semicubierta, editable o por Excel externo
# ---------------------------------------------------------------------------
st.header("4. Resultados de las mediciones (punto 7 del informe)")
st.caption(
    "La tabla se rellena con los datos ya disponibles en el Excel de detectores "
    "(código de zona, código de detector, fechas, puestos de trabajo). Introduce "
    "manualmente el resultado y la incertidumbre de cada detector, o sube un Excel "
    "de resultados del laboratorio para completarlos automáticamente."
)

resultados_file = st.file_uploader(
    "Excel de resultados del laboratorio (opcional) — debe tener una columna 'Código' "
    "que coincida con el código de detector, y columnas de resultado/incertidumbre",
    type=["xlsx", "xlsm"],
    key="resultados_excel",
)

df_working = df_center.copy()
if resultados_file:
    try:
        resultados_df = pd.read_excel(resultados_file)
        df_working = merge_resultados(df_working, resultados_df)
        st.success("Resultados del laboratorio incorporados a la tabla.")
    except ExcelFormatError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"No se ha podido leer el Excel de resultados: {e}")

incerteza_defecto = st.text_input(
    "Incertidumbre expandida y K por defecto (se usará si dejas la celda vacía)",
    value="",
    placeholder="p.ej. ±15% (k=2)",
)

editable_cols = [
    "Código de la sala",
    "Sala",
    "Código",
    "Fecha de colocación fmt",
    "Fecha de retirada real fmt",
    "Resultado Bq/m3",
    "Incerteza expandida e K",
    "Profesionales en la sala",
]
editable_cols = [c for c in editable_cols if c in df_working.columns]

edited = st.data_editor(
    df_working[editable_cols],
    use_container_width=True,
    num_rows="fixed",
    key="results_editor",
    column_config={
        "Resultado Bq/m3": st.column_config.NumberColumn(
            "Resultado (Bq/m³)", help="Concentración de radón medida"
        ),
        "Incerteza expandida e K": st.column_config.TextColumn(
            "Incerteza expandida e K", help="p.ej. ±15% (k=2)"
        ),
    },
)

df_final = df_working.copy()
for c in editable_cols:
    df_final[c] = edited[c]

pendientes = int(df_final["Resultado Bq/m3"].isna().sum()) if "Resultado Bq/m3" in df_final else 0
if pendientes:
    st.info(f"ℹ️ Quedan {pendientes} detector(es) sin resultado. Puedes completarlos antes de generar el informe.")

exceeded = (
    int((df_final["Resultado Bq/m3"] > 300).sum()) if "Resultado Bq/m3" in df_final else 0
)
if exceeded:
    st.warning(f"⚠️ {exceeded} medición(es) superan el nivel de referencia de 300 Bq/m³.")
elif not pendientes:
    st.success("Ninguna medición supera el nivel de referencia de 300 Bq/m³.")

conclusion_default = generar_conclusion_automatica(df_final)
conclusion_manual = st.text_area(
    "Texto de conclusiones (puedes modificarlo directamente)",
    value=conclusion_default,
    height=100,
)

# ---------------------------------------------------------------------------
# 5. Anexos
# ---------------------------------------------------------------------------
st.header("5. Anexos")
st.caption(
    "Sube aquí los documentos de cada anexo (PDF, Word o imagen). Si subes alguno, al generar "
    "el informe se creará también un ZIP con el informe Word y los anexos juntos. Los anexos "
    "III y IV recuerdan el último archivo subido: no hace falta volver a subirlos si no cambian."
)

ANEXOS_INFO = [
    ("anexo1", "ANEXO I: FORMULARIOS TOMA DE DATOS"),
    ("anexo2", "ANEXO II: ESQUEMA GRÁFICO DO EDIFICIO E PLANOS DE CADA PLANTA"),
    ("anexo3", "ANEXO III: INFORME DE ENSAIO DO LABORATORIO ACREDITADO"),
    ("anexo4", "ANEXO IV: CERTIFICADO ENAC DO LABORATORIO ACREDITADO"),
]
# Los anexos III y IV son "persistentes": si no se sube uno nuevo en esta
# sesión, se reutiliza el último que se subió (útil porque suelen ser
# siempre el mismo informe de laboratorio / certificado ENAC).
ANEXOS_PERSISTENTES = {"anexo3", "anexo4"}

anexos_datos = {}
for key, label in ANEXOS_INFO:
    marca(key)
    archivo = st.file_uploader(
        label, type=["pdf", "doc", "docx", "jpg", "jpeg", "png"], key=key
    )
    if archivo is not None:
        anexos_datos[key] = (archivo.name, archivo.getvalue())
        if key in ANEXOS_PERSISTENTES:
            st.session_state[f"{key}_guardado"] = (archivo.name, archivo.getvalue())
    elif key in ANEXOS_PERSISTENTES and st.session_state.get(f"{key}_guardado"):
        nombre_guardado, _ = st.session_state[f"{key}_guardado"]
        anexos_datos[key] = st.session_state[f"{key}_guardado"]
        st.caption(f"📎 Se mantiene el último archivo subido para este anexo: **{nombre_guardado}**")
    else:
        anexos_datos[key] = None
    registrar_si_vacio(key, "x" if anexos_datos[key] else "")

pintar_campos_vacios()

# ---------------------------------------------------------------------------
# 6. Generar informe
# ---------------------------------------------------------------------------
st.header("6. Generar informe")

if st.button("📄 Generar informe Word", type="primary"):
    ctx = ReportContext(
        xerencia=xerencia,
        cif=cif,
        centro=centro_nombre,
        servizo_unidade=servizo_unidade,
        enderezo=enderezo,
        superficie_construida=superficie_construida,
        superficie_util=superficie_util,
        num_plantas=num_plantas,
        postos_bullets=postos_bullets,
        num_traballadores_total=num_traballadores_total,
        categorias_bullets=categorias_bullets,
        notas_adicionais=notas_adicionais,
        data_informacion_traballadores=data_informacion_traballadores,
        medio_informacion_traballadores=medio_informacion_traballadores,
        data_informe=data_informe.strftime("%d/%m/%Y"),
        tecnico_nome=tecnico_nome,
        incertezas_por_defecto=incerteza_defecto,
        conclusion_manual=conclusion_manual,
        logo=logo_bytes,
    )
    buffer = generate_report(ctx, df_final)
    report_name = f"Informe_radon_{selected_value}_{data_informe.isoformat()}.docx"
    st.session_state["report_buffer"] = buffer.getvalue()
    st.session_state["report_name"] = report_name
    # nuevo informe generado: vuelve a mostrar los botones de descarga
    st.session_state["report_descargado"] = False
    st.session_state["zip_descargado"] = False

    anexos_subidos = {k: v for k, v in anexos_datos.items() if v is not None}
    if anexos_subidos:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(report_name, buffer.getvalue())
            numeros = {"anexo1": "I", "anexo2": "II", "anexo3": "III", "anexo4": "IV"}
            for key, (nombre_original, contenido) in anexos_subidos.items():
                ext = nombre_original.rsplit(".", 1)[-1] if "." in nombre_original else "pdf"
                zf.writestr(f"ANEXO_{numeros[key]}.{ext}", contenido)
        st.session_state["zip_buffer"] = zip_buffer.getvalue()
        st.session_state["zip_name"] = f"Informe_radon_{selected_value}_con_anexos.zip"
    else:
        st.session_state.pop("zip_buffer", None)

if "report_buffer" in st.session_state and not st.session_state.get("report_descargado", False):
    descargado = st.download_button(
        "⬇️ Descargar informe (.docx)",
        data=st.session_state["report_buffer"],
        file_name=st.session_state["report_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    if descargado:
        st.session_state["report_descargado"] = True
        st.rerun()
elif st.session_state.get("report_descargado", False):
    st.success(f"✅ Informe «{st.session_state.get('report_name', '')}» descargado.")

if "zip_buffer" in st.session_state and not st.session_state.get("zip_descargado", False):
    zip_descargado = st.download_button(
        "⬇️ Descargar informe + anexos (.zip)",
        data=st.session_state["zip_buffer"],
        file_name=st.session_state["zip_name"],
        mime="application/zip",
    )
    if zip_descargado:
        st.session_state["zip_descargado"] = True
        st.rerun()
elif st.session_state.get("zip_descargado", False):
    st.success(f"✅ Paquete «{st.session_state.get('zip_name', '')}» descargado.")
