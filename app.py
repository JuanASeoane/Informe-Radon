import streamlit as st
import pandas as pd
import io
from datetime import datetime
import base64
from utils.pdf_generator import PDFGenerator
from utils.data_processor import DataProcessor

# Configuración de la página
st.set_page_config(
    page_title="Generador de Informes PDF",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("📊 Generador de Informes PDF desde Excel")
st.markdown("---")

# Barra lateral
with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown("---")
    
    # Opciones de configuración
    include_stats = st.checkbox("Incluir estadísticas descriptivas", value=True)
    include_charts = st.checkbox("Incluir gráficos", value=True)
    include_summary = st.checkbox("Incluir resumen ejecutivo", value=True)
    
    st.markdown("---")
    st.markdown("### 📋 Instrucciones")
    st.markdown("""
    1. Sube un archivo Excel (.xlsx, .xls)
    2. Selecciona las columnas a analizar
    3. Personaliza el informe
    4. Genera y descarga el PDF
    """)
    
    st.markdown("---")
    st.markdown("### 📊 Ejemplo de datos")
    st.markdown("""
    Puedes usar este formato de ejemplo:
    - **Ventas**: Datos numéricos
    - **Producto**: Datos categóricos
    - **Fecha**: Datos de fecha
    - **Cantidad**: Datos numéricos
    - **Precio**: Datos numéricos
    """)

# Área principal
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📤 Carga de Archivo")
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel",
        type=['xlsx', 'xls'],
        help="Formatos soportados: .xlsx, .xls"
    )

    if uploaded_file is not None:
        try:
            # Procesar el archivo
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ Archivo cargado exitosamente! {len(df)} filas, {len(df.columns)} columnas")
            
            # Mostrar vista previa
            st.subheader("👁️ Vista Previa de los Datos")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Información del dataset
            with st.expander("📊 Información del Dataset"):
                col_info1, col_info2, col_info3 = st.columns(3)
                with col_info1:
                    st.metric("Total Filas", len(df))
                with col_info2:
                    st.metric("Total Columnas", len(df.columns))
                with col_info3:
                    st.metric("Valores Nulos", df.isnull().sum().sum())
                
                # Tipos de datos
                st.write("**Tipos de datos:**")
                dtype_df = pd.DataFrame(df.dtypes.reset_index())
                dtype_df.columns = ['Columna', 'Tipo']
                st.dataframe(dtype_df, use_container_width=True)
                
                # Estadísticas rápidas
                st.write("**Estadísticas rápidas (columnas numéricas):**")
                numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
                if len(numeric_cols) > 0:
                    st.dataframe(df[numeric_cols].describe(), use_container_width=True)
                else:
                    st.info("No hay columnas numéricas para mostrar estadísticas")
            
        except Exception as e:
            st.error(f"❌ Error al cargar el archivo: {str(e)}")

with col2:
    if uploaded_file is not None:
        st.subheader("🔧 Configuración del Informe")
        
        # Selección de columnas
        all_columns = df.columns.tolist()
        selected_columns = st.multiselect(
            "Selecciona las columnas para el informe",
            options=all_columns,
            default=all_columns[:min(5, len(all_columns))]
        )
        
        # Título del informe
        report_title = st.text_input(
            "Título del Informe",
            value=f"Informe de Análisis - {datetime.now().strftime('%Y-%m-%d')}"
        )
        
        # Información adicional
        author_name = st.text_input("Autor", value="Usuario")
        additional_notes = st.text_area("Notas adicionales", placeholder="Agrega comentarios o notas...")
        
        # Botón para generar el PDF
        if st.button("📄 Generar Informe PDF", type="primary", use_container_width=True):
            if selected_columns:
                with st.spinner("Generando informe PDF..."):
                    try:
                        # Procesar datos
                        processor = DataProcessor(df)
                        processed_data = processor.process_data(selected_columns, include_stats)
                        
                        # Generar PDF
                        pdf_gen = PDFGenerator()
                        pdf_buffer = pdf_gen.generate_report(
                            processed_data,
                            title=report_title,
                            author=author_name,
                            notes=additional_notes,
                            include_charts=include_charts,
                            include_summary=include_summary
                        )
                        
                        # Crear botón de descarga
                        st.success("✅ Informe generado exitosamente!")
                        
                        # Convertir a base64 para descarga
                        b64 = base64.b64encode(pdf_buffer.getvalue()).decode()
                        href = f'''
                        <a href="data:application/pdf;base64,{b64}" 
                           download="{report_title.replace(' ', '_')}.pdf" 
                           style="text-decoration: none;">
                            <button style="
                                background-color: #4CAF50;
                                color: white;
                                padding: 12px 24px;
                                border: none;
                                border-radius: 8px;
                                cursor: pointer;
                                font-size: 16px;
                                width: 100%;
                                font-weight: bold;
                                transition: all 0.3s ease;
                            ">
                                ⬇️ Descargar Informe PDF
                            </button>
                        </a>
                        '''
                        st.markdown(href, unsafe_allow_html=True)
                        
                        # Mostrar vista previa del PDF
                        with st.expander("📄 Vista previa del informe"):
                            st.info("El PDF contiene: tablas de datos, estadísticas, gráficos y notas")
                            st.markdown(f"**Título:** {report_title}")
                            st.markdown(f"**Autor:** {author_name}")
                            st.markdown(f"**Fecha:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
                            st.markdown(f"**Columnas analizadas:** {', '.join(selected_columns)}")
                            st.markdown(f"**Registros:** {len(df)}")
                        
                    except Exception as e:
                        st.error(f"❌ Error al generar el informe: {str(e)}")
            else:
                st.warning("⚠️ Por favor, selecciona al menos una columna")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666; padding: 20px;'>
        <p>Desarrollado con ❤️ usando Streamlit | Generador de Informes PDF v1.0</p>
        <p style='font-size: 12px;'>📧 Contacto: soporte@ejemplo.com</p>
    </div>
    """,
    unsafe_allow_html=True
)