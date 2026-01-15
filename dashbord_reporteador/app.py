import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import os
import io

# Configuración de la página
st.set_page_config(page_title="Dashboard JUD Información", layout="wide")

# --- CLASE PARA EL FORMATO DEL PDF ---
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'UNIDAD DEPARTAMENTAL DE INFORMACIÓN', 0, 1, 'C')
        self.set_font('Arial', 'I', 10)
        self.cell(0, 10, 'Reporte de Estatus de Proyecto', 0, 1, 'C')
        self.line(10, 30, 200, 30)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def generar_pdf(registro):
    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    def texto(t):
        return str(t).encode('latin-1', 'replace').decode('latin-1')

    # Estructura del Informe
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, texto(f"ID del Proceso: {registro['ID_Proceso']}"), ln=1)
    
    pdf.set_font("Arial", size=12)
    pdf.ln(5)
    
    # Datos Principales
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(50, 10, "Nombre del Proyecto:", 1, 0, 'L', 1)
    pdf.cell(0, 10, texto(registro['Nombre_Proyecto']), 1, 1, 'L')
    
    pdf.cell(50, 10, "Área Solicitante:", 1, 0, 'L', 1)
    pdf.cell(0, 10, texto(registro['Área_Solicitante']), 1, 1, 'L')

    pdf.cell(50, 10, "Responsable:", 1, 0, 'L', 1)
    pdf.cell(0, 10, texto(registro['Responsable']), 1, 1, 'L')
    
    pdf.ln(10)
    
    # Estatus
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "Estatus y Tiempos", ln=1)
    pdf.set_font("Arial", size=12)
    
    pdf.cell(50, 10, "Estatus Actual:", 1, 0, 'L', 1)
    pdf.cell(0, 10, texto(registro['Estatus']), 1, 1, 'L')
    
    pdf.cell(50, 10, "Prioridad:", 1, 0, 'L', 1)
    pdf.cell(0, 10, texto(registro['Prioridad']), 1, 1, 'L')
    
    pdf.ln(5)
    
    # Fechas
    fecha_inicio = registro['Fecha_Inicio'].strftime('%d/%m/%Y') if pd.notnull(registro['Fecha_Inicio']) else "N/A"
    fecha_comp = registro['Fecha_Compromiso'].strftime('%d/%m/%Y') if pd.notnull(registro['Fecha_Compromiso']) else "N/A"
    fecha_fin = registro['Fecha_Finalizacion'].strftime('%d/%m/%Y') if pd.notnull(registro['Fecha_Finalizacion']) else "Pendiente"
    
    pdf.cell(63, 10, f"Inicio: {fecha_inicio}", 1, 0, 'C')
    pdf.cell(63, 10, f"Compromiso: {fecha_comp}", 1, 0, 'C')
    pdf.cell(63, 10, f"Finalización: {fecha_fin}", 1, 1, 'C')

    pdf.ln(20)
    pdf.cell(0, 10, "_"*60, ln=1, align='C')
    pdf.cell(0, 5, "Firma del Responsable", ln=1, align='C')

    return pdf.output(dest='S').encode('latin-1')

# --- CARGA DE DATOS ---
@st.cache_data
def cargar_datos():
    try:
        # Lee el archivo asegurando que ID sea string para evitar problemas
        BASE_DIR= os.path.dirname(os.path.abspath(__file__))
        ruta_excel=os.path.join(BASE_DIR,"seguimiento_procesos.xlsx")
        df = pd.read_excel(ruta_excel, dtype={'ID_Proceso': str})
        df.columns = df.columns.str.strip()
        df['Fecha_Inicio'] = pd.to_datetime(df['Fecha_Inicio'])
        df['Fecha_Compromiso'] = pd.to_datetime(df['Fecha_Compromiso'])
        return df
    except Exception as e:
        return None

df = cargar_datos()

if df is not None:
    
    # --- [NUEVO] VALIDACIÓN DE DUPLICADOS ---
    # Esto verifica si hay IDs repetidos antes de mostrar nada más
    if df['ID_Proceso'].duplicated().any():
        ids_duplicados = df[df['ID_Proceso'].duplicated()]['ID_Proceso'].unique().tolist()
        st.error(f"⚠️ ¡ALERTA DE ERROR! Se encontraron IDs duplicados en el archivo Excel: {ids_duplicados}")
        st.warning("Por favor, corrige el archivo Excel y recarga la página. No se pueden generar reportes confiables con IDs repetidos.")
        st.stop() # Detiene la ejecución para obligar a corregir el error
    
    # Si no hay duplicados, continúa normalmente:
    st.title("📊 Sistema de Seguimiento - JUD Información")
    st.markdown("---")

    # --- BARRA LATERAL ---
    st.sidebar.header("🔍 Filtros")
    
    area_options = df["Área_Solicitante"].unique() if "Área_Solicitante" in df.columns else []
    area = st.sidebar.multiselect("Área Solicitante:", options=area_options, default=area_options)
    estatus = st.sidebar.multiselect("Estatus:", options=df["Estatus"].unique(), default=df["Estatus"].unique())
    prio = st.sidebar.multiselect("Prioridad:", options=df["Prioridad"].unique(), default=df["Prioridad"].unique()) #Filtro de prioridad, agregado 09/01/2026

    # Filtrar DataFrame
    df_selection = df.query("Área_Solicitante == @area & Estatus == @estatus & Prioridad == @prio")

    # --- GENERADOR DE INFORMES ---
    st.sidebar.markdown("---")
    st.sidebar.header("📄 Generar Informe Individual")
    
    lista_proyectos = df['ID_Proceso'].tolist() #cambio de filtro de pdf _selection
    
    if lista_proyectos:
        proyecto_seleccionado = st.sidebar.selectbox("Selecciona ID para informe:", lista_proyectos)
        
        if st.sidebar.button("Generar PDF"):
            # Al no haber duplicados, .iloc[0] es seguro
            registro = df[df['ID_Proceso'] == proyecto_seleccionado].iloc[0]# cambio de filtro de pdf  .iloc[0]
            pdf_bytes = generar_pdf(registro)
            
            st.sidebar.download_button(
                label="⬇️ Descargar Informe PDF",
                data=pdf_bytes,
                file_name=f"Informe_{proyecto_seleccionado}.pdf",
                mime="application/pdf"
            )
    else:
        st.sidebar.info("No hay proyectos visibles con los filtros actuales.")

    # --- METRICAS Y GRAFICAS ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Procesos", len(df_selection))
    col2.metric("Finalizados ✅", len(df_selection[df_selection["Estatus"] == "Finalizado"]))
    col3.metric("En Proceso ⏳", len(df_selection[df_selection["Estatus"] == "Proceso"]))
    col4.metric("Por Iniciar 📁", len(df_selection[df_selection["Estatus"] == "Inicio"]))

    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        if not df_selection.empty:
            fig_area = px.bar(df_selection, x="Área_Solicitante", color="Estatus", barmode="group",
                              color_discrete_map={"Inicio": "#FECB52", "Proceso": "#636EFA", "Finalizado": "#00CC96"})
            st.plotly_chart(fig_area, use_container_width=True)
            
    with col_chart2:
        if not df_selection.empty:
            fig_prioridad = px.pie(df_selection, names="Prioridad", hole=0.4)
            st.plotly_chart(fig_prioridad, use_container_width=True)

    st.subheader("📋 Detalle de la Información")
    #st.dataframe(df_selection, use_container_width=True)   #Primera opcion de cambio de filtrado
    st.dataframe(df)

else:
    st.error("Por favor crea el archivo 'seguimiento_procesos.xlsx' en la carpeta.")