import io
import os
import re
import folium
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from streamlit_folium import st_folium

# -----------------------------------------------------------------------------
# PALETA DE COLORES OFICIAL CONSORCIO INTEGRAS
# -----------------------------------------------------------------------------
COLOR_AGUAMARINA = '#17C3B2'  # Verde / Azul Agua Marina oficial
COLOR_ROSADO_AAP = '#D89FE3'  # Morado / Rosado Orquídea para Reporte AAP
COLOR_VERDE_ABIERTO = '#28A745'  # Verde para Casos Abiertos / En Proceso
COLOR_AMARILLO_MOSTAZA = '#E5B130'  # Amarillo Mostaza para Indicadores AAP

PALETA_INTEGRAS = [
    COLOR_AGUAMARINA,
    COLOR_ROSADO_AAP,
    COLOR_AMARILLO_MOSTAZA,
    '#08327D',  # Azul Marino
    '#0072CE',  # Azul Celeste
]

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y FUENTES PERSONALIZADAS (CSS)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title='Tablero Consorcio Integras | COOPI',
    layout='wide',
    initial_sidebar_state='expanded',
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;800&family=Quicksand:wght@600;700&display=swap');

    html, body, [class*="css"], .stMarkdown, p, div, span, label, input, button {
        font-family: 'Quicksand', sans-serif !important;
        font-weight: 700 !important;
    }

    h1, h2, h3, h4, h5, h6, .stSubheader {
        font-family: 'Now', 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
    }

    .titulo-principal {
        font-family: 'Now', 'Montserrat', sans-serif !important;
        color: #17C3B2 !important;
        margin-bottom: 5px !important;
        font-weight: 800 !important;
        font-size: 2.2rem !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# ENCABEZADO Y LOGO
# -----------------------------------------------------------------------------
col_header_title, col_header_logo = st.columns([3, 1])

with col_header_title:
    st.markdown(
        "<h1 class='titulo-principal'>Tablero de Monitoreo - Consorcio Íntegras</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '**Socio Prime / Líder:** COOPI | **Colíder:** HIAS | **Socios Implementadores:** FLM, PLAFAM, PALUZ'
    )

with col_header_logo:
    logo_path = 'integras.jpg' if os.path.exists('integras.jpg') else None
    if logo_path:
        st.image(logo_path, width='stretch')
    else:
        st.warning("⚠️ No se pudo cargar el logo 'integras.jpg'.")

st.markdown('---')

META_PARTICIPANTES_UNICOS = 46122

MAPA_MUNICIPIOS = {
    'Libertador': [10.5000, -66.9167],
    'Caroní': [8.2833, -62.7167],
    'Angostura': [7.5333, -63.8833],
    'Pedernales': [9.9667, -62.2500],
    'Tucupita': [9.0622, -62.0531],
    'Paz Castillo': [10.2167, -66.6667],
    'Sucre': [10.4833, -66.8167],
    'Urdaneta': [10.1500, -66.8833],
    'Arismendi': [10.7167, -62.5167],
    'Bermúdez': [10.6333, -63.2500],
    'Cruz Salmerón Acosta': [10.6167, -64.2000],
    'Vargas': [10.6000, -66.9333],
}

font_layout = dict(family='Quicksand', size=13)

# -----------------------------------------------------------------------------
# 2. CARGA INTELIGENTE DE DATOS DESDE EXCEL
# -----------------------------------------------------------------------------
st.sidebar.header('Carga de Datos (Excel)')
uploaded_file = st.sidebar.file_uploader(
    "Sube el archivo Excel de reporte consolidado", 
    type=['xlsx', 'xls']
)

@st.cache_data
def cargar_datos_excel(file_bytes_or_path):
    try:
        xls = pd.ExcelFile(file_bytes_or_path)
        sheet_target = 'group_beneficiario' if 'group_beneficiario' in xls.sheet_names else xls.sheet_names[0]
        
        # Búsqueda automática de la fila de encabezados reales
        df_test = pd.read_excel(file_bytes_or_path, sheet_name=sheet_target, header=None)
        header_row = 0
        for idx, row in df_test.iloc[:15].iterrows():
            row_str = " ".join([str(v) for v in row.values]).lower()
            if 'codigoid' in row_str or 'nombres' in row_str:
                header_row = idx
                break
                
        df = pd.read_excel(file_bytes_or_path, sheet_name=sheet_target, header=header_row)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        return df
    except Exception as e:
        st.error(f"Error al procesar el Excel: {e}")
        return pd.DataFrame()

if uploaded_file is not None:
    df_raw = cargar_datos_excel(uploaded_file)
else:
    default_path = 'reporte_septimbre_indicadores_integras.xlsx'
    if os.path.exists(default_path):
        df_raw = cargar_datos_excel(default_path)
        st.sidebar.success(f"📁 Usando archivo predeterminado: {default_path}")
    else:
        st.info("👋 Por favor, carga tu archivo Excel en la barra lateral.")
        st.stop()

if df_raw.empty:
    st.warning("El DataFrame está vacío. Revisa la estructura del archivo Excel.")
    st.stop()

# Mapeo robusto de columnas del Excel consolidado
def buscar_col(posibles):
    for col in df_raw.columns:
        if any(p.lower() in str(col).lower() for p in posibles):
            return col
    return None

c_ong = buscar_col(['organización', 'ong', 'socio'])
c_estado = buscar_col(['estado'])
c_muni = buscar_col(['municipio'])
c_sector = buscar_col(['resultado', 'sector'])
c_act = buscar_col(['actividad'])
c_sexo = buscar_col(['sexo'])
c_edad = buscar_col(['edad_calculada', 'edad'])
c_id = buscar_col(['codigoid', 'id'])

df_raw['ONG'] = df_raw[c_ong].astype(str).str.upper().str.strip() if c_ong else 'COOPI'
df_raw['Estado'] = df_raw[c_estado].astype(str).str.strip() if c_estado else 'Distrito Capital'
df_raw['Municipio'] = df_raw[c_muni].astype(str).str.strip() if c_muni else 'Libertador'
df_raw['Sector'] = df_raw[c_sector].astype(str).str.strip() if c_sector else 'Protección'
df_raw['Actividad'] = df_raw[c_act].astype(str).str.strip() if c_act else 'General'
df_raw['Sexo'] = df_raw[c_sexo].astype(str).str.capitalize() if c_sexo else 'Otro'
df_raw['Edad'] = pd.to_numeric(df_raw[c_edad], errors='coerce').fillna(25) if c_edad else 25
df_raw['ID_Unico'] = df_raw[c_id].astype(str) if c_id else [f"ID_{i}" for i in range(len(df_raw))]

def calcular_grupo(row):
    e, s = row['Edad'], row['Sexo']
    if e < 18:
        return 'Niña' if s in ['Femenino', 'Mujer'] else 'Niño'
    else:
        return 'Mujer' if s in ['Femenino', 'Mujer'] else 'Hombre'

df_raw['Grupo_Demografico'] = df_raw.apply(calcular_grupo, axis=1)

# -----------------------------------------------------------------------------
# 3. FILTROS LATERALES
# -----------------------------------------------------------------------------
st.sidebar.markdown('---')
st.sidebar.header('Filtros de Consulta General')

socios_disp = ['Todos'] + sorted([x for x in df_raw['ONG'].dropna().unique() if x != 'nan'])
socio_sel = st.sidebar.selectbox('Socio / ONG:', socios_disp)

estados_disp = ['Todos'] + sorted([x for x in df_raw['Estado'].dropna().unique() if x != 'nan'])
estado_sel = st.sidebar.selectbox('Estado:', estados_disp)

df_temp = df_raw if estado_sel == 'Todos' else df_raw[df_raw['Estado'] == estado_sel]
munis_disp = ['Todos'] + sorted([x for x in df_temp['Municipio'].dropna().unique() if x != 'nan'])
muni_sel = st.sidebar.selectbox('Municipio:', munis_disp)

sectores_disp = ['Todos'] + sorted([x for x in df_raw['Sector'].dropna().unique() if x != 'nan'])
sector_sel = st.sidebar.selectbox('Sector de Implementación:', sectores_disp)

sexo_disp = ['Todos', 'Hombre', 'Mujer', 'Niña', 'Niño']
sexo_sel = st.sidebar.selectbox('Grupo Demográfico / Sexo:', sexo_disp)

# Aplicar filtros
df_filtered = df_raw.copy()
if socio_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['ONG'] == socio_sel]
if estado_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Estado'] == estado_sel]
if muni_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Municipio'] == muni_sel]
if sector_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Sector'] == sector_sel]
if sexo_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Grupo_Demografico'] == sexo_sel]

# -----------------------------------------------------------------------------
# 4. MÉTRICAS CLAVE
# -----------------------------------------------------------------------------
total_impactados = len(df_filtered)
df_unicos = df_filtered.drop_duplicates(subset=['ID_Unico'])
total_unicos = len(df_unicos)
pct_meta = (total_unicos / META_PARTICIPANTES_UNICOS) * 100

col1, col2, col3 = st.columns(3)
col1.metric('Total de Servicios a Participantes', f'{total_impactados:,}')
col2.metric('Total Participantes Únicos', f'{total_unicos:,}')
col3.metric('% Alcance de la Meta (46.122 pers.)', f'{pct_meta:.2f}%')

st.markdown('---')

# -----------------------------------------------------------------------------
# 5. GRÁFICOS INTERACTIVOS
# -----------------------------------------------------------------------------
g1, g2 = st.columns(2)

with g1:
    st.subheader('Desglose por Grupo Demográfico')
    if total_unicos > 0:
        df_demo = df_unicos.groupby('Grupo_Demografico').size().reset_index(name='Cantidad')
        fig_pie = px.pie(df_demo, values='Cantidad', names='Grupo_Demografico', hole=0.4, color_discrete_sequence=PALETA_INTEGRAS)
        fig_pie.update_traces(textinfo='label+value+percent')
        st.plotly_chart(fig_pie, width='stretch')

with g2:
    st.subheader('Participantes Únicos por Sector')
    if total_unicos > 0:
        df_sec = df_unicos.groupby('Sector').size().reset_index(name='Unicos')
        fig_bar = px.bar(df_sec, x='Sector', y='Unicos', color='Sector', color_discrete_sequence=PALETA_INTEGRAS)
        st.plotly_chart(fig_bar, width='stretch')

st.markdown('---')
