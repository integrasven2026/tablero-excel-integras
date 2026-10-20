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
    page_title='Tablero de Monitoreo Excel | Consorcio Íntegras',
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

    .titulo-aap {
        font-family: 'Now', 'Montserrat', sans-serif !important;
        color: #D89FE3 !important;
        margin-top: 15px !important;
        margin-bottom: 5px !important;
        font-weight: 800 !important;
        font-size: 1.8rem !important;
    }

    .titulo-indicadores-aap {
        font-family: 'Now', 'Montserrat', sans-serif !important;
        color: #E5B130 !important;
        margin-top: 25px !important;
        margin-bottom: 10px !important;
        font-weight: 800 !important;
        font-size: 1.8rem !important;
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
        "<h1 class='titulo-principal'>Tablero de Monitoreo - Consorcio Íntegras (Excel)</h1>",
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

# METAS DEL PROYECTO
META_PARTICIPANTES_UNICOS = 46122
META_BENEFICIARIOS_PROYECTO = 32000
META_5_PORCIENTO = META_BENEFICIARIOS_PROYECTO * 0.05  # 1,600 personas

MESES_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
}

MAPA_ESTADOS = {
    'VE01': 'Distrito Capital', 'VE07': 'Bolívar', 'VE10': 'Delta Amacuro',
    'VE15': 'Miranda', 'VE19': 'Sucre', 'VE24': 'La Guaira',
    'Distrito Capital': 'Distrito Capital', 'Bolívar': 'Bolívar',
    'Delta Amacuro': 'Delta Amacuro', 'Miranda': 'Miranda',
    'Sucre': 'Sucre', 'La Guaira': 'La Guaira'
}

MAPA_MUNICIPIOS = {
    'Libertador': [10.5000, -66.9167],
    'Caroní': [8.2833, -62.7167],
    'Angostura': [7.5333, -63.8833],
    'Pedernales': [9.9667, -62.2500],
    'Tucupita': [9.0622, -62.0531],
    'Paz Castillo': [10.2167, -66.6667],
    'Sucre (Miranda)': [10.4833, -66.8167],
    'Urdaneta': [10.1500, -66.8833],
    'Arismendi': [10.7167, -62.5167],
    'Bermúdez': [10.6333, -63.2500],
    'Cruz Salmerón Acosta': [10.6167, -64.2000],
    'Sucre (Sucre)': [10.2833, -63.8833],
    'Vargas': [10.6000, -66.9333],
}

font_layout = dict(family='Quicksand', size=13)

# -----------------------------------------------------------------------------
# 2. CARGA DE DATOS DESDE ARCHIVO EXCEL (SUBIDO POR EL USUARIO)
# -----------------------------------------------------------------------------
st.sidebar.header('Carga de Datos (Excel)')
uploaded_file = st.sidebar.file_uploader(
    "Sube el archivo Excel de reporte consolidado", 
    type=['xlsx', 'xls']
)

@st.cache_data
def procesar_excel(file_bytes):
    try:
        xls = pd.ExcelFile(file_bytes)
        sheet_names = xls.sheet_names
        # Si existe una hoja específica de indicadores o principal, la leemos, sino la primera
        target_sheet = sheet_names[0]
        for s in sheet_names:
            if 'indicador' in s.lower() or 'principal' in s.lower() or 'base' in s.lower():
                target_sheet = s
                break
        df = pd.read_excel(file_bytes, sheet_name=target_sheet)
        return df, sheet_names
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        return pd.DataFrame(), []

if uploaded_file is not None:
    df_raw, hojas_disponibles = procesar_excel(uploaded_file)
else:
    # Ruta de respaldo local si el usuario colocó el archivo en el repositorio
    default_path = 'reporte_septimbre_indicadores_integras.xlsx'
    if os.path.exists(default_path):
        df_raw, hojas_disponibles = procesar_excel(default_path)
        st.sidebar.success(f"📁 Usando archivo predeterminado: {default_path}")
    else:
        st.info("👋 Por favor, carga un archivo Excel en la barra lateral para iniciar el tablero.")
        st.stop()

if df_raw.empty:
    st.warning("El archivo Excel cargado está vacío o no se pudo procesar.")
    st.stop()

# Normalización de columnas clave para asegurar compatibilidad
df_raw.columns = [str(c).strip() for c in df_raw.columns]

# Mapeo flexible de columnas si tienen nombres diferentes
def buscar_columna(df, posibles_nombres):
    for col in df.columns:
        if any(p.lower() in col.lower() for p in posibles_nombres):
            return col
    return None

col_ong = buscar_columna(df_raw, ['ong', 'socio', 'organizacion'])
col_estado = buscar_columna(df_raw, ['estado', 'state'])
col_muni = buscar_columna(df_raw, ['municipio', 'muni'])
col_sector = buscar_columna(df_raw, ['sector', 'resultado', 'rango'])
col_actividad = buscar_columna(df_raw, ['actividad', 'activity'])
col_sexo = buscar_columna(df_raw, ['sexo', 'genero'])
col_edad = buscar_columna(df_raw, ['edad', 'age'])
col_id = buscar_columna(df_raw, ['id', 'codigo', 'documento', 'cedula'])
col_fecha = buscar_columna(df_raw, ['fecha', 'date'])

# Creamos columnas estándar normalizadas en el DataFrame
df_raw['ONG'] = df_raw[col_ong].astype(str).str.upper().str.strip() if col_ong else 'COOPI'
df_raw['Estado'] = df_raw[col_estado].astype(str).str.strip() if col_estado else 'Distrito Capital'
df_raw['Municipio'] = df_raw[col_muni].astype(str).str.strip() if col_muni else 'Libertador'
df_raw['Sector'] = df_raw[col_sector].astype(str).str.strip() if col_sector else 'Protección'
df_raw['Actividad'] = df_raw[col_actividad].astype(str).str.strip() if col_actividad else 'General'
df_raw['Sexo'] = df_raw[col_sexo].astype(str).str.capitalize() if col_sexo else 'Otro'
df_raw['Edad'] = pd.to_numeric(df_raw[col_edad], errors='coerce').fillna(25) if col_edad else 25

# Identificador único por fila o registro
if col_id:
    df_raw['ID_Unico'] = df_raw[col_id].astype(str)
else:
    df_raw['ID_Unico'] = [f"REG_{i}" for i in range(len(df_raw))]

# Grupo demográfico
def calcular_grupo(row):
    edad = row['Edad']
    sexo = row['Sexo']
    if edad < 18:
        return 'Niña' if sexo in ['Femenino', 'Mujer'] else 'Niño'
    else:
        return 'Mujer' if sexo in ['Femenino', 'Mujer'] else 'Hombre'

df_raw['Grupo_Demografico'] = df_raw.apply(calcular_grupo, axis=1)

# Fechas y Meses
if col_fecha:
    df_raw['Fecha_DT'] = pd.to_datetime(df_raw[col_fecha], errors='coerce')
    df_raw['Mes_Reporte'] = df_raw['Fecha_DT'].apply(lambda x: f"{x.year} - {MESES_ES.get(x.month, '')}" if pd.notnull(x) else 'Septiembre 2026')
else:
    df_raw['Mes_Reporte'] = 'Septiembre 2026'

# Indicadores de vulnerabilidad simulados o adaptados si existen en el excel
for col_vuln in ['Es_Discapacidad', 'Es_Indigena', 'Es_Embarazada_Lactante', 'Es_LGBTIQ']:
    if col_vuln not in df_raw.columns:
        df_raw[col_vuln] = 0

# -----------------------------------------------------------------------------
# 3. FILTROS LATERALES
# -----------------------------------------------------------------------------
st.sidebar.markdown('---')
st.sidebar.header('Filtros de Consulta General')

meses_ordenados = sorted([m for m in df_raw['Mes_Reporte'].unique() if m != 'Sin Fecha'])
meses_disp = ['Todos'] + meses_ordenados
mes_sel = st.sidebar.selectbox('Mes del Reporte:', meses_disp, key='f_mes')

socios_disp = ['Todos'] + sorted([x for x in df_raw['ONG'].dropna().unique() if x])
socio_sel = st.sidebar.selectbox('Socio / ONG:', socios_disp, key='f_socio')

estados_disp = ['Todos'] + sorted([x for x in df_raw['Estado'].dropna().unique() if x])
estado_sel = st.sidebar.selectbox('Estado:', estados_disp, key='f_estado')

df_temp = df_raw if estado_sel == 'Todos' else df_raw[df_raw['Estado'] == estado_sel]
munis_disp = ['Todos'] + sorted([x for x in df_temp['Municipio'].dropna().unique() if x])
muni_sel = st.sidebar.selectbox('Municipio:', munis_disp, key='f_muni')

sectores_disp = ['Todos'] + sorted([x for x in df_raw['Sector'].dropna().unique() if x])
sector_sel = st.sidebar.selectbox('Sector de Implementación:', sectores_disp, key='f_sector')

sexo_disp = ['Todos', 'Hombre', 'Mujer', 'Otro']
sexo_sel = st.sidebar.selectbox('Sexo del Participante:', sexo_disp, key='f_sexo')

grupo_demo_disp = ['Todos', 'Mujer', 'Hombre', 'Niña', 'Niño']
grupo_demo_sel = st.sidebar.selectbox('Grupo Demográfico:', grupo_demo_disp, key='f_grupo_demo')

# Aplicar Filtros Generales
df_filtered = df_raw.copy()
if mes_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Mes_Reporte'] == mes_sel]
if socio_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['ONG'] == socio_sel]
if estado_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Estado'] == estado_sel]
if muni_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Municipio'] == muni_sel]
if sector_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Sector'] == sector_sel]
if sexo_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Sexo'] == sexo_sel]
if grupo_demo_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Grupo_Demografico'] == grupo_demo_sel]

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
# 5. GRÁFICOS INTERACTIVOS (POR PARTICIPANTES ÚNICOS)
# -----------------------------------------------------------------------------
g1, g2 = st.columns(2)

with g1:
    st.subheader('Desglose por Sexo y Rango Etario')
    if total_unicos > 0:
        df_demo = df_unicos.groupby('Grupo_Demografico').size().reset_index(name='Cantidad')
        fig_pie_demo = px.pie(
            df_demo, values='Cantidad', names='Grupo_Demografico', hole=0.4,
            title='Participantes Únicos por Rango Etario y Sexo',
            color_discrete_sequence=PALETA_INTEGRAS,
        )
        fig_pie_demo.update_traces(textinfo='label+value+percent')
        fig_pie_demo.update_layout(showlegend=True, font=font_layout)
        st.plotly_chart(fig_pie_demo, width='stretch')

with g2:
    st.subheader('Participantes Únicos por Sector')
    if total_unicos > 0:
        df_sec_unicos = df_filtered.drop_duplicates(subset=['ID_Unico', 'Sector'])
        df_sec_cnt = df_sec_unicos['Sector'].value_counts().reset_index()
        df_sec_cnt.columns = ['Sector', 'Unicos']
        fig_bar_sec = px.bar(
            df_sec_cnt, x='Sector', y='Unicos', color='Sector',
            title='Distribución por Sector de Implementación',
            color_discrete_sequence=PALETA_INTEGRAS,
        )
        fig_bar_sec.update_layout(showlegend=False, font=font_layout)
        st.plotly_chart(fig_bar_sec, width='stretch')

st.markdown('---')

# -----------------------------------------------------------------------------
# 6. UBICACIÓN GEOGRÁFICA Y MAPA
# -----------------------------------------------------------------------------
st.subheader('Ubicación Geográfica por Municipio')
m1, m2 = st.columns([1, 1])

with m1:
    st.markdown('### Participantes Únicos por Municipio')
    if total_unicos > 0:
        df_muni = df_unicos.groupby(['Estado', 'Municipio']).size().reset_index(name='Participantes_Unicos')
        df_muni = df_muni.sort_values(by='Participantes_Unicos', ascending=True)
        fig_muni = px.bar(
            df_muni, y='Municipio', x='Participantes_Unicos', color='Estado',
            orientation='h', height=400, color_discrete_sequence=PALETA_INTEGRAS,
        )
        fig_muni.update_layout(showlegend=False, font=font_layout)
        st.plotly_chart(fig_muni, width='stretch')

with m2:
    st.markdown('### Cobertura de la Intervención')
    mapa = folium.Map(location=[7.8, -65.5], zoom_start=6, tiles='CartoDB positron')
    if total_unicos > 0:
        muni_totales = df_unicos.groupby(['Estado', 'Municipio']).size().reset_index(name='Total_Unicos')
        for _, m_row in muni_totales.iterrows():
            est, mun, tot = m_row['Estado'], m_row['Municipio'], m_row['Total_Unicos']
            coords = MAPA_MUNICIPIOS.get(mun, [7.8, -65.5])
            folium.CircleMarker(
                location=coords, radius=min(tot * 3, 20) + 6,
                popup=f"<b>{mun} ({est})</b><br>Participantes: {tot}",
                color=COLOR_AGUAMARINA, fill=True, fill_color=COLOR_AGUAMARINA, fill_opacity=0.7,
            ).add_to(mapa)
    st_folium(mapa, width=500, height=380)

st.markdown('---')

# -----------------------------------------------------------------------------
# 7. DESCARGA DE DATOS FILTRADOS EN EXCEL
# -----------------------------------------------------------------------------
st.subheader('Exportar Datos Consolidados')
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, index=False, sheet_name='Consolidado_Integras')
buffer.seek(0)

st.download_button(
    label='📥 Descargar Consolidado Filtrado en Excel',
    data=buffer,
    file_name='Consolidado_Consorcio_Integras.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
)