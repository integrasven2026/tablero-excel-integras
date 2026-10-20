import io
import os
import re
import folium
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS (CSS)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title='Tablero Consorcio Íntegras | COOPI',
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

# METAS OFICIALES DE ACTIVIDADES DEL PROYECTO
METAS_ACTIVIDADES = {
    'R1A1': {
        'Sector': 'Protección',
        'Nombre': 'R1A1: Protección infantil (Individual)',
        'Meta_Proyecto': 3194,
        'Metas_Socios': {'COOPI': 1200, 'HIAS': 1000, 'FLM': 994},
    },
    'R1A3': {
        'Sector': 'Protección',
        'Nombre': 'R1A3: Asistencia jurídica',
        'Meta_Proyecto': 1047,
        'Metas_Socios': {'COOPI': 400, 'HIAS': 350, 'FLM': 297},
    },
    'R1A4': {
        'Sector': 'Protección',
        'Nombre': 'R1A4: Apoyo psicosocial',
        'Meta_Proyecto': 2108,
        'Metas_Socios': {'COOPI': 1000, 'FLM': 608, 'HIAS': 500},
    },
    'R1A5': {
        'Sector': 'Protección',
        'Nombre': 'R1A5: Sesiones informativas y sensibilización',
        'Meta_Proyecto': 4635,
        'Metas_Socios': {'COOPI': 2000, 'HIAS': 1500, 'FLM': 1135},
    },
    'R1A6': {
        'Sector': 'Protección',
        'Nombre': 'R1A6: Asistencia material y directa',
        'Meta_Proyecto': 800,
        'Metas_Socios': {'COOPI': 400, 'HIAS': 400},
    },
    'R2A2': {
        'Sector': 'Salud',
        'Nombre': 'R2A2: Servicios clínicos SSR',
        'Meta_Proyecto': 1144,
        'Metas_Socios': {'PLAFAM': 750, 'PALUZ': 394},
    },
    'R4A4': {
        'Sector': 'WASH',
        'Nombre': 'R4A4: Fortalecimiento capacidades WASH/PCI',
        'Meta_Proyecto': 2800,
        'Metas_Socios': {'COOPI': 1800, 'PALUZ': 1000},
    },
    'R5A1': {
        'Sector': 'Respuesta a Emergencia',
        'Nombre': 'R5A1: Preparación operativa',
        'Meta_Proyecto': 500,
        'Metas_Socios': {'COOPI': 300, 'HIAS': 200},
    },
    'R5A2': {
        'Sector': 'Respuesta a Emergencia',
        'Nombre': 'R5A2: Respuesta mecanismo modificador de crisis',
        'Meta_Proyecto': 5100,
        'Metas_Socios': {'COOPI': 3000, 'HIAS': 2100},
    },
}

COORDENADAS_MUNICIPIOS = {
    'Libertador': [10.5000, -66.9167],
    'Caroní': [8.2833, -62.7167],
    'Angostura del Orinoco': [7.5333, -63.8833],
    'Pedernales': [9.9667, -62.2500],
    'Tucupita': [9.0622, -62.0531],
    'Paz Castillo': [10.2167, -66.6667],
    'Sucre': [10.4833, -66.8167],
    'Urdaneta': [10.1500, -66.8833],
    'Arismendi': [10.7167, -62.5167],
    'Bermudez': [10.6333, -63.2500],
    'Cruz Salmeron Acosta': [10.6167, -64.2000],
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

# Mapeo robusto de columnas
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
df_raw['Actividad_Raw'] = df_raw[c_act].astype(str).str.strip() if c_act else 'General'
df_raw['Sexo'] = df_raw[c_sexo].astype(str).str.capitalize() if c_sexo else 'Otro'
df_raw['Edad'] = pd.to_numeric(df_raw[c_edad], errors='coerce').fillna(25) if c_edad else 25
df_raw['ID_Unico'] = df_raw[c_id].astype(str) if c_id else [f"ID_{i}" for i in range(len(df_raw))]

def extraer_codigo_act(txt):
    m = re.search(r'(R\d+A\d+)', str(txt), re.IGNORECASE)
    return m.group(1).upper() if m else 'OTRO'

df_raw['Actividad_Cod'] = df_raw['Actividad_Raw'].apply(extraer_codigo_act)

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

# Aplicar filtros generales
df_filtered = df_raw.copy()
if socio_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['ONG'] == socio_sel]
if estado_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Estado'] == estado_sel]
if muni_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Municipio'] == muni_sel]
if sector_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Sector'] == sector_sel]

# -----------------------------------------------------------------------------
# 4. MÉTRICAS CLAVE
# -----------------------------------------------------------------------------
total_servicios = len(df_filtered)
df_unicos = df_filtered.drop_duplicates(subset=['ID_Unico'])
total_unicos = len(df_unicos)
pct_meta = (total_unicos / META_PARTICIPANTES_UNICOS) * 100

col1, col2, col3 = st.columns(3)
col1.metric('Total de Servicios a Participantes', f'{total_servicios:,}')
col2.metric('Total Participantes Únicos', f'{total_unicos:,}')
col3.metric('% Alcance de la Meta (46.122 pers.)', f'{pct_meta:.2f}%')

st.markdown('---')

# -----------------------------------------------------------------------------
# 5. SEGUIMIENTO DE ACTIVIDADES (GRÁFICO DUAL LIMPIO Y GARANTIZADO)
# -----------------------------------------------------------------------------
st.subheader('Seguimiento de Actividades por Socio, Sector y Localidad')

col_act1, col_act2, col_act3 = st.columns(3)

with col_act1:
    socios_act_list = ['TODOS'] + sorted(list(set(df_filtered['ONG'].unique()).union({'COOPI', 'HIAS', 'FLM', 'PALUZ', 'PLAFAM'})))
    socio_reporte = st.selectbox('Filtrar Reporte por Socio:', options=socios_act_list, key='act_socio_sel')

with col_act2:
    sectores_act_list = ['TODOS'] + sorted(list(set(d['Sector'] for d in METAS_ACTIVIDADES.values())))
    sector_reporte = st.selectbox('Filtrar por Sector:', options=sectores_act_list, key='act_sector_sel')

with col_act3:
    localidades_list = ['TODOS'] + sorted([x for x in df_filtered['Estado'].dropna().unique() if x])
    localidad_reporte = st.selectbox('Filtrar por Localidad (Estado):', options=localidades_list, key='act_localidad_sel')

df_act_base = df_filtered.copy()
if localidad_reporte != 'TODOS':
    df_act_base = df_act_base[df_act_base['Estado'] == localidad_reporte]

filas_reporte = []

for cod_act, datos in METAS_ACTIVIDADES.items():
    sec = datos['Sector']
    nombre_act = datos['Nombre']
    meta_proy = datos['Meta_Proyecto']
    metas_socios = datos['Metas_Socios']

    if sector_reporte != 'TODOS' and sec != sector_reporte:
        continue

    if socio_reporte != 'TODOS':
        if socio_reporte not in metas_socios:
            continue
        meta_socio = metas_socios[socio_reporte]

        df_act = df_act_base[
            (df_act_base['ONG'] == socio_reporte) & 
            (df_act_base['Actividad_Cod'] == cod_act)
        ]
        alcanzado_val = len(df_act)
        pct = (alcanzado_val / meta_socio * 100) if meta_socio > 0 else 0.0

        filas_reporte.append({
            'Sector': sec,
            'Actividad': nombre_act,
            'Socio': socio_reporte,
            'Meta Proyecto': meta_proy,
            'Meta Socio': meta_socio,
            'Alcanzados': alcanzado_val,
            '% Avance Socio': round(pct, 1),
        })
    else:
        df_act = df_act_base[df_act_base['Actividad_Cod'] == cod_act]
        alcanzado_val = len(df_act)
        pct = (alcanzado_val / meta_proy * 100) if meta_proy > 0 else 0.0

        filas_reporte.append({
            'Sector': sec,
            'Actividad': nombre_act,
            'Socio': 'TODOS (Consorcio)',
            'Meta Proyecto': meta_proy,
            'Meta Socio': meta_proy,
            'Alcanzados': alcanzado_val,
            '% Avance Socio': round(pct, 1),
        })

df_reporte_act = pd.DataFrame(filas_reporte)

if not df_reporte_act.empty:
    st.dataframe(
        df_reporte_act.style.format({
            'Meta Proyecto': '{:,}',
            'Meta Socio': '{:,}',
            'Alcanzados': '{:,}',
            '% Avance Socio': '{:.1f}%',
        }),
        width='stretch',
        hide_index=True,
    )

    buffer_act = io.BytesIO()
    with pd.ExcelWriter(buffer_act, engine='openpyxl') as writer:
        df_reporte_act.to_excel(writer, index=False, sheet_name='Actividades')
    buffer_act.seek(0)

    st.download_button(
        label='📥 Descargar Seguimiento de Actividades en Excel',
        data=buffer_act,
        file_name='Seguimiento_Actividades_Consorcio_Integras.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    st.markdown('<br>', unsafe_allow_html=True)

    # PREPARACIÓN DE DATOS EN FORMATO LONG PARA PLOTLY EXPRESS (GARANTIZA VISUALIZACIÓN)
    df_melted = df_reporte_act.melt(
        id_vars=['Actividad'],
        value_vars=['Meta Socio', 'Alcanzados'],
        var_name='Tipo',
        value_name='Valor'
    )

    # Mapear etiquetas y orden
    df_melted['Tipo'] = df_melted['Tipo'].replace({'Meta Socio': 'Meta de la Actividad', 'Alcanzados': 'Alcanzados y Avance'})
    
    # Ordenamos el dataframe base de mayor a menor por Meta para el eje Y
    orden_actividades = df_reporte_act.sort_values(by='Meta Socio', ascending=True)['Actividad'].tolist()

    fig_act = px.bar(
        df_melted,
        x='Valor',
        y='Actividad',
        color='Tipo',
        barmode='group',
        orientation='h',
        category_orders={'Actividad': orden_actividades, 'Tipo': ['Meta de la Actividad', 'Alcanzados y Avance']},
        color_discrete_map={
            'Meta de la Actividad': '#08327D',
            'Alcanzados y Avance': COLOR_ROSADO_AAP  # Color violeta / rosado orquídea (#D89FE3)
        },
        height=550,
        title='Resultados por Actividad (Ordenado)'
    )

    fig_act.update_layout(
        font=font_layout,
        xaxis_title='Servicios / Personas',
        yaxis_title='Actividad',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )

    st.plotly_chart(fig_act, width='stretch')
else:
    st.info('No se registraron actividades correspondientes a los filtros seleccionados.')

st.markdown('---')

# -----------------------------------------------------------------------------
# 6. GRÁFICOS: TORTA (DONUT) Y SECTOR
# -----------------------------------------------------------------------------
g1, g2 = st.columns(2)

with g1:
    st.subheader('Desglose por Sexo y Rango Etario')
    if total_unicos > 0:
        df_demo = df_unicos.groupby('Grupo_Demografico').size().reset_index(name='Cantidad')
        fig_pie = px.pie(
            df_demo, values='Cantidad', names='Grupo_Demografico', hole=0.4,
            title='Participantes Únicos por Rango Etario y Sexo',
            color_discrete_sequence=PALETA_INTEGRAS
        )
        fig_pie.update_traces(textinfo='label+value+percent')
        fig_pie.update_layout(showlegend=True, font=font_layout)
        st.plotly_chart(fig_pie, width='stretch')

with g2:
    st.subheader('Participantes Únicos por Sector')
    if total_unicos > 0:
        df_sec = df_unicos.groupby('Sector').size().reset_index(name='Unicos')
        df_sec['Porcentaje'] = (df_sec['Unicos'] / total_unicos) * 100
        df_sec['Etiqueta'] = df_sec.apply(lambda r: f"{r['Unicos']:,} ({r['Porcentaje']:.1f}%)", axis=1)
        
        fig_bar = px.bar(
            df_sec, x='Sector', y='Unicos', color='Sector',
            text='Etiqueta', title='Distribución por Sector de Implementación',
            color_discrete_sequence=PALETA_INTEGRAS
        )
        fig_bar.update_traces(textposition='outside')
        fig_bar.update_layout(showlegend=False, font=font_layout)
        st.plotly_chart(fig_bar, width='stretch')

st.markdown('---')

# -----------------------------------------------------------------------------
# 7. MUNICIPIOS ORDENADOS Y MAPA INTERACTIVO
# -----------------------------------------------------------------------------
m1, m2 = st.columns([1, 1])

with m1:
    st.subheader('Población Única por Municipio (Ordenado)')
    if total_unicos > 0:
        df_muni = df_unicos.groupby(['Estado', 'Municipio']).size().reset_index(name='Participantes_Unicos')
        df_muni = df_muni.sort_values(by='Participantes_Unicos', ascending=True)
        
        fig_muni = px.bar(
            df_muni, y='Municipio', x='Participantes_Unicos', color='Estado',
            orientation='h', text='Participantes_Unicos', height=450,
            color_discrete_sequence=PALETA_INTEGRAS
        )
        fig_muni.update_traces(textposition='outside')
        fig_muni.update_layout(showlegend=True, font=font_layout, xaxis_title='Participantes Únicos', yaxis_title='Municipio')
        st.plotly_chart(fig_muni, width='stretch')

with m2:
    st.subheader('Mapa de Cobertura Territorial')
    
    mapa = folium.Map(location=[7.8, -65.5], zoom_start=6, tiles='CartoDB positron')
    
    if total_unicos > 0:
        muni_mapa = df_filtered.groupby(['Estado', 'Municipio']).agg(
            Participantes_Unicos=('ID_Unico', 'nunique'),
            Total_Servicios=('ID_Unico', 'count')
        ).reset_index()
        
        for _, m_row in muni_mapa.iterrows():
            est = m_row['Estado']
            mun = m_row['Municipio']
            unicos = m_row['Participantes_Unicos']
            servicios = m_row['Total_Servicios']
            
            coords = COORDENADAS_MUNICIPIOS.get(mun, [7.8, -65.5])
            
            popup_html = f"""
            <div style='font-family: Quicksand; font-size: 13px; width: 190px;'>
                <h4 style='color: {COLOR_AGUAMARINA}; margin-bottom: 5px;'>{mun}</h4>
                <b>Estado:</b> {est}<br>
                <b>Participantes Únicos:</b> {unicos:,}<br>
                <b>Total Servicios:</b> {servicios:,}
            </div>
            """
            
            folium.CircleMarker(
                location=coords,
                radius=min(unicos * 2, 22) + 6,
                popup=folium.Popup(popup_html, max_width=220),
                color=COLOR_AGUAMARINA,
                fill=True,
                fill_color=COLOR_AGUAMARINA,
                fill_opacity=0.7,
            ).add_to(mapa)
            
    st_folium(mapa, width=500, height=430)

st.markdown('---')
