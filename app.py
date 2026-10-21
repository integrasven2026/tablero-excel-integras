import io
import os
import re
import folium
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_folium import st_folium

# -----------------------------------------------------------------------------
# PALETA DE COLORES OFICIAL CONSORCIO INTEGRAS / COOPI
# -----------------------------------------------------------------------------
COLOR_AGUAMARINA = '#17C3B2'  # Verde / Azul Agua Marina oficial
COLOR_ROSADO_AAP = '#D89FE3'  # Morado / Rosado Orquídea para Reporte AAP
COLOR_VERDE_ABIERTO = '#28A745'  # Verde para Casos Abiertos / En Proceso
COLOR_AMARILLO_MOSTAZA = '#E5B130'  # Amarillo Mostaza para Indicadores AAP
COLOR_AZUL_COOPI = '#0072CE'
COLOR_AZUL_OSCURO = '#08327D'
COLOR_NARANJA_ABIERTO = '#E67E22'

PALETA_INTEGRAS = [
    COLOR_AGUAMARINA,
    COLOR_ROSADO_AAP,
    COLOR_AMARILLO_MOSTAZA,
    COLOR_AZUL_OSCURO,
    COLOR_AZUL_COOPI,
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
    logo_path = 'integras.jpg' if os.path.exists('integras.jpg') else ('logo_coopi.jpg' if os.path.exists('logo_coopi.jpg') else None)
    if logo_path:
        st.image(logo_path, width=180)
    else:
        st.warning("⚠️ No se pudo cargar el logo.")

st.markdown('---')

META_PARTICIPANTES_UNICOS = 34131
META_5_PORCIENTO = META_PARTICIPANTES_UNICOS * 0.05

COORDENADAS_MUNICIPIOS = {
    'Libertador': [10.5061, -66.9146],
    'Caroní': [8.3538, -62.6455],
    'Caroni': [8.3538, -62.6455],
    'Angostura del Orinoco': [8.1256, -63.5424],
    'Pedernales': [9.9833, -62.2167],
    'Tucupita': [9.0622, -62.0531],
    'Paz Castillo': [10.2167, -66.6667],
    'Sucre': [10.4806, -66.8036],
    'Urdaneta': [10.1583, -66.8964],
    'Arismendi': [10.7042, -63.1111],
    'Bermudez': [10.6667, -63.2500],
    'Cruz Salmeron Acosta': [10.5583, -64.2583],
    'Vargas': [10.6027, -66.9325],
}

MAPA_ESTADOS = {
    'VE01': 'Distrito Capital', 'VE07': 'Bolívar', 'VE10': 'Delta Amacuro',
    'VE15': 'Miranda', 'VE19': 'Sucre', 'VE24': 'La Guaira',
    'Distrito Capital': 'Distrito Capital', 'Bolívar': 'Bolívar',
    'Delta Amacuro': 'Delta Amacuro', 'Miranda': 'Miranda',
    'Sucre': 'Sucre', 'La Guaira': 'La Guaira',
}

MESES_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
}

font_layout = dict(family='Quicksand', size=13)

# -----------------------------------------------------------------------------
# 2. CARGA DE DATOS DESDE EXCEL (CON OPCIÓN DE SUBIR ARCHIVO LOCAL O USAR EL DE GITHUB)
# -----------------------------------------------------------------------------
st.sidebar.header('Carga de Datos (Excel)')
uploaded_file = st.sidebar.file_uploader(
    "Sube un archivo Excel alternativo (Opcional)", 
    type=['xlsx', 'xls']
)

def obtener_bytes_excel(file_or_path):
    if hasattr(file_or_path, 'read'):
        file_or_path.seek(0)
        return io.BytesIO(file_or_path.read())
    elif isinstance(file_or_path, str) and os.path.exists(file_or_path):
        with open(file_or_path, 'rb') as f:
            return io.BytesIO(f.read())
    return None

if uploaded_file is not None:
    source_ref = uploaded_file
    st.sidebar.success("📁 Usando archivo subido manualmente.")
else:
    default_path = 'reporte_septimbre_indicadores_integras.xlsx'
    if os.path.exists(default_path):
        source_ref = default_path
        st.sidebar.success("📁 Base de datos cargada automáticamente.")
    else:
        st.error(f"❌ No se encontró el archivo predeterminado '{default_path}'.")
        st.stop()

@st.cache_data
def cargar_datos_crudos(file_bytes_io):
    try:
        io_copy = io.BytesIO(file_bytes_io.getvalue())
        xls = pd.ExcelFile(io_copy)
        sheet_target = 'group_beneficiario' if 'group_beneficiario' in xls.sheet_names else xls.sheet_names[0]
        
        io_copy2 = io.BytesIO(file_bytes_io.getvalue())
        df_test = pd.read_excel(io_copy2, sheet_name=sheet_target, header=None)
        header_row = 0
        for idx, row in df_test.iloc[:15].iterrows():
            row_str = " ".join([str(v) for v in row.values]).lower()
            if 'codigoid' in row_str or 'nombres' in row_str:
                header_row = idx
                break
                
        io_copy3 = io.BytesIO(file_bytes_io.getvalue())
        df = pd.read_excel(io_copy3, sheet_name=sheet_target, header=header_row)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        def buscar_col(posibles):
            for col in df.columns:
                if any(p.lower() in str(col).lower() for p in posibles):
                    return col
            return None

        c_ong = buscar_col(['organización', 'ong', 'socio'])
        c_estado = buscar_col(['estado'])
        c_muni = buscar_col(['municipio'])
        c_parroquia = buscar_col(['parroquia'])
        c_sector = buscar_col(['resultado', 'sector'])
        
        c_act = None
        for col in df.columns:
            if 'actividad' in str(col).lower() and 'fecha' not in str(col).lower():
                c_act = col
                break
        if not c_act:
            c_act = buscar_col(['actividad'])

        c_sexo = buscar_col(['sexo'])
        c_edad = buscar_col(['edad_calculada', 'edad'])
        c_id = buscar_col(['codigoid', 'id'])

        df['ONG'] = df[c_ong].astype(str).str.upper().str.strip() if c_ong else 'COOPI'
        df['Estado'] = df[c_estado].astype(str).str.strip() if c_estado else 'Distrito Capital'
        df['Municipio'] = df[c_muni].astype(str).str.strip() if c_muni else 'Libertador'
        df['Parroquia'] = df[c_parroquia].astype(str).str.strip() if c_parroquia else 'Sin Parroquia'
        df['Sector'] = df[c_sector].astype(str).str.strip() if c_sector else 'Protección'
        df['Actividad_Raw'] = df[c_act].astype(str).str.strip() if c_act else 'General'
        df['Sexo'] = df[c_sexo].astype(str).str.capitalize() if c_sexo else 'Otro'
        df['Edad'] = pd.to_numeric(df[c_edad], errors='coerce').fillna(25) if c_edad else 25
        df['ID_Unico'] = df[c_id].astype(str) if c_id else [f"ID_{i}" for i in range(len(df))]

        def extraer_codigo_act(txt):
            m = re.search(r'(R\d+A\d+)', str(txt), re.IGNORECASE)
            return m.group(1).upper() if m else 'OTRO'

        df['Actividad_Cod'] = df['Actividad_Raw'].apply(extraer_codigo_act)

        def calcular_grupo(row):
            e, s = row['Edad'], row['Sexo']
            if e < 18:
                return 'Niña' if s in ['Femenino', 'Mujer'] else 'Niño'
            else:
                return 'Mujer' if s in ['Femenino', 'Mujer'] else 'Hombre'

        df['Grupo_Demografico'] = df.apply(calcular_grupo, axis=1)
        return df
    except Exception as e:
        st.error(f"Error al procesar el Excel: {e}")
        return pd.DataFrame()

excel_bytes = obtener_bytes_excel(source_ref)
df_raw = cargar_datos_crudos(excel_bytes)

# -----------------------------------------------------------------------------
# 3. FUNCIONES DE LIMPIEZA Y CARGA DE KOBOTOOLBOX (3 FORMULARIOS)
# -----------------------------------------------------------------------------
KOBO_TOKEN = '5618d295bdaac1c331e0395a3ac0699594c9664c'
ASSET_ID_AAP = 'aRbFg8ig22Ts5JFFvsWNaE'
ASSET_ID_SEGUIMIENTO = 'aav4FtiNC62seD5LAaJJJJ'
ASSET_ID_IND_AAP = 'aMYumvwLQ4rQeq5iFDSboS'

def calcular_etiquetas_con_porcentaje(serie_cantidades, total, decimales=1):
    etiquetas = []
    for x in serie_cantidades:
        if total > 0:
            pct = (x / total) * 100
            etiquetas.append(f'{x} ({pct:.{decimales}f}%)')
        else:
            etiquetas.append(str(x))
    return etiquetas

def extraer_valor_booleano(diccionario_beneficiario, lista_posibles_claves):
    val_afirmativos = ['sí', 'si', 'yes', '1', 's', 'true']
    for clave in lista_posibles_claves:
        for k_item, v_item in diccionario_beneficiario.items():
            if clave.lower() in str(k_item).lower():
                if str(v_item).lower().strip() in val_afirmativos or v_item == 1 or v_item is True:
                    return 1
    return 0

def limpiar_canal(texto):
    if not texto or str(texto).lower() in ['none', 'null', '', 'nan', 'sin especificar']:
        return 'Buzón'
    t_low = str(texto).strip().lower()
    if 'buzon' in t_low: return 'Buzón'
    elif 'telefon' in t_low or 'llamada' in t_low or 'tlf' in t_low: return 'Línea Telefónica'
    elif 'whatsapp' in t_low or 'wsp' in t_low or 'mensaje' in t_low: return 'Mensaje de Texto / WhatsApp'
    elif 'cara' in t_low or 'presencial' in t_low: return 'Cara a Cara'
    elif 'correo' in t_low or 'email' in t_low: return 'Correo Electrónico'
    return 'Buzón'

def limpiar_estatus_caso(texto):
    if not texto or str(texto).lower() in ['none', 'null', '', 'nan']:
        return 'Sin Especificar'
    t = str(texto).strip().title()
    if 'abiert' in t.lower(): return 'Abierto'
    elif 'proceso' in t.lower() or 'pend' in t.lower(): return 'En Proceso'
    elif 'cerrad' in t.lower() or 'atendid' in t.lower(): return 'Cerrado'
    return t

def limpiar_satisfaccion(texto):
    if not texto or str(texto).lower() in ['none', 'null', '', 'nan']: return 'Sin Especificar'
    t_low = str(texto).replace('_', ' ').strip().lower()
    if 'muy' in t_low and 'satisfech' in t_low: return 'Muy Satisfecho'
    elif 'satisfech' in t_low: return 'Satisfecho'
    elif 'insatisfech' in t_low: return 'Insatisfecho'
    return t_low.title()

@st.cache_data(ttl=3600)
def cargar_datos_kobo(asset_id, token_aap):
    headers = {'Authorization': f'Token {token_aap}'}
    url = f'https://eu.kobotoolbox.org/api/v2/assets/{asset_id}/data.json'
    try:
        response = requests.get(url, headers=headers)
        if response.status_code != 200: return []
        return response.json().get('results', [])
    except Exception:
        return []

@st.cache_data(ttl=3600)
def cargar_todos_datos_aap(asset_id_aap, asset_id_seg, token_aap):
    data1 = cargar_datos_kobo(asset_id_aap, token_aap)
    data2 = cargar_datos_kobo(asset_id_seg, token_aap)
    
    all_raw = []
    for r in data1: r['_form_source'] = 'form1'; all_raw.append(r)
    for r in data2: r['_form_source'] = 'form2'; all_raw.append(r)
    
    if not all_raw: return pd.DataFrame()

    aap_rows = []
    for r in all_raw:
        canal_raw = r.get('Mecanismo por el cual se recogió la retroalimentación') or 'Buzón'
        tipo_pqrs = str(r.get('Tipo_PQRS') or r.get('Categoría de la retroalimentación') or 'Demanda de Información').title()
        
        discapacidad = extraer_valor_booleano(r, ['discapacidad', 'pcd'])
        indigena = extraer_valor_booleano(r, ['indigena', 'etnia'])
        lgbtiq = extraer_valor_booleano(r, ['lgbtiq', 'lgbt'])
        embarazada = extraer_valor_booleano(r, ['embarazada', 'lactante'])

        sexo_raw = str(r.get('sexo') or r.get('Sexo') or '').lower().strip()
        try: edad = float(r.get('edad') or r.get('Edad') or 0)
        except: edad = 0

        es_nina = 1 if (edad < 18 and sexo_raw in ['femenino', 'f', 'mujer']) else 0
        es_nino = 1 if (edad < 18 and sexo_raw in ['masculino', 'm', 'hombre']) else 0
        
        estado_geo_val = r.get('Estado') or 'Distrito Capital'

        aap_rows.append({
            'Canal': limpiar_canal(canal_raw),
            'Tipo_PQRS': tipo_pqrs,
            'Estado_Caso': limpiar_estatus_caso(r.get('estatus') or r.get('seguimiento')),
            'Fecha': r.get('_submission_time'),
            'Discapacidad': discapacidad,
            'Indigena': indigena,
            'LGBTIQ': lgbtiq,
            'Embarazada': embarazada,
            'Es_Nina': es_nina,
            'Es_Nino': es_nino,
            'Proyecto': 'INTEGRAS',
            'Estado_Geo': MAPA_ESTADOS.get(estado_geo_val, estado_geo_val),
        })

    df_aap = pd.DataFrame(aap_rows)
    if not df_aap.empty and 'Fecha' in df_aap.columns:
        df_aap['Fecha_DT'] = pd.to_datetime(df_aap['Fecha'], errors='coerce')
        df_aap['Mes_Reporte'] = df_aap['Fecha_DT'].apply(lambda x: f"{x.year} - {MESES_ES.get(x.month, '')}" if pd.notnull(x) else 'Sin Fecha')
    else:
        df_aap['Mes_Reporte'] = 'Sin Fecha'
    return df_aap

@st.cache_data(ttl=3600)
def cargar_datos_indicadores_aap(asset_id_ind, token_ind):
    headers = {'Authorization': f'Token {token_ind}'}
    url = f'https://eu.kobotoolbox.org/api/v2/assets/{asset_id_ind}/data.json'
    try:
        response = requests.get(url, headers=headers)
        if response.status_code != 200: return pd.DataFrame()
        data = response.json().get('results', [])
        if not data: return pd.DataFrame()
        df_ind = pd.DataFrame(data)
        sat_col = [c for c in df_ind.columns if 'satisfac' in c.lower()]
        if sat_col:
            df_ind[sat_col[0]] = df_ind[sat_col[0]].apply(limpiar_satisfaccion)
        return df_ind
    except Exception:
        return pd.DataFrame()

df_aap_raw = cargar_todos_datos_aap(ASSET_ID_AAP, ASSET_ID_SEGUIMIENTO, KOBO_TOKEN)
df_eval_aap = cargar_datos_indicadores_aap(ASSET_ID_IND_AAP, KOBO_TOKEN)

# -----------------------------------------------------------------------------
# 4. FILTROS LATERALES GENERALES
# -----------------------------------------------------------------------------
st.sidebar.markdown('---')
st.sidebar.header('Filtros de Consulta General')

socios_disp = ['Todos'] + sorted([x for x in df_raw['ONG'].dropna().unique() if x != 'nan'])
socio_sel = st.sidebar.selectbox('Socio / Organización:', socios_disp)

estados_disp = ['Todos'] + sorted([x for x in df_raw['Estado'].dropna().unique() if x != 'nan'])
estado_sel = st.sidebar.selectbox('Estado:', estados_disp)

df_temp = df_raw if estado_sel == 'Todos' else df_raw[df_raw['Estado'] == estado_sel]
munis_disp = ['Todos'] + sorted([x for x in df_temp['Municipio'].dropna().unique() if x != 'nan'])
muni_sel = st.sidebar.selectbox('Municipio:', munis_disp)

sectores_disp = ['Todos'] + sorted([x for x in df_raw['Sector'].dropna().unique() if x != 'nan'])
sector_sel = st.sidebar.selectbox('Sector de Implementación:', sectores_disp)

# Aplicar filtros generales al DF principal
df_filtered = df_raw.copy()
if socio_sel != 'Todos': df_filtered = df_filtered[df_filtered['ONG'] == socio_sel]
if estado_sel != 'Todos': df_filtered = df_filtered[df_filtered['Estado'] == estado_sel]
if muni_sel != 'Todos': df_filtered = df_filtered[df_filtered['Municipio'] == muni_sel]
if sector_sel != 'Todos': df_filtered = df_filtered[df_filtered['Sector'] == sector_sel]

# -----------------------------------------------------------------------------
# NAVEGACIÓN POR PESTAÑAS (TABLERO PRINCIPAL + RENDICIÓN DE CUENTAS AAP)
# -----------------------------------------------------------------------------
tab_principal, tab_aap = st.tabs(["📊 Tablero de Indicadores y Cobertura", "🗣️ Rendición de Cuentas y PQRS (Íntegras)"])

with tab_principal:
    total_servicios = len(df_filtered)
    df_unicos = df_filtered.drop_duplicates(subset=['ID_Unico'])
    total_unicos = len(df_unicos)
    pct_meta = (total_unicos / META_PARTICIPANTES_UNICOS) * 100

    col1, col2, col3 = st.columns(3)
    col1.metric('Total de Servicios a Participantes', f'{total_servicios:,}')
    col2.metric('Total Participantes Únicos', f'{total_unicos:,}')
    col3.metric(f'% Alcance de la Meta ({META_PARTICIPANTES_UNICOS:,} pers.)'.replace(',', '.'), f'{pct_meta:.2f}%')

    st.markdown('---')

    st.subheader('Seguimiento de Actividades (Metas Oficiales vs Alcanzados)')
    
    @st.cache_data
    def cargar_metas_y_alcanzados_filtrados(file_bytes_io, org_seleccionada, estado_seleccionado, muni_seleccionado, sector_seleccionado):
        try:
            io_copy = io.BytesIO(file_bytes_io.getvalue())
            xls = pd.ExcelFile(io_copy)
            if 'Benef. x Actividad - Automatico' in xls.sheet_names:
                io_copy2 = io.BytesIO(file_bytes_io.getvalue())
                df_act_sheet = pd.read_excel(io_copy2, sheet_name='Benef. x Actividad - Automatico', header=9)
                
                df_detalle = cargar_datos_crudos(file_bytes_io)
                if org_seleccionada != 'Todos': df_detalle = df_detalle[df_detalle['ONG'] == org_seleccionada]
                if estado_seleccionado != 'Todos': df_detalle = df_detalle[df_detalle['Estado'] == estado_seleccionado]
                if muni_seleccionado != 'Todos': df_detalle = df_detalle[df_detalle['Municipio'] == muni_seleccionado]
                if sector_seleccionado != 'Todos': df_detalle = df_detalle[df_detalle['Sector'] == sector_seleccionado]

                alcanzados_por_act = df_detalle.groupby('Actividad_Cod').size().to_dict()

                datos_actividades = []
                for _, r in df_act_sheet.iterrows():
                    fila_txt = str(r.iloc[0])
                    if pd.notnull(fila_txt) and any(f'R{i}A' in fila_txt for i in range(1, 6)):
                        alcanzado_excel = r.iloc[13] if len(r) > 13 else 0
                        meta_val = r.iloc[14] if len(r) > 14 else 0
                        try:
                            m_int = int(float(meta_val)) if pd.notnull(meta_val) else 0
                            m_cod = re.search(r'(R\d+A\d+)', fila_txt, re.IGNORECASE)
                            codigo_corto = m_cod.group(1).upper() if m_cod else 'ACT'
                            
                            if org_seleccionada == 'Todos' and estado_seleccionado == 'Todos' and muni_seleccionado == 'Todos' and sector_seleccionado == 'Todos':
                                a_int = int(float(alcanzado_excel)) if pd.notnull(alcanzado_excel) else 0
                            else:
                                a_int = int(alcanzados_por_act.get(codigo_corto, 0))

                            pct = (a_int / m_int * 100) if m_int > 0 else 0.0
                            sec = 'Protección'
                            if 'R2' in fila_txt: sec = 'Salud'
                            elif 'R3' in fila_txt: sec = 'Nutrición'
                            elif 'R4' in fila_txt: sec = 'WASH'
                            elif 'R5' in fila_txt: sec = 'Respuesta a Emergencia'
                            
                            datos_actividades.append({
                                'Sector': sec, 'Codigo_Corto': codigo_corto,
                                'Actividad_Completa': fila_txt.strip(),
                                'Meta de la Actividad': m_int, 'Alcanzados y Avance': a_int,
                                '% Avance Socio': round(pct, 1)
                            })
                        except: pass
                return pd.DataFrame(datos_actividades)
        except Exception: pass
        return pd.DataFrame()

    df_reporte_act = cargar_metas_y_alcanzados_filtrados(excel_bytes, socio_sel, estado_sel, muni_sel, sector_sel)

    if not df_reporte_act.empty:
        df_reporte_act = df_reporte_act.sort_values(by='Meta de la Actividad', ascending=True)
        st.dataframe(
            df_reporte_act[['Sector', 'Actividad_Completa', 'Meta de la Actividad', 'Alcanzados y Avance', '% Avance Socio']].style.format({
                'Meta de la Actividad': '{:,}', 'Alcanzados y Avance': '{:,}', '% Avance Socio': '{:.1f}%',
            }),
            width='stretch', hide_index=True,
        )

        buffer_act = io.BytesIO()
        with pd.ExcelWriter(buffer_act, engine='openpyxl') as writer:
            df_reporte_act.to_excel(writer, index=False, sheet_name='Actividades')
        buffer_act.seek(0)

        st.download_button(
            label='📥 Descargar Seguimiento de Actividades en Excel',
            data=buffer_act, file_name='Seguimiento_Actividades_Consorcio_Integras.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        st.markdown('<br>', unsafe_allow_html=True)

        df_melted = df_reporte_act.melt(
            id_vars=['Codigo_Corto', 'Actividad_Completa', '% Avance Socio'],
            value_vars=['Meta de la Actividad', 'Alcanzados y Avance'],
            var_name='Tipo', value_name='Valor'
        )
        df_melted['Etiqueta'] = [f"{r['Valor']:,} ({r['% Avance Socio']:.1f}%)" if r['Tipo'] == 'Alcanzados y Avance' else f"{r['Valor']:,}" for _, r in df_melted.iterrows()]
        
        fig_act = px.bar(
            df_melted, x='Valor', y='Codigo_Corto', color='Tipo', barmode='group', orientation='h', text='Etiqueta',
            color_discrete_map={'Meta de la Actividad': '#08327D', 'Alcanzados y Avance': COLOR_ROSADO_AAP},
            height=520, title='Resultados por Actividad (Metas Oficiales vs Alcanzados)'
        )
        fig_act.update_traces(textposition='outside')
        st.plotly_chart(fig_act, width='stretch')

    st.markdown('---')

    g1, g2 = st.columns(2)
    with g1:
        st.subheader('Desglose por Sexo y Rango Etario')
        if total_unicos > 0:
            df_demo = df_unicos.groupby('Grupo_Demografico').size().reset_index(name='Cantidad')
            fig_pie = px.pie(df_demo, values='Cantidad', names='Grupo_Demografico', hole=0.4, color_discrete_sequence=PALETA_INTEGRAS)
            fig_pie.update_traces(textinfo='label+value+percent')
            fig_pie.update_layout(showlegend=False, font=font_layout)
            st.plotly_chart(fig_pie, width='stretch')

    with g2:
        st.subheader('Participantes Únicos por Sector')
        if total_unicos > 0:
            df_sec = df_unicos.groupby('Sector').size().reset_index(name='Unicos')
            df_sec['Porcentaje'] = (df_sec['Unicos'] / total_unicos) * 100
            df_sec['Etiqueta'] = df_sec.apply(lambda r: f"{r['Unicos']:,} ({r['Porcentaje']:.1f}%)", axis=1)
            fig_bar = px.bar(df_sec, x='Sector', y='Unicos', color='Sector', text='Etiqueta', color_discrete_sequence=PALETA_INTEGRAS)
            fig_bar.update_traces(textposition='outside')
            fig_bar.update_layout(showlegend=False, font=font_layout)
            st.plotly_chart(fig_bar, width='stretch')

    st.markdown('---')

    m1, m2 = st.columns([1, 1])
    with m1:
        st.subheader('Población Única por Municipio')
        if total_unicos > 0:
            df_muni = df_unicos.groupby(['Estado', 'Municipio']).size().reset_index(name='Participantes_Unicos')
            df_muni = df_muni.sort_values(by='Participantes_Unicos', ascending=True)
            fig_muni = px.bar(df_muni, y='Municipio', x='Participantes_Unicos', color='Estado', orientation='h', text='Participantes_Unicos', height=450, color_discrete_sequence=PALETA_INTEGRAS)
            fig_muni.update_traces(textposition='outside')
            st.plotly_chart(fig_muni, width='stretch')

    with m2:
        st.subheader('Mapa de Cobertura y Resultados por Actividad')
        mapa = folium.Map(location=[7.8, -65.5], zoom_start=6, tiles='OpenStreetMap')
        if total_unicos > 0:
            muni_mapa = df_filtered.groupby(['Estado', 'Municipio']).agg(Participantes_Unicos=('ID_Unico', 'nunique'), Total_Servicios=('ID_Unico', 'count')).reset_index()
            for _, m_row in muni_mapa.iterrows():
                mun = m_row['Municipio']
                unicos = m_row['Participantes_Unicos']
                coords = COORDENADAS_MUNICIPIOS.get(mun, [7.8, -65.5])
                
                df_mun_acts = df_filtered[df_filtered['Municipio'] == mun]
                acts_resumen = df_mun_acts.groupby(['Actividad_Cod']).size().reset_index(name='Cant')
                html_acts = "".join([f"• <b>{r['Actividad_Cod']}</b>: {r['Cant']:,} serv.<br>" for _, r in acts_resumen.iterrows()])
                
                popup_html = f"<div style='font-family: Quicksand; font-size: 12px; width: 220px;'><b>{mun}</b><br>Participantes Únicos: {unicos:,}<hr>{html_acts}</div>"
                folium.CircleMarker(location=coords, radius=min(unicos * 0.75, 10) + 3, popup=folium.Popup(popup_html, max_width=240), color=COLOR_AGUAMARINA, fill=True, fill_color=COLOR_AGUAMARINA, fill_opacity=0.75).add_to(mapa)
        st_folium(mapa, width=500, height=430)

with tab_aap:
    st.subheader('🗣️ Módulo de Rendición de Cuentas y PQRS - Consorcio Íntegras')
    
    total_pqrs = len(df_aap_raw)
    pct_cumplimiento_pqrs = (total_pqrs / META_5_PORCIENTO) * 100

    ap1, ap2, ap3 = st.columns(3)
    ap1.metric('Total PQRS Registradas', f'{total_pqrs:,}')
    ap2.metric('Meta AAP (5% Participantes)', f'{int(META_5_PORCIENTO):,}')
    ap3.metric('% Cumplimiento Meta AAP', f'{pct_cumplimiento_pqrs:.2f}%')

    st.markdown('---')

    p1, p2 = st.columns(2)
    with p1:
        st.markdown('### Canal más utilizado')
        if total_pqrs > 0:
            df_c = df_aap_raw['Canal'].value_counts().reset_index()
            df_c.columns = ['Canal', 'Cantidad']
            fig_c = px.pie(df_c, names='Canal', values='Cantidad', hole=0.4, color_discrete_sequence=PALETA_INTEGRAS)
            fig_c.update_traces(textinfo='percent+value')
            st.plotly_chart(fig_c, width='stretch')
        else:
            st.info("Sin registros de canales.")

    with p2:
        st.markdown('### Tipos de PQRS Recibidos')
        if total_pqrs > 0:
            df_t = df_aap_raw['Tipo_PQRS'].value_counts().reset_index()
            df_t.columns = ['Tipo_PQRS', 'Cantidad']
            fig_t = px.bar(df_t, y='Tipo_PQRS', x='Cantidad', orientation='h', color_discrete_sequence=[COLOR_AZUL_COOPI])
            st.plotly_chart(fig_t, width='stretch')
        else:
            st.info("Sin registros de tipos de PQRS.")

    st.markdown('---')
    st.markdown('### Indicadores de Satisfacción y Conocimiento AAP')
    if not df_eval_aap.empty:
        sat_col = [c for c in df_eval_aap.columns if 'satisfac' in c.lower()]
        if sat_col:
            df_sat = df_eval_aap[sat_col[0]].value_counts().reset_index()
            df_sat.columns = ['Nivel', 'Cantidad']
            fig_sat = px.pie(df_sat, names='Nivel', values='Cantidad', hole=0.4, color_discrete_sequence=PALETA_INTEGRAS, title='Nivel de Satisfacción')
            st.plotly_chart(fig_sat, width='stretch')
    else:
        st.info("Esperando registros del formulario de Indicadores AAP.")
