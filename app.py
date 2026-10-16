import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title='UDI Connect · Piloto', page_icon='🔎', layout='wide')
st.title('UDI Connect')
st.caption('Piloto para distribuidores · Carga, conciliación y exportación de referencias')
st.info('Modo demostración: no existe todavía conexión automática a EUDAMED. Puedes importar un catálogo público exportado de EUDAMED para cruzar referencias sin inventar resultados.')

with st.sidebar:
    st.header('Configuración')
    st.write('1. Carga tu listado de referencias. 2. Opcional: carga un catálogo de EUDAMED. 3. Asocia columnas y ejecuta el cruce.')
    st.markdown('[Abrir EUDAMED](https://ec.europa.eu/tools/eudamed/)')

def load(upload):
    if upload is None: return None
    if upload.name.lower().endswith('.csv'):
        raw=upload.getvalue()
        for enc in ('utf-8-sig','latin-1'):
            try: return pd.read_csv(io.BytesIO(raw), dtype=str, sep=None, engine='python', encoding=enc).fillna('')
            except Exception: pass
        raise ValueError('No se pudo leer el CSV')
    return pd.read_excel(upload, dtype=str).fillna('')

def normalize(s):
    return str(s).strip().upper().replace(' ', '').replace('–','-').replace('—','-')

left,right=st.columns(2)
with left:
    user_file=st.file_uploader('Listado de referencias (.xlsx o .csv)',type=['xlsx','csv'],key='user')
with right:
    eud_file=st.file_uploader('Catálogo de EUDAMED (.xlsx o .csv) · opcional',type=['xlsx','csv'],key='eud')

try:
    original=load(user_file)
    catalog=load(eud_file)
except Exception as exc:
    st.error(f'Error al abrir el archivo: {exc}')
    st.stop()

if original is None:
    st.write('Carga un archivo para comenzar. Puedes probar con el Excel de ejemplo incluido en el paquete.')
    st.stop()

st.subheader('Asociar columnas')
col1,col2=st.columns(2)
with col1:
    ref_col=st.selectbox('Columna de referencia de tu Excel',list(original.columns),index=next((i for i,c in enumerate(original.columns) if 'ref' in c.lower() or 'item' in c.lower()),0))
with col2:
    maker_col=st.selectbox('Columna de fabricante (opcional)', ['(Sin columna)']+list(original.columns))

if catalog is not None:
    a,b,c=st.columns(3)
    with a: cat_ref=st.selectbox('Referencia en catálogo EUDAMED',list(catalog.columns))
    with b: cat_udi=st.selectbox('UDI-DI en catálogo',list(catalog.columns))
    with c: cat_basic=st.selectbox('Basic UDI-DI (opcional)',['(No disponible)']+list(catalog.columns))
    cat_master=st.selectbox('Master UDI-DI (solo si el catálogo contiene ese dato)',['(No disponible)']+list(catalog.columns))

if st.button('Procesar referencias',type='primary'):
    result=original.copy()
    result['Referencia consultada']=original[ref_col].astype(str)
    result['Fabricante consultado']=original[maker_col].astype(str) if maker_col!='(Sin columna)' else ''
    for field in ['UDI-DI','Basic UDI-DI','Master UDI-DI']:
        result[field]=''
    result['Estado']='Pendiente de consulta EUDAMED'
    if catalog is not None:
        cat=catalog.copy()
        cat['_key']=cat[cat_ref].map(normalize)
        counts=cat.groupby('_key').size().to_dict()
        unique=cat.drop_duplicates('_key').set_index('_key')
        for idx,ref in result['Referencia consultada'].items():
            key=normalize(ref)
            if not key:
                result.at[idx,'Estado']='Referencia vacía'
            elif counts.get(key,0)==0:
                result.at[idx,'Estado']='No encontrada en catálogo cargado'
            elif counts[key]>1:
                result.at[idx,'Estado']='Coincidencia múltiple: revisar'
            else:
                entry=unique.loc[key]
                result.at[idx,'UDI-DI']=entry[cat_udi]
                if cat_basic!='(No disponible)': result.at[idx,'Basic UDI-DI']=entry[cat_basic]
                if cat_master!='(No disponible)': result.at[idx,'Master UDI-DI']=entry[cat_master]
                result.at[idx,'Estado']='Coincidencia exacta en catálogo cargado'
    st.session_state['results']=result

if 'results' in st.session_state:
    result=st.session_state['results']
    x,y,z=st.columns(3)
    x.metric('Referencias',len(result))
    y.metric('Coincidencias exactas',int((result['Estado']=='Coincidencia exacta en catálogo cargado').sum()))
    z.metric('Por revisar',int((result['Estado']!='Coincidencia exacta en catálogo cargado').sum()))
    st.dataframe(result,use_container_width=True,hide_index=True)
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as writer:
        result.to_excel(writer,index=False,sheet_name='Resultados')
    st.download_button('Descargar resultados Excel',out.getvalue(),file_name='UDI_Connect_resultados.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.divider()
st.caption('Los resultados son coincidencias de texto contra el catálogo aportado, no una verificación regulatoria. Validar siempre el fabricante, el dispositivo y el identificador antes de comunicar a AEMPS.')
