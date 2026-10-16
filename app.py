import io
import requests
from urllib.parse import urlparse
import pandas as pd
import streamlit as st

st.set_page_config(page_title='UDI Connect · Piloto', page_icon='🔎', layout='wide')
st.title('UDI Connect')
st.caption('Piloto para distribuidores · Carga, conciliación y exportación de referencias')
st.info('Consulta pública EUDAMED (producción): resultados sujetos a disponibilidad, paginación y verificación. Los códigos Master UDI-DI solo se muestran si la fuente los identifica expresamente.')

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

columnas = list(original.columns)

preferida = next(
    (c for c in columnas if c.strip().upper() == "MANUFACT_ITEMNUM"),
    columnas[0]
)

ref_col = st.selectbox(
    "Columna de referencia de tu Excel",
    columnas,
    index=columnas.index(preferida)
)

with col2:
    maker_col=st.selectbox('Columna de fabricante (opcional)', ['(Sin columna)']+list(original.columns))

st.subheader('Conexión pública EUDAMED')
api_mode = st.checkbox('Consultar EUDAMED directamente (API oficial)', value=False)
srn = ''
max_pages = 5
if api_mode:
    srn = st.text_input('SRN del fabricante (recomendado)', placeholder='US-MF-000013697').strip().upper()
    max_pages = st.number_input('Máximo de páginas (100 registros por página)', 1, 100, 5)
    st.caption('La consulta por SRN reduce las búsquedas. La API devuelve hasta 100 registros por página. No se declaran ausentes las referencias si la descarga está incompleta.')

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_eudamed(srn_value, pages):
    base='https://api.datalake.sante.service.ec.europa.eu/eudamed/udi'
    params={'api-version':'v1.0','format':'json'}
    if srn_value: params['MF_SRN']=srn_value
    rows=[]; next_url=base; complete=False
    with requests.Session() as session:
        for page in range(pages):
            response=session.get(next_url, params=params if page==0 else None, timeout=35, headers={'Accept':'application/json','User-Agent':'UDIConnectPilot/2.0'})
            response.raise_for_status()
            payload=response.json()
            if not isinstance(payload,dict) or not isinstance(payload.get('value'),list):
                raise ValueError('Respuesta no reconocida de EUDAMED; comprueba el endpoint en la documentación oficial.')
            rows.extend(payload['value'])
            nxt=payload.get('nextLink') or payload.get('@odata.nextLink')
            if not nxt:
                complete=True; break
            if urlparse(nxt).netloc != urlparse(base).netloc:
                raise ValueError('Enlace de paginación fuera del servidor oficial')
            next_url=nxt
    return rows, complete

def source_col(frame, *names):
    lookup={str(k).upper().replace('_','').replace(' ',''):k for k in frame.columns}
    for name in names:
        match=lookup.get(name.upper().replace('_','').replace(' ',''))
        if match is not None: return match
    return None

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
    if api_mode:
        try:
            with st.spinner('Consultando la API pública de EUDAMED...'):
                rows, complete = fetch_eudamed(srn, int(max_pages))
            live=pd.DataFrame(rows)
            if live.empty:
                st.warning('La API no ha devuelto registros. Verifica el SRN o el acceso al servicio.')
            else:
                st.write(f'Registros recuperados: {len(live)}. Descarga completa: {"sí" if complete else "no"}.')
                st.write('Campos devueltos:', ', '.join(map(str,live.columns)))
                r=source_col(live,'REFERENCE','DEVICE_REFERENCE','MANUFACTURER_ITEM_NUMBER','CATALOGUE_NUMBER')
                u=source_col(live,'PRIMARY_DI','PRIMARY_UDI','UDI_DI')
                b=source_col(live,'BASIC_UDI','BASIC_UDI_DI')
                m=source_col(live,'MASTER_UDI_DI')
                if r is None or u is None:
                    st.error('La respuesta no incluye campos de referencia y UDI-DI reconocibles. No se realizará un cruce inseguro.')
                else:
                    live['_key']=live[r].map(normalize)
                    live=live[live['_key']!='']
                    groups=live.groupby('_key',dropna=False)
                    for idx,ref in result['Referencia consultada'].items():
                        key=normalize(ref)
                        if not key:
                            result.at[idx,'Estado']='Referencia vacía'
                        elif key not in groups.groups:
                            result.at[idx,'Estado']='No encontrada (consulta completa)' if complete else 'Pendiente: consulta parcial'
                        else:
                            matches=groups.get_group(key)
                            if len(matches)!=1:
                                result.at[idx,'Estado']='Coincidencia múltiple: revisar'
                            else:
                                entry=matches.iloc[0]
                                result.at[idx,'UDI-DI']=str(entry[u]) if pd.notna(entry[u]) else ''
                                if b: result.at[idx,'Basic UDI-DI']=str(entry[b]) if pd.notna(entry[b]) else ''
                                if m: result.at[idx,'Master UDI-DI']=str(entry[m]) if pd.notna(entry[m]) else ''
                                result.at[idx,'Estado']='Coincidencia exacta EUDAMED (verificar fabricante)'
        except Exception as exc:
            st.error(f'No se pudo completar la consulta EUDAMED: {exc}')
            st.warning('No se han inventado resultados. Puedes seguir usando el cruce con un catálogo descargado.')
    elif catalog is not None:
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
