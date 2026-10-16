import io
import re
import time
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title='UDI Connect', page_icon='🔎', layout='wide')
st.title('UDI Connect')
st.caption('Consulta de referencias de fabricantes en el registro público de EUDAMED')
st.info('La aplicación muestra UDI-DI y Basic UDI-DI como identificadores diferentes. Master UDI-DI solo se informará cuando exista un dato explícito y verificable.')

BASE = 'https://ec.europa.eu/tools/eudamed/api/devices/udiDiData'
PUBLIC = 'https://ec.europa.eu/tools/eudamed/#/screen/search-device'

def norm(value):
    return re.sub(r'\s+', '', str(value or '')).upper().replace('–', '-').replace('—', '-')

def read_file(upload):
    data = upload.getvalue()
    if upload.name.lower().endswith('.csv'):
        for encoding in ('utf-8-sig', 'latin-1'):
            try:
                return pd.read_csv(io.BytesIO(data), dtype=str, sep=None, engine='python', encoding=encoding).fillna('')
            except (UnicodeError, pd.errors.ParserError):
                continue
        raise ValueError('No se pudo interpretar el archivo CSV.')
    return pd.read_excel(io.BytesIO(data), dtype=str).fillna('')

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_page(srn, page, page_size):
    params = {
        'page': page,
        'pageSize': page_size,
        'size': page_size,
        'iso2Code': 'en',
        'languageIso2Code': 'en',
        'srn': srn,
    }
    response = requests.get(BASE, params=params, timeout=45, headers={
        'Accept': 'application/json',
        'User-Agent': 'UDIConnectPilot/3.0',
    })
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict) or not isinstance(payload.get('content'), list):
        raise ValueError('EUDAMED ha devuelto un formato inesperado; no se pueden verificar las coincidencias.')
    return payload

with st.sidebar:
    st.header('Fuente de datos')
    st.markdown(f'[Abrir buscador oficial de EUDAMED]({PUBLIC})')
    st.caption('Se consulta el servicio público utilizado por la pantalla Devices/SPPs. La disponibilidad y el formato pueden cambiar.')

srn = st.text_input('SRN del fabricante', value='US-MF-000013697', help='Puedes introducir el SRN de cualquier fabricante, por ejemplo US-MF-000013697.').strip().upper()
file = st.file_uploader('Sube tu Excel o CSV con MANUFACT_ITEMNUM', type=['xlsx', 'csv'])
page_size = st.selectbox('Registros por página', [25, 50, 100, 300], index=2)
max_pages = st.number_input('Máximo de páginas que se consultarán', min_value=1, max_value=10000, value=20, step=1)
st.caption('Si se alcanza el límite de páginas, los productos no localizados quedan pendientes, nunca se marcan como ausentes.')

if file is not None:
    try:
        original = read_file(file)
        col = next((c for c in original.columns if str(c).strip().upper() == 'MANUFACT_ITEMNUM'), None)
        if col is None:
            st.error('El archivo debe contener una columna llamada MANUFACT_ITEMNUM.')
            st.stop()
        refs = original[col].astype(str).tolist()
        st.success(f'Archivo leído: {len(refs)} filas. Columna identificada: {col}.')
        st.dataframe(pd.DataFrame({'MANUFACT_ITEMNUM': refs}).head(10), hide_index=True, use_container_width=True)
    except Exception as exc:
        st.error(f'No se pudo abrir el archivo: {exc}')
        st.stop()
else:
    refs = None

if st.button('Buscar UDI-DI en EUDAMED', type='primary', disabled=(refs is None or not srn)):
    if not re.fullmatch(r'[A-Z]{2}-MF-\d+', srn):
        st.error('Revisa el SRN: debe tener el formato de fabricante, por ejemplo US-MF-000013697.')
        st.stop()

    progress = st.progress(0, text='Iniciando consulta...')
    records = []
    complete = False
    error = None
    pages_checked = 0
    total_pages = None
    try:
        for page in range(int(max_pages)):
            payload = fetch_page(srn, page, int(page_size))
            chunk = payload['content']
            # Verificar el SRN en cada registro: el servidor podría ignorar filtros desconocidos.
            for item in chunk:
                if not isinstance(item, dict):
                    continue
                record_srn = str(item.get('manufacturerSrn') or '').strip().upper()
                if record_srn == srn:
                    records.append(item)
            pages_checked += 1
            if total_pages is None:
                try:
                    total_pages = int(payload['totalPages'])
                except (KeyError, ValueError, TypeError):
                    total_pages = None
            progress.progress(min((page + 1) / int(max_pages), 1.0), text=f'Páginas consultadas: {pages_checked}' + (f' de {total_pages}' if total_pages is not None else ''))
            if payload.get('last') is True or (total_pages is not None and page + 1 >= total_pages) or not chunk or len(chunk) < int(page_size):
                complete = True
                break
    except Exception as exc:
        error = str(exc)
    progress.empty()

    if error:
        st.error(f'La consulta no se completó: {error}')
    elif not complete:
        st.warning('Consulta parcial: se alcanzó el límite de páginas. Los productos sin coincidencia permanecerán pendientes.')
    else:
        st.success('Se han consultado todas las páginas devueltas por el servicio.')

    st.write(f'Registros del SRN recibidos: **{len(records)}** · Páginas consultadas: **{pages_checked}**')
    lookup = {}
    for item in records:
        reference = norm(item.get('reference'))
        if reference:
            lookup.setdefault(reference, []).append(item)

    result_rows = []
    for ref in refs:
        key = norm(ref)
        matches = lookup.get(key, []) if key else []
        row = {
            'MANUFACT_ITEMNUM': ref,
            'SRN': srn,
            'UDI-DI': '',
            'Basic UDI-DI': '',
            'Master UDI-DI': '',
            'Fabricante EUDAMED': '',
            'Referencia EUDAMED': '',
            'Estado': '',
            'Enlace EUDAMED': '',
        }
        if not key:
            row['Estado'] = 'Referencia vacía'
        elif not matches:
            row['Estado'] = 'No encontrada en la consulta completa' if complete and not error else 'Pendiente: consulta incompleta'
        else:
            # Versiones históricas del mismo dispositivo no deben confundirse con referencias diferentes.
            latest = [x for x in matches if x.get('latestVersion') is True]
            candidates = latest or matches
            distinct = {(str(x.get('primaryDi') or ''), str(x.get('uuid') or '')) for x in candidates}
            if len(distinct) != 1:
                row['Estado'] = 'Coincidencias múltiples: revisar'
            else:
                item = candidates[0]
                row['UDI-DI'] = str(item.get('primaryDi') or '')
                row['Basic UDI-DI'] = str(item.get('basicUdi') or '')
                row['Fabricante EUDAMED'] = str(item.get('manufacturerName') or '')
                row['Referencia EUDAMED'] = str(item.get('reference') or '')
                uuid = str(item.get('uuid') or '')
                if uuid:
                    row['Enlace EUDAMED'] = f'{PUBLIC}/{uuid}'
                row['Estado'] = 'Coincidencia exacta por referencia y SRN' if row['UDI-DI'] else 'Coincidencia sin UDI-DI: revisar'
        result_rows.append(row)

    result = pd.DataFrame(result_rows)
    st.session_state['udi_results'] = result
    st.session_state['udi_query'] = {'srn': srn, 'complete': complete, 'error': error, 'pages': pages_checked, 'records': len(records)}

if 'udi_results' in st.session_state:
    result = st.session_state['udi_results']
    info = st.session_state.get('udi_query', {})
    st.subheader('Resultados')
    a, b, c = st.columns(3)
    a.metric('Referencias', len(result))
    b.metric('Encontradas', int(result['Estado'].eq('Coincidencia exacta por referencia y SRN').sum()))
    c.metric('Pendientes o por revisar', int((result['Estado'] != 'Coincidencia exacta por referencia y SRN').sum()))
    st.dataframe(result, hide_index=True, use_container_width=True)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='Resultados', index=False)
        pd.DataFrame([info]).to_excel(writer, sheet_name='Control de consulta', index=False)
        sheet = writer.sheets['Resultados']
        for cells in sheet.columns:
            for cell in cells[1:]:
                cell.number_format = '@'
    st.download_button('Descargar Excel de resultados', data=out.getvalue(), file_name='UDI_Connect_resultados.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.divider()
st.caption('Herramienta piloto no afiliada a la Comisión Europea. Verifica los identificadores y datos regulatorios en EUDAMED antes de comunicarlos a AEMPS.')
