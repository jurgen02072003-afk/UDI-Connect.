import io
import re
from datetime import datetime, timezone
import requests
import streamlit as st
from openpyxl import load_workbook

API = 'https://api.datalake.sante.service.ec.europa.eu/eudamed/udi'
SRN_DEFAULT = 'US-MF-000013697'
LIMIT = 1000

def norm(v):
    return re.sub(r'[^A-Z0-9]', '', str(v or '').upper())

def fetch(params):
    p = {'api-version': 'v1.0', 'format': 'json', **params}
    resp = requests.get(API, params=p, timeout=90, headers={'Accept': 'application/json'})
    resp.raise_for_status()
    data = resp.json()
    if not isinstance(data, dict) or not isinstance(data.get('value'), list):
        raise ValueError('Formato inesperado de la API. No se pueden validar los resultados.')
    return data['value']

def run_match(source, srn, code_col, dest_col, obs_col, first, last):
    wb = load_workbook(io.BytesIO(source))
    ws = wb.active
    codes = [(row, str(ws.cell(row, code_col).value or '').strip()) for row in range(first, min(last, ws.max_row) + 1)]
    # Do not treat a partial 1000-row response as the complete manufacturer dataset.
    rows = fetch({'MF_SRN': srn})
    complete = len(rows) < LIMIT
    by_ref = {}
    if complete:
        for item in rows:
            by_ref.setdefault(norm(item.get('REFERENCE')), []).append(item)
    counts = {'Encontrado': 0, 'Sin coincidencia': 0, 'Varias coincidencias': 0, 'Sin referencia': 0, 'Error de consulta': 0}
    date = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')
    for row, code in codes:
        if not code:
            status, di = 'Sin referencia', None
        else:
            try:
                candidates = by_ref.get(norm(code), []) if complete else fetch({'MF_SRN': srn, 'REFERENCE': code})
                # Strictly validate manufacturer and reference, even when API filters were supplied.
                candidates = [x for x in candidates if str(x.get('MF_SRN', '')).upper() == srn.upper() and norm(x.get('REFERENCE')) == norm(code)]
                ids = sorted({str(x.get('PRIMARY_DI')).strip() for x in candidates if x.get('PRIMARY_DI')})
                if len(ids) == 1:
                    status, di = 'Encontrado', ids[0]
                elif len(ids) > 1:
                    status, di = 'Varias coincidencias', None
                else:
                    status, di = 'Sin coincidencia', None
            except Exception as exc:
                status, di = 'Error de consulta', None
                st.session_state.setdefault('api_errors', []).append(f'Fila {row}: {exc}')
        counts[status] += 1
        if di:
            ws.cell(row, dest_col).value = di
            ws.cell(row, dest_col).number_format = '@'
        elif status != 'Error de consulta':
            ws.cell(row, dest_col).value = None
        ws.cell(row, obs_col).value = f'{status} | SRN {srn} | {date}'
        if status == 'Varias coincidencias':
            ws.cell(row, obs_col).value += f' | Revisar: {", ".join(ids[:8])}'
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), counts, complete

st.set_page_config(page_title='UDI Connect · EUDAMED', page_icon='🩺', layout='wide')
st.title('UDI Connect')
st.caption('Piloto de búsqueda masiva UDI-DI · Fuente: API pública EUDAMED (Comisión Europea)')
st.info('Los resultados son candidatos regulatorios que deben validarse antes de comunicar a AEMPS. «Sin coincidencia» no demuestra ausencia definitiva en EUDAMED.')
file = st.file_uploader('1. Sube el Excel (.xlsx)', type=['xlsx'])
col1, col2, col3 = st.columns(3)
with col1: srn = st.text_input('SRN del fabricante', SRN_DEFAULT).strip().upper()
with col2: code_letter = st.selectbox('Columna referencia fabricante', list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), index=2)
with col3: dest_letter = st.selectbox('Columna UDI-DI', list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), index=3)
obs_letter = st.selectbox('Columna observaciones', list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), index=6)
c1, c2 = st.columns(2)
with c1: first = st.number_input('Primera fila de datos', min_value=2, value=2)
with c2: last = st.number_input('Última fila a consultar', min_value=2, value=356)
if st.button('2. Consultar EUDAMED y preparar Excel', type='primary', disabled=file is None):
    if not re.fullmatch(r'[A-Z]{2}-[A-Z]{2,3}-\d{9}', srn):
        st.error('Formato SRN inválido.')
    elif len({code_letter, dest_letter, obs_letter}) != 3:
        st.error('Selecciona columnas distintas para referencia, UDI y observaciones.')
    elif last < first:
        st.error('La última fila debe ser mayor o igual que la primera.')
    else:
        try:
            with st.spinner('Consultando EUDAMED y cruzando referencias...'):
                result, counts, complete = run_match(file.getvalue(), srn, ord(code_letter)-65+1, ord(dest_letter)-65+1, ord(obs_letter)-65+1, first, last)
            st.session_state['result'] = result
            st.session_state['counts'] = counts
            st.session_state['complete'] = complete
        except Exception as exc:
            st.error(f'No se pudo completar la consulta. No se ha generado un Excel engañoso. Detalle: {exc}')
if 'result' in st.session_state:
    st.subheader('Resultados')
    cols = st.columns(5)
    for c, (name, n) in zip(cols, st.session_state['counts'].items()): c.metric(name, n)
    if not st.session_state['complete']:
        st.warning('El fabricante supera el límite de 1.000 registros de la API: se ha consultado cada referencia por separado.')
    if st.session_state.get('api_errors'):
        with st.expander('Errores de conexión por fila'):
            st.write('\n'.join(st.session_state['api_errors']))
    st.download_button('3. Descargar Excel actualizado', st.session_state['result'], file_name='UDI_CONNECT_RESULTADOS.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.markdown('Fuente: [API pública EUDAMED](https://webgate.ec.europa.eu/eudamed-help/en/files/Public%20API%20-%20user%20guide.pdf). No es una herramienta oficial de la Comisión Europea.')
