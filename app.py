
import io
import re
import json
import hashlib
import time
from datetime import datetime, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

st.set_page_config(page_title='UDI Connect', page_icon='🔎', layout='wide')
st.title('🔎 UDI Connect')
st.caption('EUDAMED · Búsqueda automática por SRN y referencia · Hasta 5.000 referencias')

API = 'https://api.datalake.sante.service.ec.europa.eu/eudamed/udi'
COLS = ['MANUFACT_ITEMNUM', 'SRN', 'UDI-DI', 'Basic UDI-DI', 'Master UDI-DI', 'Fabricante', 'Estado', 'Detalle', 'Fecha UTC']


def normalize(value):
    return re.sub(r'\s+', '', str(value or '')).upper().replace('–', '-').replace('—', '-')


def read_excel(upload):
    data = io.BytesIO(upload.getvalue())
    if upload.name.lower().endswith('.csv'):
        return pd.read_csv(data, dtype=str, keep_default_na=False, sep=None, engine='python')
    return pd.read_excel(data, dtype=str, keep_default_na=False)


def session():
    s = requests.Session()
    retry = Retry(total=2, connect=2, read=1, backoff_factor=0.6,
                  status_forcelist=[429, 500, 502, 503, 504], allowed_methods=['GET'])
    s.mount('https://', HTTPAdapter(max_retries=retry))
    s.headers.update({'Accept': 'application/json', 'User-Agent': 'UDIConnect/7.0'})
    return s


def fetch(srn, reference):
    if '-PR-' in srn:
        return {'Estado': 'Revisar PR', 'Detalle': 'Consulta de productor PR no implementada en este servicio MF'}
    with session() as s:
        response = s.get(API, params={'api-version': 'v1.0', 'format': 'json',
                                      'MF_SRN': srn, 'REFERENCE': reference}, timeout=(7, 25))
        response.raise_for_status()
        body = response.json()
    if not isinstance(body, dict) or not isinstance(body.get('value'), list):
        raise ValueError('La API devolvió un formato de datos inesperado')
    rows = body['value']
    if len(rows) >= 1000:
        return {'Estado': 'Revisar límite', 'Detalle': 'Respuesta de 1.000 registros: posible truncamiento'}
    exact = [r for r in rows if isinstance(r, dict) and
             normalize(r.get('MF_SRN')) == normalize(srn) and
             normalize(r.get('REFERENCE')) == normalize(reference)]
    if not exact:
        return {'Estado': 'No localizado en descarga', 'Detalle': 'Sin coincidencia exacta en la API; no confirma ausencia en EUDAMED'}
    unique = {(str(r.get('PRIMARY_DI') or ''), str(r.get('BASIC_UDI') or '')) for r in exact}
    if len(unique) != 1:
        return {'Estado': 'Revisar múltiples', 'Detalle': f'{len(unique)} combinaciones UDI para esta referencia'}
    primary, basic = next(iter(unique))
    if not primary:
        return {'Estado': 'Revisar sin UDI', 'Detalle': 'Coincidencia sin PRIMARY_DI'}
    return {'UDI-DI': primary, 'Basic UDI-DI': basic,
            'Fabricante': str(exact[0].get('MF_NAME') or ''), 'Estado': 'Encontrado',
            'Detalle': 'Coincidencia exacta de SRN y referencia'}


def excel_bytes(frame, fingerprint):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        frame.to_excel(writer, index=False, sheet_name='Resultados')
        pd.DataFrame([{'Huella': fingerprint, 'Exportado UTC': datetime.now(timezone.utc).isoformat()}]).to_excel(
            writer, index=False, sheet_name='Control')
        sheet = writer.sheets['Resultados']
        for cells in sheet.iter_rows(min_row=2):
            for cell in cells:
                cell.number_format = '@'
    return buffer.getvalue()


srn_default = st.text_input('SRN del fabricante', placeholder='LI-MF-000000522').strip().upper()
upload = st.file_uploader('Subir Excel con MANUFACT_ITEMNUM', type=['xlsx', 'csv'])
with st.expander('Opciones y recuperación de progreso'):
    batch_size = st.select_slider('Tamaño del lote', options=[10, 25, 50, 100], value=100)
    restore = st.file_uploader('Recuperar Excel de progreso anterior (opcional)', type=['xlsx'], key='restore')

if upload is None:
    st.info('Introduce un SRN, sube tu archivo y pulsa Iniciar búsqueda. También puedes incluir una columna SRN en el Excel.')
    st.stop()

try:
    source = read_excel(upload)
    ref_col = next((c for c in source.columns if str(c).strip().upper() == 'MANUFACT_ITEMNUM'), None)
    if ref_col is None:
        raise ValueError('El Excel debe incluir la columna MANUFACT_ITEMNUM')
    if not 1 <= len(source) <= 5000:
        raise ValueError('El Excel debe contener entre 1 y 5.000 filas')
    srn_col = next((c for c in source.columns if str(c).strip().upper() == 'SRN'), None)
    refs = [str(v).strip() for v in source[ref_col].tolist()]
    srns = [str(v).strip().upper() or srn_default for v in source[srn_col].tolist()] if srn_col else [srn_default] * len(refs)
    invalid = [i + 2 for i, v in enumerate(srns) if not re.fullmatch(r'[A-Z]{2}-(MF|PR)-\d{9}', v)]
    if invalid:
        raise ValueError(f'SRN inválido o vacío en filas: {invalid[:10]}')
except Exception as exc:
    st.error(f'Error de archivo: {exc}')
    st.stop()

fingerprint = hashlib.sha256(json.dumps([srns, refs], ensure_ascii=False).encode()).hexdigest()
if st.session_state.get('fingerprint') != fingerprint:
    st.session_state.fingerprint = fingerprint
    st.session_state.results = pd.DataFrame([
        {'MANUFACT_ITEMNUM': ref, 'SRN': srns[i], 'UDI-DI': '', 'Basic UDI-DI': '',
         'Master UDI-DI': '', 'Fabricante': '', 'Estado': 'Pendiente', 'Detalle': '', 'Fecha UTC': ''}
        for i, ref in enumerate(refs)], columns=COLS)
    st.session_state.cache = {}
    st.session_state.restored_hash = None

if restore is not None:
    restore_hash = hashlib.sha256(restore.getvalue()).hexdigest()
    if st.session_state.restored_hash != restore_hash:
        try:
            restored = pd.read_excel(io.BytesIO(restore.getvalue()), sheet_name='Resultados', dtype=str, keep_default_na=False)
            control = pd.read_excel(io.BytesIO(restore.getvalue()), sheet_name='Control', dtype=str, keep_default_na=False)
            if (control.empty or control.iloc[0].get('Huella') != fingerprint or
                    not set(COLS).issubset(restored.columns) or
                    restored['MANUFACT_ITEMNUM'].tolist() != refs or restored['SRN'].tolist() != srns):
                raise ValueError('El progreso no corresponde al Excel y SRN actuales')
            st.session_state.results = restored[COLS].copy()
            st.session_state.restored_hash = restore_hash
            st.success('Progreso recuperado')
        except Exception as exc:
            st.error(f'No se pudo recuperar el progreso: {exc}')

frame = st.session_state.results
processed = int((frame['Estado'] != 'Pendiente').sum())
found = int((frame['Estado'] == 'Encontrado').sum())
a, b, c = st.columns(3)
a.metric('Referencias', len(frame))
b.metric('Procesadas', processed)
c.metric('UDI-DI encontrados', found)
st.progress(processed / len(frame))

run = st.button('▶ Iniciar búsqueda / Continuar', type='primary', use_container_width=True)
if run:
    pending = [i for i in range(len(frame)) if frame.at[i, 'Estado'] in ('Pendiente', 'Error temporal')]
    if not pending:
        st.success('No quedan referencias pendientes.')
    else:
        progress = st.progress(processed / len(frame))
        status = st.empty()
        cache = st.session_state.cache
        fatal = None
        total_batches = (len(pending) + batch_size - 1) // batch_size
        for batch_number, start in enumerate(range(0, len(pending), batch_size), 1):
            indexes = pending[start:start + batch_size]
            groups = {}
            for i in indexes:
                ref = refs[i]
                if not ref:
                    frame.at[i, 'Estado'] = 'Referencia vacía'
                    frame.at[i, 'Detalle'] = 'Sin código en MANUFACT_ITEMNUM'
                    continue
                key = (srns[i], normalize(ref))
                groups.setdefault(key, []).append(i)
            todo = [key for key in groups if key not in cache]
            status.info(f'Lote {batch_number}/{total_batches} · {len(todo)} consultas nuevas')
            with ThreadPoolExecutor(max_workers=3) as executor:
                futures = {executor.submit(fetch, srn, ref): (srn, ref) for srn, ref in todo}
                for future in as_completed(futures):
                    key = futures[future]
                    try:
                        cache[key] = future.result()
                    except requests.HTTPError as exc:
                        code = exc.response.status_code if exc.response is not None else None
                        if code in (400, 401, 403, 404):
                            fatal = f'Error HTTP {code}: la conexión de EUDAMED no responde como se esperaba.'
                            cache[key] = {'Estado': 'Error temporal', 'Detalle': fatal}
                        else:
                            cache[key] = {'Estado': 'Error temporal', 'Detalle': str(exc)[:250]}
                    except Exception as exc:
                        cache[key] = {'Estado': 'Error temporal', 'Detalle': str(exc)[:250]}
            for key, positions in groups.items():
                item = cache.get(key, {'Estado': 'Error temporal', 'Detalle': 'Sin respuesta'})
                for i in positions:
                    for field, value in item.items():
                        frame.at[i, field] = value
                    frame.at[i, 'Fecha UTC'] = datetime.now(timezone.utc).isoformat()
            st.session_state.results = frame
            st.session_state.cache = cache
            done = int((frame['Estado'] != 'Pendiente').sum())
            progress.progress(done / len(frame))
            if fatal:
                st.error(f'{fatal} Se detuvo la búsqueda para evitar solicitudes repetidas.')
                break
            time.sleep(0.2)
        if not fatal:
            st.success('Búsqueda terminada. Descarga los resultados.')

st.download_button('⬇ Descargar resultados y progreso (Excel)',
                   data=excel_bytes(st.session_state.results, fingerprint),
                   file_name='UDI_Connect_resultados.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                   use_container_width=True)
st.dataframe(st.session_state.results, use_container_width=True, hide_index=True)
st.caption('Fuente: API pública EUDAMED. Master UDI-DI queda vacío si la fuente no lo proporciona. '
           'La descarga de progreso es necesaria para recuperar el trabajo tras cerrar la sesión. '
           'La conexión del servicio aún debe verificarse en el despliegue real.')
