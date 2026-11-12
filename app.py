import hashlib
import io
import json
import re
import time
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

st.set_page_config(page_title='UDI Connect 6.0', page_icon='🔎', layout='wide')
st.title('UDI Connect 6.0')
st.caption('Búsqueda piloto de UDI-DI · 5.000 referencias · progreso recuperable')

API = 'https://api.datalake.sante.service.ec.europa.eu/eudamed/udi'
MAX_ROWS = 5000
COLS = ['MANUFACT_ITEMNUM', 'SRN', 'UDI-DI', 'Basic UDI-DI', 'Master UDI-DI', 'Estado', 'Detalle', 'Fecha UTC']


def text(value):
    if value is None or pd.isna(value):
        return ''
    return str(value).strip()


def normalized(value):
    return re.sub(r'\s+', '', text(value)).upper().replace('–', '-').replace('—', '-')


def read_source(upload):
    data = io.BytesIO(upload.getvalue())
    if upload.name.lower().endswith('.csv'):
        for encoding in ('utf-8-sig', 'latin-1'):
            try:
                data.seek(0)
                return pd.read_csv(data, dtype=str, keep_default_na=False, sep=None, engine='python', encoding=encoding)
            except (UnicodeError, pd.errors.ParserError):
                pass
        raise ValueError('No se ha podido interpretar el CSV.')
    return pd.read_excel(data, dtype=str, keep_default_na=False)


def signature(srn, references):
    content = json.dumps([srn, references], ensure_ascii=False).encode('utf-8')
    return hashlib.sha256(content).hexdigest()


def empty_results(srn, references):
    return pd.DataFrame([{
        'MANUFACT_ITEMNUM': ref, 'SRN': srn, 'UDI-DI': '',
        'Basic UDI-DI': '', 'Master UDI-DI': '',
        'Estado': 'Pendiente', 'Detalle': '', 'Fecha UTC': ''
    } for ref in references], columns=COLS)


def field(record, *names):
    lookup = {re.sub(r'[^a-z0-9]', '', str(k).lower()): v for k, v in record.items()}
    for name in names:
        key = re.sub(r'[^a-z0-9]', '', name.lower())
        value = lookup.get(key)
        if value is not None and text(value):
            return text(value)
    return ''


def make_session():
    s = requests.Session()
    retry = Retry(total=1, connect=1, read=0, status=1, backoff_factor=0.5,
                  status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=['GET'], respect_retry_after_header=True)
    s.mount('https://', HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4))
    s.headers.update({'Accept': 'application/json', 'User-Agent': 'UDIConnectPilot/6.0'})
    return s


def lookup_device(session, srn, reference):
    # IMPORTANTE: esta API y su esquema requieren validación real con EUDAMED.
    # No enviar SRN PR al filtro MF_SRN: no son roles equivalentes.
    if '-PR-' in srn:
        return {'Estado': 'Consulta PR no disponible',
                'Detalle': 'El servicio configurado solo dispone de una consulta piloto por MF_SRN. Verificar PR en EUDAMED.'}

    try:
        response = session.get(API, params={
            'api-version': 'v1.0', 'format': 'json',
            'MF_SRN': srn, 'REFERENCE': reference
        }, timeout=(5, 12))
        response.raise_for_status()
        payload = response.json()
    except requests.exceptions.Timeout:
        return {'Estado': 'Error de conexión', 'Detalle': 'Tiempo de espera agotado; se puede reintentar.'}
    except requests.exceptions.RequestException as exc:
        return {'Estado': 'Error de conexión', 'Detalle': str(exc)[:220]}
    except ValueError:
        return {'Estado': 'Error de formato', 'Detalle': 'La respuesta no es JSON válido.'}

    if not isinstance(payload, dict) or not isinstance(payload.get('value'), list):
        return {'Estado': 'Error de formato', 'Detalle': 'La API no devolvió la lista value esperada; comprobar documentación del endpoint.'}

    if payload.get('nextLink') or payload.get('@odata.nextLink'):
        return {'Estado': 'Consulta incompleta', 'Detalle': 'La API indica páginas adicionales; no se puede asegurar coincidencia exhaustiva.'}

    exact = []
    for record in payload['value']:
        if not isinstance(record, dict):
            continue
        returned_srn = field(record, 'MF_SRN', 'manufacturerSrn')
        returned_ref = field(record, 'REFERENCE', 'manufacturerItemNumber', 'catalogueNumber')
        if normalized(returned_srn) == normalized(srn) and normalized(returned_ref) == normalized(reference):
            exact.append(record)

    if not exact:
        return {'Estado': 'Sin coincidencia verificable',
                'Detalle': 'No se encontró coincidencia exacta en esta respuesta; no implica ausencia en EUDAMED.'}

    identifiers = set()
    for record in exact:
        identifiers.add((field(record, 'PRIMARY_DI', 'primaryDi', 'UDI_DI'),
                         field(record, 'BASIC_UDI', 'basicUdi'),
                         field(record, 'MASTER_UDI_DI', 'masterUdiDi')))
    if len(identifiers) != 1:
        return {'Estado': 'Coincidencias múltiples', 'Detalle': 'Distintos identificadores para el mismo SRN y referencia; revisión manual.'}

    primary, basic, master = next(iter(identifiers))
    if not primary:
        return {'Estado': 'UDI-DI no identificado', 'Detalle': 'Coincidencia exacta, pero sin UDI-DI primario reconocible.'}
    return {'Estado': 'Encontrado', 'UDI-DI': primary, 'Basic UDI-DI': basic,
            'Master UDI-DI': master, 'Detalle': 'Coincidencia exacta en la respuesta de la API piloto.'}


def export_excel(results, srn, key):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        results.to_excel(writer, sheet_name='Resultados', index=False)
        pd.DataFrame([{'SRN': srn, 'Fingerprint': key,
                       'Fecha UTC': datetime.now(timezone.utc).isoformat(),
                       'Fuente': API, 'Version': '6.0'}]).to_excel(writer, sheet_name='Control', index=False)
        for sheet in writer.book.worksheets:
            for cells in sheet.iter_rows(min_row=2):
                for cell in cells:
                    cell.number_format = '@'
    return buffer.getvalue()


st.info('La API de descarga de EUDAMED no está validada para todas las referencias. Los resultados son orientativos y deben contrastarse antes de RECOPS.')
st.markdown('[Abrir búsqueda pública de EUDAMED](https://ec.europa.eu/tools/eudamed/#/screen/search-device)')

srn = st.text_input('SRN del actor', placeholder='ES-MF-000002991 o ES-PR-000042695').strip().upper()
source_file = st.file_uploader('Archivo Excel o CSV con MANUFACT_ITEMNUM', type=['xlsx', 'csv'], key='source')
checkpoint_file = st.file_uploader('Recuperar Excel de progreso (opcional)', type=['xlsx'], key='checkpoint')
batch_size = st.select_slider('Referencias por lote', options=[10, 25, 50, 100], value=25)

if not re.fullmatch(r'[A-Z]{2}-(MF|PR)-\d{9}', srn):
    st.info('Introduce un SRN válido de tipo MF o PR (por ejemplo, ES-MF-000002991).')
    st.stop()
if source_file is None:
    st.info('Carga el Excel de referencias para comenzar.')
    st.stop()

try:
    source = read_source(source_file)
    ref_col = next((col for col in source.columns if str(col).strip().upper() == 'MANUFACT_ITEMNUM'), None)
    if ref_col is None:
        raise ValueError('No existe la columna MANUFACT_ITEMNUM.')
    if not 1 <= len(source) <= MAX_ROWS:
        raise ValueError(f'El archivo debe contener entre 1 y {MAX_ROWS} filas.')
    references = [text(x) for x in source[ref_col].tolist()]
    key = signature(srn, references)
except Exception as exc:
    st.error(f'Error al cargar referencias: {exc}')
    st.stop()

if st.session_state.get('active_key') != key:
    st.session_state['active_key'] = key
    st.session_state['results'] = empty_results(srn, references)
    st.session_state['checkpoint_digest'] = None

if checkpoint_file is not None:
    digest = hashlib.sha256(checkpoint_file.getvalue()).hexdigest()
    if digest != st.session_state.get('checkpoint_digest'):
        try:
            previous = pd.read_excel(io.BytesIO(checkpoint_file.getvalue()), sheet_name='Resultados', dtype=str, keep_default_na=False)
            control = pd.read_excel(io.BytesIO(checkpoint_file.getvalue()), sheet_name='Control', dtype=str, keep_default_na=False)
            if control.empty or text(control.iloc[0].get('Fingerprint', '')) != key:
                raise ValueError('El archivo de progreso corresponde a otro SRN u otra lista de referencias.')
            if not set(COLS).issubset(previous.columns):
                raise ValueError('El archivo de progreso no tiene todas las columnas necesarias.')
            if len(previous) != len(references) or previous['MANUFACT_ITEMNUM'].astype(str).tolist() != references:
                raise ValueError('Las referencias del progreso no coinciden con el Excel actual.')
            st.session_state['results'] = previous[COLS].copy().reset_index(drop=True)
            st.session_state['checkpoint_digest'] = digest
            st.success('Progreso recuperado correctamente.')
        except Exception as exc:
            st.error(f'No se pudo recuperar el progreso: {exc}')

results = st.session_state['results']
retry = st.checkbox('Reintentar errores y referencias sin coincidencia', value=False)
retryable = ['Error de conexión', 'Error de formato', 'Consulta incompleta',
             'Sin coincidencia verificable', 'UDI-DI no identificado']
statuses = ['Pendiente'] + (retryable if retry else [])
pending = results.index[results['Estado'].isin(statuses)].tolist()
found = int(results['Estado'].eq('Encontrado').sum())

c1, c2, c3 = st.columns(3)
c1.metric('Referencias', len(results))
c2.metric('UDI-DI encontrados', found)
c3.metric('Pendientes', int(results['Estado'].eq('Pendiente').sum()))
st.progress((len(results) - int(results['Estado'].eq('Pendiente').sum())) / len(results))

if '-PR-' in srn:
    st.warning('El SRN PR es válido, pero esta versión no dispone de un endpoint verificado para buscar dispositivos por productor PR. No se ejecutarán peticiones MF_SRN con este SRN.')

if st.button('Procesar siguiente lote', type='primary', disabled=not pending or '-PR-' in srn):
    selected = pending[:batch_size]
    bar = st.progress(0)
    status = st.empty()
    with make_session() as session:
        for n, idx in enumerate(selected, start=1):
            ref = references[idx]
            status.write(f'Consultando {n}/{len(selected)}: {ref}')
            update = ({'Estado': 'Referencia vacía', 'Detalle': 'Sin referencia de fabricante.'}
                      if not ref else lookup_device(session, srn, ref))
            # Al reintentar, borrar identificadores antiguos antes de actualizar.
            for col in ('UDI-DI', 'Basic UDI-DI', 'Master UDI-DI'):
                results.at[idx, col] = ''
            for col, value in update.items():
                results.at[idx, col] = value
            results.at[idx, 'Fecha UTC'] = datetime.now(timezone.utc).isoformat()
            st.session_state['results'] = results
            bar.progress(n / len(selected))
            if n < len(selected):
                time.sleep(0.1)
    status.empty()
    st.success(f'Lote terminado ({len(selected)} referencias). Descarga ahora el Excel de progreso.')

st.subheader('Resultados')
st.dataframe(st.session_state['results'], use_container_width=True, hide_index=True)

st.download_button('Descargar resultados y progreso (Excel)',
                   data=export_excel(st.session_state['results'], srn, key),
                   file_name=f'UDI_Connect_{srn}_progreso.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.caption('El progreso se mantiene durante la sesión. Para recuperarlo después de un reinicio, descarga y vuelve a subir el Excel de progreso. El almacenamiento automático permanente requiere un servicio externo.')
