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
    except requests.exceptions.Timeout:
        error = 'Tiempo de espera agotado al contactar con EUDAMED, incluso después de los reintentos.'
    except requests.exceptions.RequestException as exc:
        error = f'Error de red o HTTP: {exc}'
    except Exception as exc:
        error = str(exc)
    progress.empty()

    if error:
        st.error(f'La consulta no se completó: {error}')
        st.info('No se han confirmado ausencias. Si se recibieron páginas anteriores, sus coincidencias se conservan; el resto queda pendiente. Prueba con 10 registros por página o más tarde.')
    elif not complete:
        st.warning('Consulta parcial: se alcanzó el límite de páginas. Los productos sin coincidencia permanecerán pendientes.')
    else:
        st.success('Se han consultado todas las páginas devueltas por el servicio.')

    st.write(f'Registros del SRN recibidos: **{len(records)}** · Páginas consultadas: **{pages_checked}**')
    if pages_checked > 0 and not records:
        st.warning('No hay registros con el campo manufacturerSrn coincidente en las páginas recibidas. Esto podría deberse a un campo SRN diferente en la respuesta. No se considera prueba de ausencia.')
        complete = False
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
