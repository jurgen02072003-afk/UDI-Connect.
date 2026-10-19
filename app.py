
import io
import json
import re
import time
import hashlib
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


st.set_page_config(
    page_title="UDI Connect 5.0",
    page_icon="🔎",
    layout="wide"
)

st.title("UDI Connect 5.0")

st.caption(
    "Consulta por SRN y MANUFACT_ITEMNUM "
    "· hasta 5.000 referencias · lotes de 100"
)

st.warning(
    "Piloto: la API pública de descarga puede "
    "diferir de la búsqueda web de EUDAMED. "
    "Un resultado vacío no demuestra que "
    "el producto no esté registrado."
)


API = (
    "https://api.datalake.sante.service.ec.europa.eu"
    "/eudamed/udi"
)

PUBLIC = (
    "https://ec.europa.eu/tools/eudamed/"
    "#/screen/search-device"
)

MAX_ROWS = 5000
BATCH_SIZE = 100

RESULT_COLS = [
    "MANUFACT_ITEMNUM",
    "SRN",
    "UDI-DI",
    "Basic UDI-DI",
    "Master UDI-DI",
    "Estado",
    "Detalle"
]


def clean(value):
    if value is None:
        return ""

    return re.sub(
        r"\s+", "", str(value)
    ).upper().replace(
        "–", "-"
    ).replace("—", "-")


def read_excel_or_csv(upload):
    raw = upload.getvalue()

    if upload.name.lower().endswith(".csv"):
        for encoding in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(
                    io.BytesIO(raw),
                    dtype=str,
                    sep=None,
                    engine="python",
                    encoding=encoding,
                    keep_default_na=False
                )
            except (
                UnicodeError,
                pd.errors.ParserError
            ):
                pass

        raise ValueError("CSV ilegible")

    return pd.read_excel(
        io.BytesIO(raw),
        dtype=str,
        keep_default_na=False
    )


def value_of(record, *keys):
    normalized = {
        str(k).replace(
            "_", ""
        ).replace(
            "-", ""
        ).lower(): v
        for k, v in record.items()
    }

    for key in keys:
        normalized_key = key.replace(
            "_", ""
        ).replace(
            "-", ""
        ).lower()

        value = normalized.get(normalized_key)

        if value is not None and str(value).strip():
            return str(value).strip()

    return ""


def new_session(srn, references):
    rows = []

    for ref in references:
        rows.append({
            "MANUFACT_ITEMNUM": ref,
            "SRN": srn,
            "UDI-DI": "",
            "Basic UDI-DI": "",
            "Master UDI-DI": "",
            "Estado": "Pendiente",
            "Detalle": ""
        })

    return pd.DataFrame(
        rows,
        columns=RESULT_COLS
    )


def fingerprint(srn, refs):
    payload = json.dumps(
        [srn, refs],
        ensure_ascii=False
    ).encode("utf-8")

    return hashlib.sha256(
        payload
    ).hexdigest()


def api_lookup(srn, ref):
    retry = Retry(
        total=2,
        connect=2,
        read=2,
        status=2,
        backoff_factor=1.0,
        status_forcelist=[
            429, 500, 502, 503, 504
        ],
        allowed_methods=["GET"],
        respect_retry_after_header=True
    )

    with requests.Session() as session:
        session.mount(
            "https://",
            HTTPAdapter(max_retries=retry)
        )

        response = session.get(
            API,
            params={
                "api-version": "v1.0",
                "format": "json",
                "MF_SRN": srn,
                "REFERENCE": ref
            },
            timeout=(10, 35),
            headers={
                "Accept": "application/json",
                "User-Agent": "UDIConnectPilot/5.0"
            }
        )

        response.raise_for_status()
        payload = response.json()

    if (
        not isinstance(payload, dict)
        or not isinstance(
            payload.get("value"), list
        )
    ):
        raise ValueError(
            "Formato de respuesta no reconocido. "
            "Verificar endpoint y esquema de la API."
        )

    more = bool(
        payload.get("nextLink")
        or payload.get("@odata.nextLink")
    )

    records = payload["value"]
    exact = []

    for record in records:
        if not isinstance(record, dict):
            continue

        returned_srn = value_of(
            record,
            "MF_SRN",
            "manufacturerSrn"
        )

        returned_ref = value_of(
            record,
            "REFERENCE",
            "manufacturerItemNumber",
            "catalogueNumber"
        )

        if (
            clean(returned_srn) == clean(srn)
            and clean(returned_ref) == clean(ref)
        ):
            exact.append(record)

    if not exact:
        detail = (
            "Sin coincidencia exacta SRN + "
            "referencia en esta respuesta"
        )

        if more:
            detail += "; hay más páginas"

        return {
            "Estado": "Sin coincidencia verificable",
            "Detalle": detail
        }

    distinct = {
        (
            value_of(
                x, "PRIMARY_DI",
                "primaryDi", "UDI_DI"
            ),
            value_of(
                x, "BASIC_UDI",
                "basicUdi"
            ),
            value_of(
                x, "MASTER_UDI_DI",
                "masterUdiDi"
            )
        )
        for x in exact
    }

    if len(distinct) != 1 or more:
        if len(distinct) != 1:
            state = (
                "Coincidencias múltiples: revisar"
            )
        else:
            state = "Consulta parcial: revisar"

        return {
            "Estado": state,
            "Detalle": (
                "Varias combinaciones de "
                "identificadores o paginación pendiente"
            )
        }

    primary, basic, master = next(
        iter(distinct)
    )

    if not primary:
        return {
            "Estado": (
                "Encontrado sin UDI-DI: revisar"
            ),
            "Detalle": (
                "Registro coincidente sin "
                "identificador primario reconocido"
            )
        }

    return {
        "Estado": "Encontrado",
        "UDI-DI": primary,
        "Basic UDI-DI": basic,
        "Master UDI-DI": master,
        "Detalle": (
            "Coincidencia exacta SRN + "
            "referencia en respuesta API"
        )
    }


def excel_bytes(frame, srn, key):
    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ) as writer:

        frame.to_excel(
            writer,
            index=False,
            sheet_name="Resultados"
        )

        control = pd.DataFrame([{
            "SRN": srn,
            "Fingerprint": key,
            "Fecha_UTC": datetime.now(
                timezone.utc
            ).isoformat(),
            "Fuente": API
        }])

        control.to_excel(
            writer,
            index=False,
            sheet_name="Control"
        )

        sheet = writer.sheets["Resultados"]

        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.number_format = "@"

    return out.getvalue()


st.markdown(
    f"[Abrir búsqueda pública de EUDAMED]({PUBLIC})"
)

srn = st.text_input(
    "SRN del fabricante",
    placeholder="ES-MF-000002991"
).strip().upper()

file = st.file_uploader(
    "Excel con columna MANUFACT_ITEMNUM",
    type=["xlsx", "csv"],
    key="source"
)

checkpoint = st.file_uploader(
    "Archivo de progreso anterior (opcional)",
    type=["xlsx"],
    key="checkpoint"
)

st.caption(
    "Cada pulsación procesa hasta 100 "
    "referencias pendientes. Descarga el "
    "progreso al terminar cada lote."
)


if not file:
    st.info("Carga el Excel para empezar.")
    st.stop()

if not re.fullmatch(
    r"[A-Z]{2}-MF-\d+", srn
):
    st.info(
        "Introduce un SRN de fabricante válido."
    )
    st.stop()


try:
    source = read_excel_or_csv(file)

    ref_col = next(
        (
            c for c in source.columns
            if str(c).strip().upper()
            == "MANUFACT_ITEMNUM"
        ),
        None
    )

    if ref_col is None:
        st.error(
            "El Excel debe incluir la columna "
            "MANUFACT_ITEMNUM."
        )
        st.stop()

    if not 1 <= len(source) <= MAX_ROWS:
        st.error(
            f"El archivo debe tener entre "
            f"1 y {MAX_ROWS} filas."
        )
        st.stop()

    references = [
        str(v).strip()
        for v in source[ref_col].tolist()
    ]

    key = fingerprint(
        srn,
        references
    )

except Exception as exc:
    st.error(
        f"No se pudo leer el archivo: {exc}"
    )
    st.stop()


if st.session_state.get("active_key") != key:
    st.session_state["active_key"] = key

    st.session_state["results"] = new_session(
        srn,
        references
    )

    st.session_state[
        "last_checkpoint_signature"
    ] = None


if checkpoint is not None:
    signature = hashlib.sha256(
        checkpoint.getvalue()
    ).hexdigest()

    if (
        st.session_state.get(
            "last_checkpoint_signature"
        ) != signature
    ):
        try:
            previous = pd.read_excel(
                io.BytesIO(
                    checkpoint.getvalue()
                ),
                sheet_name="Resultados",
                dtype=str,
                keep_default_na=False
            )

            control = pd.read_excel(
                io.BytesIO(
                    checkpoint.getvalue()
                ),
                sheet_name="Control",
                dtype=str,
                keep_default_na=False
            )

            if not set(
                RESULT_COLS
            ).issubset(previous.columns):
                raise ValueError(
                    "El archivo de progreso no "
                    "contiene las columnas esperadas"
                )

            if (
                control.empty
                or str(
                    control.iloc[0].get(
                        "Fingerprint", ""
                    )
                ) != key
            ):
                raise ValueError(
                    "El progreso corresponde a "
                    "otro SRN u otro listado"
                )

            if (
                len(previous) != len(references)
                or previous[
                    "MANUFACT_ITEMNUM"
                ].astype(str).tolist() != references
            ):
                raise ValueError(
                    "El orden o contenido de las "
                    "referencias no coincide"
                )

            st.session_state[
                "results"
            ] = previous[RESULT_COLS].copy()

            st.session_state[
                "last_checkpoint_signature"
            ] = signature

            st.success(
                "Progreso anterior recuperado."
            )

        except Exception as exc:
            st.error(
                f"No se pudo recuperar el "
                f"progreso: {exc}"
            )


results = st.session_state["results"]

retry_states = [
    "Error de consulta",
    "Sin coincidencia verificable",
    "Consulta parcial: revisar",
    "Encontrado sin UDI-DI: revisar"
]

retry_unverified = st.checkbox(
    "Reintentar referencias sin confirmar",
    value=False,
    help=(
        "Actívalo para volver a consultar "
        "referencias que ya fueron revisadas "
        "pero no obtuvieron un resultado definitivo."
    )
)

pending_states = ["Pendiente"]

if retry_unverified:
    pending_states.extend(retry_states)

pending = results.index[
    results["Estado"].isin(pending_states)
].tolist()

found = int(
    (results["Estado"] == "Encontrado").sum()
)

a, b, c = st.columns(3)

a.metric(
    "Referencias",
    len(results)
)

b.metric(
    "UDI-DI encontrados",
    found
)

c.metric(
    "Pendientes / por revisar",
    len(results) - found
)


if st.button(
    "Procesar siguiente lote (hasta 100)",
    type="primary",
    disabled=not pending
):
    selected = pending[:BATCH_SIZE]

    progress = st.progress(0)
    note = st.empty()

    for count, idx in enumerate(selected, 1):
        ref = references[idx]

        note.write(
            f"Consultando {count}/"
            f"{len(selected)}: {ref}"
        )

        if not clean(ref):
            update = {
                "Estado": "Referencia vacía",
                "Detalle": (
                    "Sin código de fabricante"
                )
            }

        else:
            try:
                update = api_lookup(
                    srn,
                    ref
                )

            except requests.exceptions.RequestException as exc:
                update = {
                    "Estado": "Error de consulta",
                    "Detalle": str(exc)[:300]
                }

            except Exception as exc:
                update = {
                    "Estado": "Error de consulta",
                    "Detalle": str(exc)[:300]
                }

        for col, val in update.items():
            results.at[idx, col] = val

        progress.progress(
            count / len(selected)
        )

        if count < len(selected):
            time.sleep(0.2)

    st.session_state["results"] = results

    note.empty()

    st.success(
        f"Lote terminado: {len(selected)} "
        f"referencias revisadas. "
        f"Descarga el progreso para conservarlo."
    )


st.dataframe(
    st.session_state["results"],
    hide_index=True,
    use_container_width=True
)

st.download_button(
    "Descargar resultados y progreso (Excel)",
    data=excel_bytes(
        st.session_state["results"],
        srn,
        key
    ),
    file_name=(
        f"UDI_Connect_{srn}_progreso.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-"
        "officedocument.spreadsheetml.sheet"
    )
)

st.caption(
    "Para reanudar después de un reinicio, "
    "sube el Excel de progreso descargado. "
    "Los resultados requieren validación "
    "regulatoria antes de comunicarse a AEMPS."
)
