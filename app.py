
import io
import re
import json
import hashlib
import time
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


st.set_page_config(
    page_title="UDI Connect 6.2",
    page_icon="🔎",
    layout="wide"
)

st.title("🔎 UDI Connect 6.2")
st.caption(
    "Hasta 5.000 referencias · SRN dinámicos · "
    "Lotes de 10, 25, 50 o 100 · Recuperación mediante Excel"
)

API = st.secrets.get(
    "EUDAMED_API_URL",
    "https://api.datalake.sante.service.ec.europa.eu/eudamed/udi"
)

MAX_ROWS = 5000

COLUMNS = [
    "MANUFACT_ITEMNUM",
    "SRN",
    "UDI-DI",
    "Basic UDI-DI",
    "Master UDI-DI",
    "Fabricante",
    "Estado",
    "Detalle",
    "Fecha consulta UTC"
]


def clean(value):
    return (
        re.sub(r"\s+", "", str(value or ""))
        .upper()
        .replace("–", "-")
        .replace("—", "-")
    )


def read_input(upload):
    raw = upload.getvalue()

    if upload.name.lower().endswith(".csv"):
        for encoding in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(
                    io.BytesIO(raw),
                    dtype=str,
                    keep_default_na=False,
                    sep=None,
                    engine="python",
                    encoding=encoding
                )
            except (UnicodeError, pd.errors.ParserError):
                continue

        raise ValueError("No se pudo leer el CSV")

    return pd.read_excel(
        io.BytesIO(raw),
        dtype=str,
        keep_default_na=False
    )


def request_session():
    session = requests.Session()

    retries = Retry(
        total=1,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )

    session.mount(
        "https://",
        HTTPAdapter(max_retries=retries)
    )

    session.headers.update({
        "Accept": "application/json",
        "User-Agent": "UDIConnect/6.2"
    })

    return session


def query(session, srn, reference):
    response = session.get(
        API,
        params={
            "api-version": "v1.0",
            "format": "json",
            "MF_SRN": srn,
            "REFERENCE": reference
        },
        timeout=(8, 20)
    )

    response.raise_for_status()
    payload = response.json()

    if (
        not isinstance(payload, dict)
        or not isinstance(payload.get("value"), list)
    ):
        raise ValueError(
            "Respuesta inesperada: falta la lista value"
        )

    rows = payload["value"]

    if len(rows) >= 1000:
        return {
            "Estado": "Revisar: límite API",
            "Detalle": (
                "La respuesta alcanza 1.000 registros "
                "y podría estar truncada"
            )
        }

    matches = [
        row
        for row in rows
        if isinstance(row, dict)
        and clean(row.get("MF_SRN")) == clean(srn)
        and clean(row.get("REFERENCE")) == clean(reference)
    ]

    if not matches:
        return {
            "Estado": "Sin coincidencia en API",
            "Detalle": (
                "No se encontró coincidencia exacta. "
                "Verificar en EUDAMED."
            )
        }

    variants = {
        (
            str(row.get("PRIMARY_DI") or ""),
            str(row.get("BASIC_UDI") or "")
        )
        for row in matches
    }

    if len(variants) != 1:
        return {
            "Estado": "Revisar: múltiples UDI",
            "Detalle": (
                f"{len(variants)} combinaciones diferentes. "
                "Revisar manualmente."
            )
        }

    primary, basic = next(iter(variants))

    if not primary:
        return {
            "Estado": "Revisar: sin UDI-DI",
            "Detalle": "Coincidencia sin PRIMARY_DI"
        }

    return {
        "Estado": "Encontrado",
        "UDI-DI": primary,
        "Basic UDI-DI": basic,
        "Fabricante": str(
            matches[0].get("MF_NAME") or ""
        ),
        "Detalle": (
            "Coincidencia exacta de SRN y referencia"
        )
    }


def digest(srns, refs):
    data = json.dumps(
        [srns, refs],
        ensure_ascii=False
    )

    return hashlib.sha256(
        data.encode()
    ).hexdigest()


def export_xlsx(df, srn, signature):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Resultados"
        )

        control = pd.DataFrame([{
            "SRN": srn,
            "Fingerprint": signature,
            "Fecha UTC": datetime.now(
                timezone.utc
            ).isoformat(),
            "Fuente": API
        }])

        control.to_excel(
            writer,
            index=False,
            sheet_name="Control"
        )

        for row in writer.sheets[
            "Resultados"
        ].iter_rows(min_row=2):
            for cell in row:
                cell.number_format = "@"

    return output.getvalue()


# ========================================
# ENTRADA DE DATOS
# ========================================

srn = st.text_input(
    "SRN predeterminado",
    placeholder="ES-MF-000002991",
    help=(
        "Opcional si el Excel contiene "
        "una columna SRN para cada referencia."
    )
).strip().upper()

upload = st.file_uploader(
    "Archivo Excel o CSV",
    type=["xlsx", "csv"],
    help="Debe incluir MANUFACT_ITEMNUM"
)

checkpoint = st.file_uploader(
    "Recuperar progreso anterior (opcional)",
    type=["xlsx"],
    key="progress_upload"
)

size = st.selectbox(
    "Referencias por lote",
    [10, 25, 50, 100],
    index=1
)

if not upload:
    st.info(
        "Introduce un SRN y adjunta el Excel, "
        "o utiliza un Excel con columnas "
        "SRN y MANUFACT_ITEMNUM."
    )
    st.stop()


# ========================================
# VALIDACIÓN DEL EXCEL
# ========================================

try:
    source = read_input(upload)

    col = next(
        (
            c for c in source.columns
            if str(c).strip().upper()
            == "MANUFACT_ITEMNUM"
        ),
        None
    )

    if col is None:
        st.error(
            "Falta la columna MANUFACT_ITEMNUM."
        )
        st.stop()

    if not 1 <= len(source) <= MAX_ROWS:
        st.error(
            "El archivo debe contener entre "
            "1 y 5.000 filas."
        )
        st.stop()

    refs = [
        str(value).strip()
        for value in source[col].tolist()
    ]

    srn_col = next(
        (
            c for c in source.columns
            if str(c).strip().upper() == "SRN"
        ),
        None
    )

    if srn_col:
        srns = [
            str(value).strip().upper() or srn
            for value in source[srn_col].tolist()
        ]
    else:
        srns = [srn] * len(refs)

    invalid = [
        i + 2
        for i, value in enumerate(srns)
        if not re.fullmatch(
            r"[A-Z]{2}-(MF|PR)-\d{9}",
            value
        )
    ]

    if invalid:
        raise ValueError(
            "SRN no válido o vacío en las "
            f"filas: {invalid[:12]}"
        )

    if any("-PR-" in value for value in srns):
        st.warning(
            "El archivo contiene SRN de tipo PR. "
            "Se conservarán en los resultados, "
            "pero no se consultarán mediante MF_SRN."
        )

except Exception as exc:
    st.error(
        f"No se pudo abrir el archivo: {exc}"
    )
    st.stop()


# ========================================
# ESTADO Y RECUPERACIÓN
# ========================================

signature = digest(srns, refs)

if st.session_state.get("signature") != signature:

    st.session_state.signature = signature

    st.session_state.results = pd.DataFrame(
        [
            {
                "MANUFACT_ITEMNUM": ref,
                "SRN": srns[i],
                "UDI-DI": "",
                "Basic UDI-DI": "",
                "Master UDI-DI": "",
                "Fabricante": "",
                "Estado": "Pendiente",
                "Detalle": "",
                "Fecha consulta UTC": ""
            }
            for i, ref in enumerate(refs)
        ],
        columns=COLUMNS
    )

    st.session_state.loaded_checkpoint = None
    st.session_state.api_checked = False


if checkpoint:

    checkpoint_hash = hashlib.sha256(
        checkpoint.getvalue()
    ).hexdigest()

    if (
        st.session_state.get("loaded_checkpoint")
        != checkpoint_hash
    ):
        try:
            old = pd.read_excel(
                io.BytesIO(checkpoint.getvalue()),
                sheet_name="Resultados",
                dtype=str,
                keep_default_na=False
            )

            meta = pd.read_excel(
                io.BytesIO(checkpoint.getvalue()),
                sheet_name="Control",
                dtype=str,
                keep_default_na=False
            )

            if (
                meta.empty
                or str(
                    meta.iloc[0].get(
                        "Fingerprint", ""
                    )
                ) != signature
            ):
                raise ValueError(
                    "El progreso no corresponde "
                    "a este archivo."
                )

            if (
                not set(COLUMNS).issubset(
                    old.columns
                )
                or old[
                    "MANUFACT_ITEMNUM"
                ].tolist() != refs
                or old["SRN"].tolist() != srns
            ):
                raise ValueError(
                    "El archivo de progreso "
                    "no coincide con las referencias."
                )

            st.session_state.results = (
                old[COLUMNS].copy()
            )

            st.session_state.loaded_checkpoint = (
                checkpoint_hash
            )

            st.success(
                "Progreso restaurado correctamente."
            )

        except Exception as exc:
            st.error(
                f"No se pudo restaurar: {exc}"
            )


# ========================================
# INDICADORES
# ========================================

results = st.session_state.results

completed = int(
    (results["Estado"] != "Pendiente").sum()
)

found = int(
    (results["Estado"] == "Encontrado").sum()
)

c1, c2, c3 = st.columns(3)

c1.metric("Total", len(results))
c2.metric("Consultadas", completed)
c3.metric("UDI-DI encontrados", found)

st.progress(completed / len(results))


# ========================================
# PRUEBA DE CONEXIÓN
# ========================================

st.markdown("### 1. Comprobar conexión")

if st.button(
    "Probar API con la primera referencia"
):

    test_idx = next(
        (
            i for i, ref in enumerate(refs)
            if ref and "-MF-" in srns[i]
        ),
        None
    )

    test_ref = (
        refs[test_idx]
        if test_idx is not None
        else None
    )

    if not test_ref:
        st.error(
            "No hay referencias MF para probar."
        )

    else:
        try:
            with request_session() as session:
                test_result = query(
                    session,
                    srns[test_idx],
                    test_ref
                )

            st.session_state.api_checked = True

            st.success(
                "La API respondió con el "
                "formato esperado."
            )

            st.write({
                "SRN": srns[test_idx],
                "Referencia": test_ref,
                **test_result
            })

        except requests.HTTPError as exc:

            st.session_state.api_checked = False

            st.error(
                f"Error HTTP: {exc}. "
                "La conexión no está validada. "
                "No ejecutes lotes."
            )

        except Exception as exc:

            st.session_state.api_checked = False

            st.error(
                f"La API no está lista: {exc}"
            )


# ========================================
# PROCESAMIENTO POR LOTES
# ========================================

st.markdown("### 2. Procesar referencias")

pending = results.index[
    results["Estado"] == "Pendiente"
].tolist()

if st.button(
    f"Procesar siguiente lote (máximo {size})",
    type="primary",
    disabled=(
        not pending
        or not st.session_state.api_checked
    )
):

    indexes = pending[:size]

    progress = st.progress(0)
    label = st.empty()

    cache = {}
    fatal_error = None

    with request_session() as session:

        for count, idx in enumerate(indexes, 1):

            ref = refs[idx]

            label.caption(
                f"Consulta {count}/{len(indexes)}: "
                f"{ref}"
            )

            if not ref:

                item = {
                    "Estado": "Referencia vacía",
                    "Detalle": (
                        "No se proporcionó código"
                    )
                }

            elif "-PR-" in srns[idx]:

                item = {
                    "Estado": (
                        "PR: consulta no implementada"
                    ),
                    "Detalle": (
                        "No utilizar MF_SRN "
                        "para un productor PR"
                    )
                }

            elif (
                srns[idx],
                clean(ref)
            ) in cache:

                item = cache[
                    (srns[idx], clean(ref))
                ]

            else:

                try:
                    item = query(
                        session,
                        srns[idx],
                        ref
                    )

                    cache[
                        (srns[idx], clean(ref))
                    ] = item

                except requests.HTTPError as exc:

                    if (
                        exc.response is not None
                        and exc.response.status_code
                        in (400, 401, 403, 404)
                    ):

                        fatal_error = (
                            f"Error HTTP "
                            f"{exc.response.status_code}: "
                            f"{exc}"
                        )

                        st.session_state.api_checked = (
                            False
                        )

                        break

                    item = {
                        "Estado": "Error temporal",
                        "Detalle": str(exc)[:250]
                    }

                except Exception as exc:

                    item = {
                        "Estado": "Error temporal",
                        "Detalle": str(exc)[:250]
                    }

            for field, value in item.items():
                results.at[idx, field] = value

            results.at[
                idx,
                "Fecha consulta UTC"
            ] = datetime.now(
                timezone.utc
            ).isoformat()

            progress.progress(
                count / len(indexes)
            )

            if count < len(indexes):
                time.sleep(0.3)

    st.session_state.results = results
    label.empty()

    if fatal_error:
        st.error(
            "Proceso detenido para evitar "
            f"errores repetidos. {fatal_error}"
        )
    else:
        st.success(
            "Lote finalizado. "
            "Descarga el progreso."
        )


# ========================================
# RESULTADOS Y DESCARGA
# ========================================

st.markdown("### 3. Resultados")

st.dataframe(
    st.session_state.results,
    use_container_width=True,
    hide_index=True
)

st.download_button(
    "⬇️ Descargar resultados y progreso",
    data=export_xlsx(
        st.session_state.results,
        srn,
        signature
    ),
    file_name="UDI_Connect_progreso.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )
)

st.warning(
    "La dirección predeterminada de la API "
    "todavía no está validada. "
    "Anteriormente devolvió HTTP 404. "
    "No proceses lotes hasta superar "
    "la prueba de conexión."
)

st.caption(
    "La recuperación utiliza el Excel descargado. "
    "No existe guardado automático en servidor. "
    "Master UDI-DI permanecerá vacío cuando "
    "la fuente no proporcione ese dato. "
    "Verifica los resultados en EUDAMED "
    "antes de utilizarlos en RECOPS."
)
