
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
    page_title="UDI Connect 6.1",
    page_icon="🔎",
    layout="wide"
)

st.title("🔎 UDI Connect 6.1")

st.caption(
    "Hasta 5.000 referencias · "
    "lotes de 10, 25, 50 o 100 · "
    "recuperación mediante Excel"
)

API = (
    "https://api.datalake.sante.service.ec.europa.eu"
    "/eudamed/udi"
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


def clean(v):
    return re.sub(
        r"\s+", "", str(v or "")
    ).upper().replace(
        "–", "-"
    ).replace("—", "-")


def read_input(upload):
    raw = upload.getvalue()

    if upload.name.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(
                    io.BytesIO(raw),
                    dtype=str,
                    keep_default_na=False,
                    sep=None,
                    engine="python",
                    encoding=enc
                )
            except (
                UnicodeError,
                pd.errors.ParserError
            ):
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
        status_forcelist=[
            429, 500, 502, 503, 504
        ],
        allowed_methods=["GET"]
    )

    session.mount(
        "https://",
        HTTPAdapter(max_retries=retries)
    )

    session.headers.update({
        "Accept": "application/json",
        "User-Agent": "UDIConnect/6.1"
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
        or not isinstance(
            payload.get("value"), list
        )
    ):
        raise ValueError(
            "Respuesta inesperada: "
            "falta la lista value"
        )

    rows = payload["value"]

    if len(rows) >= 1000:
        return {
            "Estado": "Revisar: límite API",
            "Detalle": (
                "La respuesta alcanza 1.000 "
                "registros y podría estar truncada"
            )
        }

    matches = [
        r for r in rows
        if isinstance(r, dict)
        and clean(r.get("MF_SRN")) == clean(srn)
        and clean(r.get("REFERENCE")) == clean(reference)
    ]

    if not matches:
        return {
            "Estado": "Sin coincidencia en API",
            "Detalle": (
                "No se encontró coincidencia "
                "exacta en la descarga pública; "
                "verificar EUDAMED"
            )
        }

    variants = {
        (
            str(r.get("PRIMARY_DI") or ""),
            str(r.get("BASIC_UDI") or "")
        )
        for r in matches
    }

    if len(variants) != 1:
        return {
            "Estado": "Revisar: múltiples UDI",
            "Detalle": (
                f"{len(variants)} combinaciones "
                "diferentes; revisar manualmente"
            )
        }

    primary, basic = next(iter(variants))

    if not primary:
        return {
            "Estado": "Revisar: sin UDI-DI",
            "Detalle": (
                "Coincidencia sin PRIMARY_DI"
            )
        }

    return {
        "Estado": "Encontrado",
        "UDI-DI": primary,
        "Basic UDI-DI": basic,
        "Fabricante": str(
            matches[0].get("MF_NAME") or ""
        ),
        "Detalle": (
            "Coincidencia exacta de SRN "
            "y referencia (Data Lake)"
        )
    }


def digest(srn, refs):
    data = json.dumps(
        [srn, refs],
        ensure_ascii=False
    ).encode("utf-8")

    return hashlib.sha256(data).hexdigest()


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


srn = st.text_input(
    "SRN del fabricante",
    placeholder="ES-MF-000002991"
).strip().upper()

upload = st.file_uploader(
    "Archivo Excel o CSV "
    "(columna MANUFACT_ITEMNUM)",
    type=["xlsx", "csv"]
)

checkpoint = st.file_uploader(
    "Recuperar Excel de progreso "
    "anterior (opcional)",
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
        "Introduce un SRN y adjunta "
        "el Excel de referencias."
    )
    st.stop()


if not re.fullmatch(
    r"[A-Z]{2}-(MF|PR)-\d{9}", srn
):
    st.error(
        "El SRN debe tener un formato como "
        "ES-MF-000002991 o ES-PR-000042695."
    )
    st.stop()


if "-PR-" in srn:
    st.warning(
        "SRN PR reconocido. El filtro MF_SRN "
        "de esta API corresponde a fabricantes "
        "MF; no se realizará una búsqueda PR "
        "no verificada."
    )
    st.stop()


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
            "El archivo debe contener "
            "entre 1 y 5.000 filas."
        )
        st.stop()

    refs = [
        str(x).strip()
        for x in source[col].tolist()
    ]

except Exception as exc:
    st.error(
        f"No se pudo abrir el archivo: {exc}"
    )
    st.stop()


signature = digest(srn, refs)


if st.session_state.get(
    "signature"
) != signature:

    st.session_state.signature = signature

    st.session_state.results = pd.DataFrame(
        [
            {
                "MANUFACT_ITEMNUM": ref,
                "SRN": srn,
                "UDI-DI": "",
                "Basic UDI-DI": "",
                "Master UDI-DI": "",
                "Fabricante": "",
                "Estado": "Pendiente",
                "Detalle": "",
                "Fecha consulta UTC": ""
            }
            for ref in refs
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
        st.session_state.get(
            "loaded_checkpoint"
        ) != checkpoint_hash
    ):
        try:
            old = pd.read_excel(
                io.BytesIO(
                    checkpoint.getvalue()
                ),
                sheet_name="Resultados",
                dtype=str,
                keep_default_na=False
            )

            meta = pd.read_excel(
                io.BytesIO(
                    checkpoint.getvalue()
                ),
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
                    "a este SRN y archivo"
                )

            if (
                not set(COLUMNS).issubset(
                    old.columns
                )
                or old[
                    "MANUFACT_ITEMNUM"
                ].tolist() != refs
            ):
                raise ValueError(
                    "El archivo de progreso "
                    "no coincide con las referencias"
                )

            st.session_state.results = old[
                COLUMNS
            ].copy()

            st.session_state.loaded_checkpoint = (
                checkpoint_hash
            )

            st.success(
                "Progreso restaurado correctamente."
            )

        except Exception as exc:
            st.error(
                "No se pudo restaurar el "
                f"progreso: {exc}"
            )


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


st.markdown(
    "### 1. Comprobar la conexión"
)

if st.button(
    "Probar API con la primera "
    "referencia no vacía"
):
    test_ref = next(
        (x for x in refs if x),
        None
    )

    if not test_ref:
        st.error(
            "No hay referencias para probar."
        )

    else:
        try:
            with request_session() as session:
                test_result = query(
                    session,
                    srn,
                    test_ref
                )

            st.session_state.api_checked = True

            st.success(
                "El servicio respondió con el "
                "formato esperado. "
                "Puedes procesar lotes."
            )

            st.write({
                "Referencia": test_ref,
                **test_result
            })

        except requests.HTTPError as exc:
            st.session_state.api_checked = False

            st.error(
                f"Error HTTP: {exc}. "
                "No ejecutes lotes hasta resolver "
                "el acceso a la API."
            )

        except Exception as exc:
            st.session_state.api_checked = False

            st.error(
                f"La API no está lista: {exc}"
            )


st.markdown(
    "### 2. Procesar y descargar"
)

pending = results.index[
    results["Estado"] == "Pendiente"
].tolist()


if st.button(
    f"Procesar siguiente lote "
    f"(máximo {size})",
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
        for count, idx in enumerate(
            indexes, 1
        ):
            ref = refs[idx]

            label.caption(
                f"Consulta {count}/"
                f"{len(indexes)}: {ref}"
            )

            if not ref:
                item = {
                    "Estado": "Referencia vacía",
                    "Detalle": (
                        "No se proporcionó código"
                    )
                }

            elif clean(ref) in cache:
                item = cache[clean(ref)]

            else:
                try:
                    item = query(
                        session,
                        srn,
                        ref
                    )

                    cache[clean(ref)] = item

                except requests.HTTPError as exc:
                    if (
                        exc.response is not None
                        and exc.response.status_code
                        in (400, 401, 403, 404)
                    ):
                        fatal_error = (
                            "Error HTTP "
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
                idx, "Fecha consulta UTC"
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
            "errores repetidos. "
            f"{fatal_error}"
        )

    else:
        st.success(
            "Lote finalizado. Descarga "
            "el progreso antes de continuar."
        )


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
    file_name=(
        f"UDI_Connect_{srn}_progreso.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )
)

st.caption(
    "El Excel es el mecanismo de recuperación "
    "después de un reinicio. "
    "No hay guardado automático en servidor. "
    "Master UDI-DI permanece vacío si la fuente "
    "no lo proporciona. "
    "Verifica los datos en EUDAMED antes "
    "de usarlos en RECOPS."
)
