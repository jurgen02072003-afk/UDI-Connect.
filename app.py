
import io
import re
import time
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


st.set_page_config(
    page_title="UDI Connect",
    page_icon="🔎",
    layout="wide"
)

st.title("UDI Connect")
st.caption(
    "Consulta de referencias de fabricantes "
    "en los datos públicos de EUDAMED"
)

st.info(
    "La búsqueda utiliza el SRN y la referencia "
    "del fabricante. Los resultados deben "
    "verificarse antes de comunicarlos a AEMPS."
)

API = (
    "https://api.datalake.sante.service.ec.europa.eu"
    "/eudamed/udi"
)

PUBLIC = (
    "https://ec.europa.eu/tools/eudamed/"
    "#/screen/search-device"
)


def norm(value):
    return re.sub(
        r"\s+", "", str(value or "")
    ).upper().replace("–", "-").replace("—", "-")


def as_text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def load_file(upload):
    raw = upload.getvalue()

    if upload.name.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(
                    io.BytesIO(raw),
                    dtype=str,
                    sep=None,
                    engine="python",
                    encoding=enc,
                    keep_default_na=False
                )
            except (
                UnicodeError,
                pd.errors.ParserError
            ):
                pass

        raise ValueError("No se pudo leer el CSV.")

    return pd.read_excel(
        io.BytesIO(raw),
        dtype=str,
        keep_default_na=False
    )


@st.cache_data(ttl=1800, show_spinner=False)
def lookup(srn, reference):
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
                "REFERENCE": reference
            },
            headers={
                "Accept": "application/json",
                "User-Agent": "UDIConnectPilot/4.0"
            },
            timeout=(12, 55)
        )

        response.raise_for_status()
        data = response.json()

    if (
        not isinstance(data, dict)
        or not isinstance(data.get("value"), list)
    ):
        raise ValueError(
            "La API devolvió un formato JSON inesperado."
        )

    return data["value"]


srn = st.text_input(
    "SRN del fabricante",
    value="ES-MF-000002991"
).strip().upper()

upload = st.file_uploader(
    "Excel o CSV con columna MANUFACT_ITEMNUM",
    type=["xlsx", "csv"]
)

limit = st.number_input(
    "Máximo de referencias por ejecución",
    min_value=1,
    max_value=5000,
    value=20,
    step=1
)

st.caption(
    "Recomendamos empezar con 1 a 20 referencias. "
    "Los lotes grandes pueden tardar."
)

frame = None
column = None

if upload is not None:
    try:
        frame = load_file(upload)

        column = next(
            (
                c for c in frame.columns
                if str(c).strip().upper()
                == "MANUFACT_ITEMNUM"
            ),
            None
        )

        if column is None:
            st.error(
                "Falta la columna MANUFACT_ITEMNUM."
            )
            frame = None

        else:
            st.success(
                f"Archivo leído: {len(frame)} filas."
            )

            st.dataframe(
                frame[[column]].head(8),
                hide_index=True
            )

    except Exception as exc:
        st.error(
            f"Error al abrir el archivo: {exc}"
        )
        frame = None


if st.button(
    "Buscar UDI-DI en EUDAMED",
    type="primary",
    disabled=frame is None
):
    if not re.fullmatch(
        r"[A-Z]{2}-MF-\d+", srn
    ):
        st.error(
            "Introduce un SRN de fabricante válido."
        )
        st.stop()

    refs = frame[column].astype(str).tolist()
    selected = refs[:int(limit)]

    unique = list(
        dict.fromkeys(
            as_text(r)
            for r in selected
            if as_text(r)
        )
    )

    findings = {}
    progress = st.progress(0)
    status = st.empty()

    for i, ref in enumerate(unique):
        status.write(
            f"Consultando {i + 1}/{len(unique)}: {ref}"
        )

        try:
            raw = lookup(srn, ref)

            verified = [
                r for r in raw
                if isinstance(r, dict)
                and norm(r.get("MF_SRN")) == norm(srn)
                and norm(r.get("REFERENCE")) == norm(ref)
            ]

            if len(raw) >= 1000:
                findings[ref] = (
                    "Respuesta posiblemente truncada",
                    []
                )

            elif not verified:
                findings[ref] = (
                    "Sin coincidencia exacta en API",
                    []
                )

            else:
                findings[ref] = (
                    "Encontrado",
                    verified
                )

        except requests.exceptions.RequestException as exc:
            findings[ref] = (
                f"Error de conexión: {exc}",
                []
            )

        except Exception as exc:
            findings[ref] = (
                f"Error de respuesta: {exc}",
                []
            )

        progress.progress(
            (i + 1) / max(len(unique), 1)
        )

        if i < len(unique) - 1:
            time.sleep(0.5)

    progress.empty()
    status.empty()

    output = []

    for ref in selected:
        ref = as_text(ref)

        state, matches = findings.get(
            ref,
            ("Referencia vacía", [])
        )

        row = {
            "MANUFACT_ITEMNUM": ref,
            "SRN": srn,
            "UDI-DI": "",
            "Basic UDI-DI": "",
            "Master UDI-DI": "",
            "Estado": state,
            "Fabricante EUDAMED": "",
            "Referencia EUDAMED": ""
        }

        if matches:
            pairs = {
                (
                    as_text(r.get("PRIMARY_DI")),
                    as_text(r.get("BASIC_UDI"))
                )
                for r in matches
            }

            if len(pairs) > 1:
                row["Estado"] = (
                    "Coincidencias múltiples: revisar"
                )

            else:
                item = matches[0]

                row["UDI-DI"] = as_text(
                    item.get("PRIMARY_DI")
                )

                row["Basic UDI-DI"] = as_text(
                    item.get("BASIC_UDI")
                )

                row["Fabricante EUDAMED"] = as_text(
                    item.get("MF_NAME")
                )

                row["Referencia EUDAMED"] = as_text(
                    item.get("REFERENCE")
                )

                if row["UDI-DI"]:
                    row["Estado"] = (
                        "Coincidencia exacta"
                    )
                else:
                    row["Estado"] = (
                        "Encontrado sin UDI-DI: revisar"
                    )

        output.append(row)

    result = pd.DataFrame(output)

    st.session_state["result"] = result

    st.session_state["metadata"] = {
        "SRN": srn,
        "Fecha UTC": datetime.now(
            timezone.utc
        ).isoformat(),
        "Referencias procesadas": len(selected),
        "Filas restantes": (
            len(refs) - len(selected)
        ),
        "Fuente": API,
        "Consulta": "MF_SRN + REFERENCE"
    }


if "result" in st.session_state:
    result = st.session_state["result"]

    st.subheader("Resultados")

    a, b, c = st.columns(3)

    a.metric(
        "Consultadas",
        len(result)
    )

    b.metric(
        "Coincidencias exactas",
        int(
            result["Estado"].eq(
                "Coincidencia exacta"
            ).sum()
        )
    )

    c.metric(
        "Pendientes / revisar",
        int(
            result["Estado"].ne(
                "Coincidencia exacta"
            ).sum()
        )
    )

    st.dataframe(
        result,
        hide_index=True
    )

    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ) as writer:
        result.to_excel(
            writer,
            sheet_name="Resultados",
            index=False
        )

        pd.DataFrame(
            [st.session_state["metadata"]]
        ).to_excel(
            writer,
            sheet_name="Control",
            index=False
        )

        for sheet in writer.book.worksheets:
            for excel_row in sheet.iter_rows(
                min_row=2
            ):
                for cell in excel_row:
                    cell.number_format = "@"

    st.download_button(
        "Descargar Excel",
        data=out.getvalue(),
        file_name="UDI_Connect_resultados.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )


st.divider()

st.caption(
    "UDI Connect es una herramienta piloto "
    "independiente. La API de descarga puede "
    "actualizarse con retraso respecto a la "
    "interfaz pública de EUDAMED. Un resultado "
    "vacío no demuestra que el producto no exista."
)
