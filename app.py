
import io
import re
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import streamlit as st

st.set_page_config(
    page_title="UDI Connect",
    page_icon="🔎",
    layout="wide"
)

st.title("UDI Connect")
st.caption(
    "Consulta de referencias de fabricantes "
    "en el registro público de EUDAMED"
)

st.info(
    "UDI-DI, Basic UDI-DI y Master UDI-DI "
    "son identificadores diferentes. "
    "Los resultados requieren validación regulatoria."
)

BASE = (
    "https://ec.europa.eu/tools/eudamed/"
    "api/devices/udiDiData"
)
PUBLIC = (
    "https://ec.europa.eu/tools/eudamed/"
    "#/screen/search-device"
)


def norm(value):
    return re.sub(
        r"\s+", "", str(value or "")
    ).upper().replace("–", "-").replace("—", "-")


def read_file(upload):
    data = upload.getvalue()

    if upload.name.lower().endswith(".csv"):
        for encoding in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(
                    io.BytesIO(data),
                    dtype=str,
                    sep=None,
                    engine="python",
                    encoding=encoding
                ).fillna("")
            except (UnicodeError, pd.errors.ParserError):
                continue

        raise ValueError("No se pudo leer el CSV.")

    return pd.read_excel(
        io.BytesIO(data),
        dtype=str
    ).fillna("")


@st.cache_data(ttl=900, show_spinner=False)
def fetch_page(srn, page, page_size):
    params = {
        "page": page,
        "pageSize": page_size,
        "size": page_size,
        "iso2Code": "en",
        "languageIso2Code": "en",
        "srn": srn
    }

    retry = Retry(
        total=2,
        connect=2,
        read=2,
        status=2,
        backoff_factor=1.5,
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
            BASE,
            params=params,
            timeout=(12, 60),
            headers={
                "Accept": "application/json",
                "User-Agent": "UDIConnectPilot/3.1"
            }
        )

        response.raise_for_status()
        payload = response.json()

    if (
        not isinstance(payload, dict)
        or not isinstance(payload.get("content"), list)
    ):
        raise ValueError(
            "Respuesta EUDAMED no reconocida."
        )

    return payload


with st.sidebar:
    st.header("Fuente de datos")
    st.markdown(
        f"[Abrir EUDAMED oficial]({PUBLIC})"
    )
    st.caption(
        "La disponibilidad del servicio "
        "puede variar."
    )


srn = st.text_input(
    "SRN del fabricante",
    value="ES-MF-000002991"
).strip().upper()

file = st.file_uploader(
    "Sube tu Excel con MANUFACT_ITEMNUM",
    type=["xlsx", "csv"]
)

page_size = st.selectbox(
    "Registros por página",
    [10, 20, 25, 50, 100],
    index=1
)

max_pages = st.number_input(
    "Máximo de páginas",
    min_value=1,
    max_value=10000,
    value=10,
    step=1
)

st.caption(
    "Si la consulta queda incompleta, "
    "las referencias sin coincidencia "
    "permanecerán pendientes."
)


if file is not None:
    try:
        original = read_file(file)

        col = next(
            (
                c for c in original.columns
                if str(c).strip().upper()
                == "MANUFACT_ITEMNUM"
            ),
            None
        )

        if col is None:
            st.error(
                "El Excel debe contener la columna "
                "MANUFACT_ITEMNUM."
            )
            st.stop()

        refs = original[col].astype(str).tolist()

        st.success(
            f"Archivo leído: {len(refs)} referencias."
        )

        st.dataframe(
            pd.DataFrame({
                "MANUFACT_ITEMNUM": refs
            }).head(10),
            hide_index=True
        )

    except Exception as exc:
        st.error(
            f"No se pudo abrir el archivo: {exc}"
        )
        st.stop()

else:
    refs = None


if st.button(
    "Buscar UDI-DI en EUDAMED",
    type="primary",
    disabled=(refs is None or not srn)
):
    if not re.fullmatch(
        r"[A-Z]{2}-MF-\d+", srn
    ):
        st.error(
            "Formato de SRN incorrecto."
        )
        st.stop()

    progress = st.progress(
        0,
        text="Iniciando consulta..."
    )

    records = []
    complete = False
    error = None
    pages_checked = 0
    total_pages = None

    try:
        for page in range(int(max_pages)):
            payload = fetch_page(
                srn,
                page,
                int(page_size)
            )

            chunk = payload["content"]

            for item in chunk:
                if not isinstance(item, dict):
                    continue

                record_srn = str(
                    item.get("manufacturerSrn") or ""
                ).strip().upper()

                if record_srn == srn:
                    records.append(item)

            pages_checked += 1

            if total_pages is None:
                try:
                    total_pages = int(
                        payload["totalPages"]
                    )
                except (
                    KeyError,
                    ValueError,
                    TypeError
                ):
                    total_pages = None

            progress.progress(
                min(
                    (page + 1) / int(max_pages),
                    1.0
                ),
                text=(
                    f"Páginas consultadas: "
                    f"{pages_checked}"
                )
            )

            if (
                payload.get("last") is True
                or (
                    total_pages is not None
                    and page + 1 >= total_pages
                )
                or not chunk
                or len(chunk) < int(page_size)
            ):
                complete = True
                break

    except requests.exceptions.Timeout:
        error = (
            "Tiempo de espera agotado "
            "al contactar con EUDAMED."
        )

    except requests.exceptions.RequestException as exc:
        error = f"Error de conexión: {exc}"

    except Exception as exc:
        error = str(exc)

    progress.empty()

    if error:
        st.error(
            f"La consulta no se completó: {error}"
        )
        st.warning(
            "Las referencias no localizadas "
            "quedan pendientes de revisión."
        )

    elif not complete:
        st.warning(
            "Consulta parcial: se alcanzó "
            "el límite de páginas."
        )

    else:
        st.success(
            "Consulta de páginas completada."
        )

    st.write(
        f"Registros recibidos: {len(records)} "
        f"· Páginas: {pages_checked}"
    )

    if pages_checked > 0 and not records:
        st.warning(
            "No se pudieron confirmar registros "
            "con el SRN indicado."
        )
        complete = False

    lookup = {}

    for item in records:
        reference = norm(
            item.get("reference")
        )

        if reference:
            lookup.setdefault(
                reference, []
            ).append(item)

    result_rows = []

    for ref in refs:
        key = norm(ref)

        matches = (
            lookup.get(key, [])
            if key else []
        )

        row = {
            "MANUFACT_ITEMNUM": ref,
            "SRN": srn,
            "UDI-DI": "",
            "Basic UDI-DI": "",
            "Master UDI-DI": "",
            "Fabricante EUDAMED": "",
            "Referencia EUDAMED": "",
            "Estado": "",
            "Enlace EUDAMED": ""
        }

        if not key:
            row["Estado"] = "Referencia vacía"

        elif not matches:
            if complete and not error:
                row["Estado"] = (
                    "Sin coincidencia en consulta completa"
                )
            else:
                row["Estado"] = (
                    "Pendiente: consulta incompleta"
                )

        else:
            latest = [
                x for x in matches
                if x.get("latestVersion") is True
            ]

            candidates = latest or matches

            distinct = {
                (
                    str(x.get("primaryDi") or ""),
                    str(x.get("uuid") or "")
                )
                for x in candidates
            }

            if len(distinct) != 1:
                row["Estado"] = (
                    "Coincidencias múltiples: revisar"
                )

            else:
                item = candidates[0]

                row["UDI-DI"] = str(
                    item.get("primaryDi") or ""
                )

                row["Basic UDI-DI"] = str(
                    item.get("basicUdi") or ""
                )

                row["Fabricante EUDAMED"] = str(
                    item.get("manufacturerName") or ""
                )

                row["Referencia EUDAMED"] = str(
                    item.get("reference") or ""
                )

                uuid = str(
                    item.get("uuid") or ""
                )

                if uuid:
                    row["Enlace EUDAMED"] = (
                        f"{PUBLIC}/{uuid}"
                    )

                if row["UDI-DI"]:
                    row["Estado"] = (
                        "Coincidencia exacta "
                        "por referencia y SRN"
                    )
                else:
                    row["Estado"] = (
                        "Coincidencia sin UDI-DI: revisar"
                    )

        result_rows.append(row)

    result = pd.DataFrame(result_rows)

    st.session_state["udi_results"] = result
    st.session_state["udi_query"] = {
        "srn": srn,
        "complete": complete,
        "error": error,
        "pages": pages_checked,
        "records": len(records)
    }


if "udi_results" in st.session_state:
    result = st.session_state["udi_results"]

    info = st.session_state.get(
        "udi_query", {}
    )

    st.subheader("Resultados")

    a, b, c = st.columns(3)

    a.metric(
        "Referencias",
        len(result)
    )

    b.metric(
        "Encontradas",
        int(
            result["Estado"].eq(
                "Coincidencia exacta por referencia y SRN"
            ).sum()
        )
    )

    c.metric(
        "Pendientes o por revisar",
        int(
            (
                result["Estado"]
                != "Coincidencia exacta por referencia y SRN"
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

        pd.DataFrame([info]).to_excel(
            writer,
            sheet_name="Control de consulta",
            index=False
        )

        sheet = writer.sheets["Resultados"]

        for cells in sheet.columns:
            for cell in cells[1:]:
                cell.number_format = "@"

    st.download_button(
        "Descargar Excel de resultados",
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
    "no afiliada a la Comisión Europea. "
    "Verifica los datos en EUDAMED "
    "antes de comunicarlos a AEMPS."
)
