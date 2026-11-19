
import io
import re
import time
from datetime import datetime, timezone
from urllib.parse import urlencode

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="UDI Connect | Automático",
    page_icon="🔎",
    layout="wide"
)

st.title("🔎 UDI Connect — búsqueda automática de UDI-DI")

st.caption(
    "Consulta EUDAMED por SRN y referencia del fabricante."
)

API = "https://api.datalake.sante.service.ec.europa.eu/eudamed/udi"

DEFAULT_SRNS = """KR-MF-000026311
CN-MF-000009139
JP-MF-000011271"""

COLUMNS = [
    "MANUFACT_ITEMNUM",
    "SRN",
    "UDI-DI",
    "Basic UDI-DI",
    "Fabricante",
    "Nombre producto",
    "Estado",
    "Fuente",
    "Fecha consulta UTC",
    "Enlace EUDAMED"
]

SRN_RE = re.compile(r"^[A-Z]{2}-(?:MF|PR)-\d{9}$")


def official_url(srn, ref):
    return (
        "https://ec.europa.eu/tools/eudamed/"
        "#/screen/search-device?"
        + urlencode({
            "srn": srn,
            "reference": ref,
            "submitted": "true"
        })
    )


def get_records(srn, ref, timeout=22):
    params = {
        "api-version": "v1.0",
        "format": "json",
        "MF_SRN": srn,
        "REFERENCE": ref
    }

    response = requests.get(
        API,
        params=params,
        timeout=timeout,
        headers={
            "Accept": "application/json",
            "User-Agent": "UDI-Connect/1.0"
        }
    )

    response.raise_for_status()
    data = response.json()

    if (
        not isinstance(data, dict)
        or not isinstance(data.get("value"), list)
    ):
        raise ValueError(
            "Respuesta JSON inesperada de EUDAMED."
        )

    return data["value"]


def normalize(value):
    return str(value or "").strip().casefold()


def match_records(rows, srn, ref):
    return [
        r for r in rows
        if (
            isinstance(r, dict)
            and normalize(r.get("MF_SRN")) == normalize(srn)
            and normalize(r.get("REFERENCE")) == normalize(ref)
        )
    ]


def make_rows(srn, ref, found, now):
    base = {
        "MANUFACT_ITEMNUM": ref,
        "SRN": srn,
        "Fecha consulta UTC": now,
        "Enlace EUDAMED": official_url(srn, ref)
    }

    if not found:
        return [{
            **base,
            "UDI-DI": "",
            "Basic UDI-DI": "",
            "Fabricante": "",
            "Nombre producto": "",
            "Estado": "Sin coincidencia en API",
            "Fuente": "EUDAMED Data Lake"
        }]

    result = []

    for r in found:
        di = str(r.get("PRIMARY_DI") or "").strip()

        result.append({
            **base,
            "UDI-DI": di,
            "Basic UDI-DI": str(
                r.get("BASIC_UDI") or ""
            ).strip(),
            "Fabricante": str(
                r.get("MF_NAME") or ""
            ).strip(),
            "Nombre producto": str(
                r.get("DEVICE_NAME")
                or r.get("TRADE_NAME")
                or ""
            ).strip(),
            "Estado": (
                "Coincidencia exacta"
                if di
                else "Coincidencia sin UDI-DI"
            ),
            "Fuente": "EUDAMED Data Lake"
        })

    return result


def read_file(upload):
    raw = io.BytesIO(upload.getvalue())

    if upload.name.lower().endswith(".csv"):
        df = pd.read_csv(
            raw,
            dtype=str,
            keep_default_na=False,
            sep=None,
            engine="python"
        )
    else:
        df = pd.read_excel(
            raw,
            dtype=str,
            keep_default_na=False
        )

    df.columns = [
        str(c).strip().upper()
        for c in df.columns
    ]

    if "MANUFACT_ITEMNUM" not in df.columns:
        raise ValueError(
            "El archivo debe contener "
            "la columna MANUFACT_ITEMNUM."
        )

    return df


def download_xlsx(df):
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Resultados UDI"
        )

        ws = writer.sheets["Resultados UDI"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = {
            "A": 22,
            "B": 22,
            "C": 25,
            "D": 27,
            "E": 35,
            "F": 35,
            "G": 45,
            "H": 24,
            "I": 24,
            "J": 75
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.number_format = "@"

    return buffer.getvalue()


# COMPROBAR CONEXIÓN

with st.expander(
    "Comprobar conexión con EUDAMED"
):
    test_srn = st.text_input(
        "SRN de prueba",
        "LI-MF-000000522"
    )

    test_ref = st.text_input(
        "Referencia de prueba",
        "577275WW"
    )

    if st.button("Probar API oficial"):
        try:
            data = get_records(
                test_srn.strip().upper(),
                test_ref.strip()
            )

            matches = match_records(
                data,
                test_srn.strip().upper(),
                test_ref.strip()
            )

            if matches:
                st.success(
                    f"Conexión operativa: "
                    f"{len(matches)} coincidencias."
                )

                st.dataframe(
                    pd.DataFrame(
                        make_rows(
                            test_srn.strip().upper(),
                            test_ref.strip(),
                            matches,
                            datetime.now(
                                timezone.utc
                            ).isoformat()
                        )
                    )
                )
            else:
                st.warning(
                    "La API respondió, pero no "
                    "encontró coincidencias exactas."
                )

        except Exception as exc:
            st.error(
                f"Error de conexión: "
                f"{type(exc).__name__}: {exc}"
            )


# CARGAR EXCEL

st.subheader("1. Cargar referencias")

upload = st.file_uploader(
    "Excel o CSV con MANUFACT_ITEMNUM",
    type=["xlsx", "csv"]
)

srns_text = st.text_area(
    "SRN de los fabricantes",
    DEFAULT_SRNS,
    height=105
)

srns = list(dict.fromkeys(
    v.strip().upper()
    for v in re.split(r"[,;\s]+", srns_text)
    if v.strip()
))

invalid = [
    s for s in srns
    if not SRN_RE.fullmatch(s)
]

if invalid:
    st.error(
        "SRN incorrectos: " + ", ".join(invalid)
    )

pairs = []

if upload is not None:
    try:
        source = read_file(upload)

        if "SRN" in source.columns:
            for _, row in source.iterrows():
                ref = str(
                    row["MANUFACT_ITEMNUM"]
                ).strip()

                assigned = str(
                    row["SRN"]
                ).strip().upper()

                if ref:
                    if assigned:
                        pairs.append((ref, assigned))
                    else:
                        pairs.extend(
                            (ref, s) for s in srns
                        )
        else:
            refs = list(dict.fromkeys(
                str(v).strip()
                for v in source["MANUFACT_ITEMNUM"]
                if str(v).strip()
            ))

            pairs = [
                (ref, s)
                for ref in refs
                for s in srns
            ]

        pairs = list(dict.fromkeys(pairs))

        if len(pairs) > 15000:
            raise ValueError(
                "Máximo 15.000 combinaciones."
            )

        st.info(
            f"{len(pairs)} búsquedas preparadas."
        )

    except Exception as exc:
        st.error(f"Error leyendo Excel: {exc}")
        pairs = []


# PROCESAMIENTO AUTOMÁTICO

st.subheader("2. Ejecutar búsqueda automática")

if "results" not in st.session_state:
    st.session_state.results = []
    st.session_state.done = set()
    st.session_state.batch_signature = None

signature = (
    upload.getvalue().__hash__()
    if upload else None,
    tuple(pairs)
)

if st.session_state.batch_signature != signature:
    st.session_state.results = []
    st.session_state.done = set()
    st.session_state.batch_signature = signature

batch_size = st.slider(
    "Consultas por lote",
    1, 100, 30
)

delay = st.slider(
    "Pausa entre consultas (segundos)",
    0.5, 3.0, 0.7, 0.1
)

remaining = [
    p for p in pairs
    if p not in st.session_state.done
]

st.write(
    f"Procesadas: {len(st.session_state.done)} "
    f"· Pendientes: {len(remaining)}"
)

if st.button(
    "▶ Iniciar / continuar búsqueda",
    type="primary",
    disabled=not pairs or bool(invalid)
):
    selected = remaining[:batch_size]

    progress = st.progress(0)
    status = st.empty()

    for i, (ref, srn) in enumerate(selected):
        status.write(
            f"Consultando {i+1}/{len(selected)}: "
            f"{srn} · {ref}"
        )

        try:
            response = get_records(srn, ref)

            if len(response) >= 1000:
                raise ValueError(
                    "Posible truncamiento de resultados."
                )

            exact = match_records(
                response, srn, ref
            )

            now = datetime.now(
                timezone.utc
            ).isoformat(timespec="seconds")

            st.session_state.results.extend(
                make_rows(
                    srn, ref, exact, now
                )
            )

            st.session_state.done.add(
                (ref, srn)
            )

        except Exception as exc:
            st.error(
                "Consulta detenida: "
                f"{type(exc).__name__}: {exc}"
            )
            break

        progress.progress(
            (i + 1) / len(selected)
        )

        if i < len(selected) - 1:
            time.sleep(delay)

    st.info(
        f"Consultas completadas: "
        f"{len(st.session_state.done)} "
        f"de {len(pairs)}."
    )


# RESULTADOS Y DESCARGA

if st.session_state.results:
    results = pd.DataFrame(
        st.session_state.results,
        columns=COLUMNS
    )

    st.subheader("3. Resultados y descarga")

    st.metric(
        "UDI-DI encontrados",
        int(
            (
                results["Estado"]
                == "Coincidencia exacta"
            ).sum()
        )
    )

    st.dataframe(
        results,
        hide_index=True,
        use_container_width=True
    )

    st.download_button(
        "⬇ Descargar Excel con UDI-DI",
        data=download_xlsx(results),
        file_name="UDI_Connect_resultados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

st.caption(
    "Los resultados dependen de la disponibilidad "
    "y cobertura de la API pública de EUDAMED."
)
