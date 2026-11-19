
import io
import re
import time
import hashlib
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="UDI Connect",
    page_icon="🔎",
    layout="centered"
)

st.title("🔎 UDI Connect")
st.caption("SRN · Excel · Búsqueda automática · Descarga")

API = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"

SRN_PATTERN = re.compile(r"^[A-Z]{2}-MF-\d{9}$")

COLUMNS = [
    "MANUFACT_ITEMNUM",
    "SRN",
    "UDI-DI",
    "Basic UDI-DI",
    "Estado",
    "Fecha UTC"
]


def norm(value):
    return str(
        value if value is not None else ""
    ).strip().casefold()


def read_excel(upload):
    data = io.BytesIO(upload.getvalue())

    if upload.name.lower().endswith(".csv"):
        df = pd.read_csv(
            data,
            dtype=str,
            keep_default_na=False,
            sep=None,
            engine="python"
        )
    else:
        df = pd.read_excel(
            data,
            dtype=str,
            keep_default_na=False
        )

    df.columns = [
        str(c).strip().upper()
        for c in df.columns
    ]

    if "MANUFACT_ITEMNUM" not in df.columns:
        raise ValueError(
            "El Excel debe contener la columna MANUFACT_ITEMNUM."
        )

    return df


def fetch(session, srn, page, page_size=100):
    response = session.get(
        API,
        params={
            "page": page,
            "pageSize": page_size,
            "size": page_size,
            "iso2Code": "en",
            "languageIso2Code": "en",
            "srn": srn
        },
        timeout=35,
        headers={
            "Accept": "application/json",
            "User-Agent": "UDIConnect/3.0"
        }
    )

    if response.status_code == 429:
        retry_after = response.headers.get("Retry-After")

        message = (
            f"Reintenta después de {retry_after}."
            if retry_after
            else "Espera antes de reintentar."
        )

        raise RuntimeError(
            "EUDAMED ha limitado las consultas (429). "
            + message
        )

    response.raise_for_status()
    body = response.json()

    if (
        not isinstance(body, dict)
        or not isinstance(body.get("content"), list)
    ):
        raise RuntimeError(
            "EUDAMED devolvió un formato inesperado."
        )

    return body


def excel_bytes(rows):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:
        pd.DataFrame(
            rows,
            columns=COLUMNS
        ).to_excel(
            writer,
            sheet_name="Resultados UDI",
            index=False
        )

        sheet = writer.sheets["Resultados UDI"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        for column in sheet.columns:
            sheet.column_dimensions[
                column[0].column_letter
            ].width = 26

            for cell in column[1:]:
                cell.number_format = "@"

    return output.getvalue()


# 1. SRN

st.subheader("1. SRN del fabricante")

srn_text = st.text_area(
    "Introduce uno o varios SRN (uno por línea)",
    value="",
    placeholder="",
    height=105
)

srns = list(dict.fromkeys(
    s.strip().upper()
    for s in re.split(r"[,;\s]+", srn_text)
    if s.strip()
))

invalid_srns = [
    s for s in srns
    if not SRN_PATTERN.fullmatch(s)
]

if invalid_srns:
    st.error(
        "SRN no válidos: "
        + ", ".join(invalid_srns)
    )


# 2. ADJUNTAR DOCUMENTO

st.subheader("2. Adjuntar documento")

upload = st.file_uploader(
    "Excel o CSV",
    type=["xlsx", "csv"],
    label_visibility="collapsed"
)

pairs = []

if upload is not None:
    try:
        source = read_excel(upload)

        for _, row in source.iterrows():
            reference = str(
                row["MANUFACT_ITEMNUM"]
            ).strip()

            if not reference:
                continue

            row_srn = str(
                row.get("SRN", "")
            ).strip().upper()

            assigned_srns = (
                [row_srn]
                if row_srn
                else srns
            )

            for srn in assigned_srns:
                pairs.append((reference, srn))

        pairs = list(dict.fromkeys(pairs))

        if len(pairs) > 15000:
            raise ValueError(
                "Máximo 15.000 combinaciones."
            )

        if any(
            not SRN_PATTERN.fullmatch(srn)
            for _, srn in pairs
        ):
            raise ValueError(
                "Hay un SRN incorrecto dentro del Excel."
            )

        st.caption(
            f"{len(pairs)} combinaciones preparadas."
        )

    except Exception as exc:
        st.error(str(exc))
        pairs = []


# CONSERVAR EL PROGRESO

signature = hashlib.sha256(
    repr(pairs).encode("utf-8")
).hexdigest()

if st.session_state.get("search_signature") != signature:
    st.session_state.search_signature = signature
    st.session_state.results = []
    st.session_state.completed = set()
    st.session_state.current_pages = {}
    st.session_state.matches = {}
    st.session_state.expected_pages = {}


# 3. BÚSQUEDA AUTOMÁTICA

st.subheader(
    "3. Búsqueda de UDI-DI y Basic UDI-DI"
)

if st.button(
    "🔎 Buscar automáticamente",
    type="primary",
    use_container_width=True,
    disabled=not pairs or bool(invalid_srns)
):
    groups = {}

    for ref, srn in pairs:
        groups.setdefault(
            srn,
            set()
        ).add(norm(ref))

    with requests.Session() as session:
        for srn, refs in groups.items():
            if srn in st.session_state.completed:
                continue

            page = st.session_state.current_pages.get(
                srn, 0
            )

            st.write(
                f"Buscando fabricante {srn}..."
            )

            while page < 20:
                try:
                    body = fetch(
                        session,
                        srn,
                        page
                    )

                    content = body["content"]
                    total_pages = body.get("totalPages")

                    if total_pages is None:
                        raise RuntimeError(
                            "No se puede confirmar "
                            "una búsqueda completa."
                        )

                    total_pages = int(total_pages)

                    st.session_state.expected_pages[
                        srn
                    ] = total_pages

                    total = int(
                        body.get("totalElements", 0)
                    )

                    if total > 100000:
                        raise RuntimeError(
                            "El filtro SRN parece "
                            "no estar aplicado."
                        )

                    if content and not any(
                        norm(
                            item.get("manufacturerSrn")
                        ) == norm(srn)
                        for item in content
                        if isinstance(item, dict)
                    ):
                        raise RuntimeError(
                            "La respuesta no confirma "
                            "el SRN solicitado."
                        )

                    for item in content:
                        if not isinstance(item, dict):
                            continue

                        if norm(
                            item.get("manufacturerSrn")
                        ) != norm(srn):
                            continue

                        ref = norm(
                            item.get("reference")
                        )

                        if ref not in refs:
                            continue

                        di = str(
                            item.get("primaryDi") or ""
                        ).strip()

                        basic = str(
                            item.get("basicUdi") or ""
                        ).strip()

                        if di:
                            key = (srn, ref, di)

                            st.session_state.matches[
                                key
                            ] = basic

                    page += 1

                    st.session_state.current_pages[
                        srn
                    ] = page

                    if page >= total_pages:
                        st.session_state.completed.add(
                            srn
                        )
                        break

                    time.sleep(1.5)

                except Exception as exc:
                    st.warning(
                        f"Búsqueda pausada para "
                        f"{srn}: {exc}"
                    )
                    break

            else:
                st.info(
                    f"{srn}: límite de 20 páginas. "
                    "Puedes descargar resultados parciales."
                )

            if srn not in st.session_state.completed:
                break

    now = datetime.now(
        timezone.utc
    ).isoformat(timespec="seconds")

    rows = []

    for ref, srn in pairs:
        matches = [
            (di, basic)
            for (s, r, di), basic
            in st.session_state.matches.items()
            if s == srn and r == norm(ref)
        ]

        if matches:
            for di, basic in matches:
                rows.append({
                    "MANUFACT_ITEMNUM": ref,
                    "SRN": srn,
                    "UDI-DI": di,
                    "Basic UDI-DI": basic,
                    "Estado":
                        "Coincidencia exacta SRN + referencia",
                    "Fecha UTC": now
                })
        else:
            rows.append({
                "MANUFACT_ITEMNUM": ref,
                "SRN": srn,
                "UDI-DI": "",
                "Basic UDI-DI": "",
                "Estado": (
                    "Sin coincidencia"
                    if srn in st.session_state.completed
                    else "Pendiente / búsqueda incompleta"
                ),
                "Fecha UTC": now
            })

    st.session_state.results = rows


# 4. DESCARGAR EXCEL

st.subheader("4. Descargar documento")

if st.session_state.results:
    results = pd.DataFrame(
        st.session_state.results,
        columns=COLUMNS
    )

    st.dataframe(
        results,
        hide_index=True,
        use_container_width=True
    )

    st.download_button(
        "⬇ Descargar Excel",
        data=excel_bytes(
            st.session_state.results
        ),
        file_name="UDI_Connect_resultados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

else:
    st.caption(
        "Los resultados aparecerán aquí "
        "cuando ejecutes una búsqueda."
    )

st.caption(
    "La búsqueda depende de EUDAMED. "
    "Un resultado pendiente no significa "
    "que el producto no exista."
)
