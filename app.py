
"""UDI Connect: búsqueda en EUDAMED y registro de UDI-DI."""

import io
import re
import json
import hashlib
from datetime import datetime, timezone
from urllib.parse import urlencode

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="UDI Connect | EUDAMED",
    page_icon="🔎",
    layout="wide"
)

st.title("🔎 UDI Connect")
st.caption(
    "Busca por SRN y referencia en EUDAMED "
    "y guarda los UDI-DI encontrados"
)

FIELDS = [
    "MANUFACT_ITEMNUM",
    "SRN",
    "UDI-DI",
    "Basic UDI-DI",
    "Master UDI-DI",
    "Estado",
    "Enlace EUDAMED",
    "Fecha UTC"
]


def srn_list(raw):
    values = [
        v.strip().upper()
        for v in re.split(r"[,;\s]+", raw)
        if v.strip()
    ]
    return list(dict.fromkeys(values))


def url_for(srn, ref):
    params = {
        "srn": srn,
        "reference": ref,
        "deviceStatusCode":
            "refdata.device-model-status.on-the-market",
        "paging": json.dumps(
            {"pageIndex": 0, "pageSize": 25},
            separators=(",", ":")
        ),
        "sorting": json.dumps(
            {
                "sortField": "primaryDi",
                "sortDirection": "desc"
            },
            separators=(",", ":")
        ),
        "submitted": "true"
    }

    return (
        "https://ec.europa.eu/tools/eudamed/"
        "#/screen/search-device?"
        + urlencode(params)
    )


def export_excel(df, fingerprint):
    out = io.BytesIO()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            sheet_name="Resultados",
            index=False
        )

        pd.DataFrame([{
            "Huella": fingerprint,
            "Exportado UTC":
                datetime.now(timezone.utc).isoformat()
        }]).to_excel(
            writer,
            sheet_name="Control",
            index=False
        )

        ws = writer.sheets["Resultados"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.number_format = "@"

    return out.getvalue()


# BÚSQUEDA INDIVIDUAL

st.subheader("1. Buscar un producto")

with st.form("direct_search"):
    left, right = st.columns(2)

    srn_direct = left.text_input(
        "SRN",
        placeholder="KR-MF-000026311"
    ).strip().upper()

    ref_direct = right.text_input(
        "Referencia del fabricante"
    ).strip()

    search_direct = st.form_submit_button(
        "🔎 Preparar búsqueda",
        type="primary"
    )

if search_direct:
    if (
        not re.fullmatch(
            r"[A-Z]{2}-(MF|PR)-\d{9}",
            srn_direct
        )
        or not ref_direct
    ):
        st.error(
            "Introduce un SRN válido y una referencia."
        )
    else:
        st.session_state.direct_url = url_for(
            srn_direct,
            ref_direct
        )

if "direct_url" in st.session_state:
    st.link_button(
        "Abrir búsqueda en EUDAMED",
        st.session_state.direct_url,
        type="primary"
    )

st.divider()

# BÚSQUEDA MEDIANTE EXCEL

st.subheader("2. Buscar referencias de un Excel")

srns_raw = st.text_area(
    "SRN de los fabricantes",
    value=(
        "KR-MF-000026311\n"
        "CN-MF-000009139\n"
        "JP-MF-000011271"
    ),
    height=100
)

srns = srn_list(srns_raw)

invalid = [
    s for s in srns
    if not re.fullmatch(
        r"[A-Z]{2}-(MF|PR)-\d{9}",
        s
    )
]

if invalid:
    st.error(
        "SRN inválidos: " + ", ".join(invalid)
    )
    st.stop()

if not srns:
    st.warning("Introduce al menos un SRN.")
    st.stop()

upload = st.file_uploader(
    "Sube Excel o CSV con MANUFACT_ITEMNUM",
    type=["xlsx", "csv"]
)

restore = st.file_uploader(
    "Recuperar resultados anteriores (opcional)",
    type=["xlsx"],
    key="restore"
)

if upload is None:
    st.info(
        "También puedes buscar productos individuales "
        "sin subir un Excel."
    )
    st.stop()

try:
    raw = io.BytesIO(upload.getvalue())

    if upload.name.lower().endswith(".csv"):
        source = pd.read_csv(
            raw,
            dtype=str,
            keep_default_na=False,
            sep=None,
            engine="python"
        )
    else:
        source = pd.read_excel(
            raw,
            dtype=str,
            keep_default_na=False
        )

    cols = {
        str(c).strip().upper(): c
        for c in source.columns
    }

    if "MANUFACT_ITEMNUM" not in cols:
        raise ValueError(
            "Falta la columna MANUFACT_ITEMNUM"
        )

    refs = list(dict.fromkeys(
        str(v).strip()
        for v in source[cols["MANUFACT_ITEMNUM"]]
        if str(v).strip()
    ))

    if not refs:
        raise ValueError(
            "No hay referencias en MANUFACT_ITEMNUM"
        )

    if len(refs) > 5000:
        raise ValueError(
            "Máximo 5.000 referencias únicas"
        )

    if "SRN" in cols:
        pairs = list(dict.fromkeys(
            (
                str(row[cols["MANUFACT_ITEMNUM"]]).strip(),
                str(row[cols["SRN"]]).strip().upper()
            )
            for _, row in source.iterrows()
            if str(row[cols["MANUFACT_ITEMNUM"]]).strip()
        ))

        pairs = [
            (ref, srn)
            for ref, srn in pairs
            if srn
        ] + [
            (ref, srn)
            for ref, old_srn in pairs
            if not old_srn
            for srn in srns
        ]
    else:
        pairs = [
            (ref, srn)
            for ref in refs
            for srn in srns
        ]

    bad = [
        srn for _, srn in pairs
        if not re.fullmatch(
            r"[A-Z]{2}-(MF|PR)-\d{9}",
            srn
        )
    ]

    if bad:
        raise ValueError(
            "SRN inválidos: "
            + ", ".join(sorted(set(bad)))
        )

except Exception as exc:
    st.error(
        f"No se pudo leer el archivo: {exc}"
    )
    st.stop()


# CREAR SESIÓN DE RESULTADOS

fingerprint = hashlib.sha256(
    json.dumps(
        pairs,
        ensure_ascii=False
    ).encode()
).hexdigest()

if st.session_state.get("fingerprint") != fingerprint:
    st.session_state.fingerprint = fingerprint

    st.session_state.results = pd.DataFrame(
        [
            {
                "MANUFACT_ITEMNUM": ref,
                "SRN": srn,
                "UDI-DI": "",
                "Basic UDI-DI": "",
                "Master UDI-DI": "",
                "Estado": "Pendiente de consulta",
                "Enlace EUDAMED": url_for(srn, ref),
                "Fecha UTC": ""
            }
            for ref, srn in pairs
        ],
        columns=FIELDS
    )

    st.session_state.restored_hash = None


# RESTAURAR PROGRESO

if restore is not None:
    digest = hashlib.sha256(
        restore.getvalue()
    ).hexdigest()

    if st.session_state.restored_hash != digest:
        try:
            restored = pd.read_excel(
                io.BytesIO(restore.getvalue()),
                sheet_name="Resultados",
                dtype=str,
                keep_default_na=False
            )

            control = pd.read_excel(
                io.BytesIO(restore.getvalue()),
                sheet_name="Control",
                dtype=str,
                keep_default_na=False
            )

            if (
                control.empty
                or control.iloc[0]["Huella"]
                != fingerprint
            ):
                raise ValueError(
                    "El progreso corresponde "
                    "a otra búsqueda"
                )

            if not set(FIELDS).issubset(
                restored.columns
            ):
                raise ValueError(
                    "Faltan columnas de resultados"
                )

            st.session_state.results = (
                restored[FIELDS].copy()
            )

            st.session_state.restored_hash = digest

            st.success("Resultados recuperados")

        except Exception as exc:
            st.error(
                f"No se pudo recuperar: {exc}"
            )


# PANEL DE RESULTADOS

df = st.session_state.results

st.caption(
    f"{len(refs)} referencias únicas · "
    f"{len(srns)} SRN · "
    f"{len(df)} combinaciones"
)

found = int(
    (df["Estado"] == "Verificado manualmente").sum()
)

a, b, c = st.columns(3)

a.metric("Combinaciones", len(df))
b.metric("UDI-DI guardados", found)
c.metric("Pendientes", len(df) - found)

st.subheader(
    "3. Buscar en EUDAMED y guardar el resultado"
)

show_pending = st.checkbox(
    "Mostrar solo pendientes",
    value=True
)

choices = [
    i for i in range(len(df))
    if (
        not show_pending
        or df.at[i, "Estado"]
        != "Verificado manualmente"
    )
]

if choices:
    chosen = st.selectbox(
        "Selecciona la referencia",
        options=choices,
        format_func=lambda i: (
            f"Fila {i + 1} · "
            f"{df.at[i, 'MANUFACT_ITEMNUM']} · "
            f"{df.at[i, 'SRN']}"
        )
    )

    ref = df.at[chosen, "MANUFACT_ITEMNUM"]
    srn = df.at[chosen, "SRN"]

    st.write(
        f"**SRN:** `{srn}` · "
        f"**Referencia:** `{ref}`"
    )

    st.link_button(
        "🔎 Buscar esta referencia en EUDAMED",
        df.at[chosen, "Enlace EUDAMED"],
        type="primary"
    )

    with st.form(f"save_{chosen}"):
        udi = st.text_input(
            "UDI-DI encontrado",
            value=str(df.at[chosen, "UDI-DI"])
        )

        basic = st.text_input(
            "Basic UDI-DI (opcional)",
            value=str(df.at[chosen, "Basic UDI-DI"])
        )

        master = st.text_input(
            "Master UDI-DI (opcional)",
            value=str(df.at[chosen, "Master UDI-DI"])
        )

        save = st.form_submit_button(
            "Guardar resultado"
        )

    if save:
        if not udi.strip():
            st.error(
                "Introduce un UDI-DI antes de guardar."
            )
        else:
            df.at[chosen, "UDI-DI"] = udi.strip()
            df.at[chosen, "Basic UDI-DI"] = (
                basic.strip()
            )
            df.at[chosen, "Master UDI-DI"] = (
                master.strip()
            )

            df.at[chosen, "Estado"] = (
                "Verificado manualmente"
            )

            df.at[chosen, "Fecha UTC"] = (
                datetime.now(
                    timezone.utc
                ).isoformat()
            )

            st.session_state.results = df

            st.success(
                "Resultado guardado. "
                "Descarga el Excel para conservarlo."
            )
else:
    st.success(
        "Todas las combinaciones tienen "
        "un UDI-DI guardado."
    )


# EXPORTACIÓN

st.download_button(
    "⬇ Descargar Excel con resultados",
    data=export_excel(df, fingerprint),
    file_name="UDI_Connect_resultados.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True
)

st.dataframe(
    df,
    hide_index=True,
    use_container_width=True,
    column_config={
        "Enlace EUDAMED":
            st.column_config.LinkColumn(
                "Abrir en EUDAMED",
                display_text="Abrir búsqueda"
            )
    }
)

st.caption(
    "Los enlaces abren el buscador oficial de "
    "EUDAMED. Los UDI-DI se introducen "
    "manualmente tras verificarlos."
)
