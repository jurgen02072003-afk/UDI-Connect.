
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
    layout="wide"
)

st.title("🔎 UDI Connect — búsqueda automática")

st.caption(
    "Carga tu Excel, consulta fabricantes por SRN "
    "y descarga las coincidencias verificadas."
)

API = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"

SRN_RE = re.compile(r"^[A-Z]{2}-MF-\d{9}$")

DEFAULT = """KR-MF-000026311
CN-MF-000009139
JP-MF-000011271"""

COLS = [
    "MANUFACT_ITEMNUM",
    "SRN",
    "UDI-DI",
    "Basic UDI-DI",
    "Fabricante",
    "Nombre comercial",
    "Estado",
    "UUID",
    "Fecha UTC"
]


def norm(x):
    return str(
        x if x is not None else ""
    ).strip().casefold()


def fetch(session, srn, page=0, page_size=300):
    params = {
        "page": page,
        "pageSize": page_size,
        "size": page_size,
        "iso2Code": "en",
        "languageIso2Code": "en",
        "srn": srn
    }

    response = session.get(
        API,
        params=params,
        timeout=40,
        allow_redirects=False,
        headers={
            "Accept": "application/json",
            "User-Agent": "UDIConnect/2.0"
        }
    )

    if response.status_code in (
        301, 302, 303, 307, 308
    ):
        raise RuntimeError(
            "EUDAMED ha redirigido la consulta."
        )

    if response.status_code == 429:
        raise RuntimeError(
            "EUDAMED limita las consultas (429). "
            "Espera unos minutos."
        )

    response.raise_for_status()

    obj = response.json()

    if (
        not isinstance(obj, dict)
        or not isinstance(obj.get("content"), list)
    ):
        raise RuntimeError(
            "Formato inesperado: falta content[]."
        )

    return obj


def read_refs(upload):
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
            "Falta la columna MANUFACT_ITEMNUM."
        )

    return df


def xlsx(data):
    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ) as writer:
        data.to_excel(
            writer,
            index=False,
            sheet_name="Resultados UDI"
        )

        ws = writer.sheets["Resultados UDI"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            letter = col[0].column_letter

            width = max(
                len(str(cell.value or ""))
                for cell in col[:300]
            ) + 2

            ws.column_dimensions[letter].width = min(
                48,
                max(18, width)
            )

            for cell in col[1:]:
                cell.number_format = "@"

    return out.getvalue()


# PRUEBA DE CONEXIÓN

with st.expander(
    "Prueba de conexión",
    expanded=True
):
    test_srn = st.text_input(
        "SRN para la prueba",
        "LI-MF-000000522"
    )

    if st.button(
        "Comprobar respuesta de EUDAMED"
    ):
        try:
            with requests.Session() as session:
                obj = fetch(
                    session,
                    test_srn.strip().upper(),
                    0,
                    25
                )

            total = int(
                obj.get("totalElements", 0)
            )

            matches = sum(
                norm(x.get("manufacturerSrn"))
                == norm(test_srn)
                for x in obj["content"]
            )

            st.write(
                f"Respuesta HTTP correcta. "
                f"Registros declarados: {total:,}. "
                f"Coincidencias de SRN: "
                f"{matches}/{len(obj['content'])}."
            )

            if (
                total > 100000
                or (
                    obj["content"]
                    and matches == 0
                )
            ):
                st.error(
                    "El filtro SRN parece ignorado "
                    "o no está verificado. "
                    "No ejecutes la búsqueda masiva."
                )
            else:
                st.success(
                    "Respuesta compatible con "
                    "una búsqueda acotada. "
                    "Verifica los resultados."
                )

        except Exception as exc:
            st.error(
                "No se pudo validar la conexión: "
                f"{type(exc).__name__}: {exc}"
            )


# CARGA DEL EXCEL

st.subheader("1. Cargar referencias")

srn_text = st.text_area(
    "SRN de fabricantes",
    DEFAULT,
    height=105
)

srns = list(dict.fromkeys(
    value.strip().upper()
    for value in re.split(
        r"[,;\s]+",
        srn_text
    )
    if value.strip()
))

invalid = [
    srn for srn in srns
    if not SRN_RE.fullmatch(srn)
]

if invalid:
    st.error(
        "SRN inválidos: "
        + ", ".join(invalid)
    )

upload = st.file_uploader(
    "Subir Excel/CSV con MANUFACT_ITEMNUM",
    type=["xlsx", "csv"]
)

pairs = []

if upload is not None:
    try:
        source = read_refs(upload)

        for _, row in source.iterrows():
            ref = str(
                row["MANUFACT_ITEMNUM"]
            ).strip()

            if not ref:
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
                pairs.append((ref, srn))

        pairs = list(dict.fromkeys(pairs))

        if len(pairs) > 15000:
            raise ValueError(
                "Máximo 15.000 combinaciones."
            )

        st.info(
            f"{len(pairs)} combinaciones únicas "
            "de referencia y SRN."
        )

    except Exception as exc:
        st.error(
            f"Error al leer el Excel: {exc}"
        )

        pairs = []


# PREPARAR SESIÓN

signature = hashlib.sha256(
    repr(pairs).encode()
).hexdigest()

if st.session_state.get("signature") != signature:
    st.session_state.signature = signature
    st.session_state.results = []
    st.session_state.finished_srns = set()
    st.session_state.errors = {}

max_pages = st.number_input(
    "Límite de páginas por fabricante",
    min_value=1,
    max_value=100,
    value=20
)

st.caption(
    "Si se alcanza el límite, las referencias "
    "no encontradas quedarán como incompletas."
)


# BÚSQUEDA AUTOMÁTICA

st.subheader("2. Buscar UDI-DI")

if st.button(
    "▶ Buscar UDI automáticamente",
    type="primary",
    disabled=not pairs or bool(invalid)
):
    lookup = {}

    for ref, srn in pairs:
        lookup.setdefault(
            srn,
            set()
        ).add(norm(ref))

    pending = [
        srn
        for srn in lookup
        if srn not in st.session_state.finished_srns
    ]

    progress = st.progress(0)

    with requests.Session() as session:
        for index, srn in enumerate(pending):
            st.write(
                f"Consultando fabricante {srn}..."
            )

            found = {}
            completed = False
            error = None

            try:
                first = fetch(
                    session,
                    srn,
                    0
                )

                total = int(
                    first.get("totalElements", 0)
                )

                if total > 100000:
                    raise RuntimeError(
                        "El filtro SRN no es fiable: "
                        f"la API devuelve {total:,} "
                        "dispositivos."
                    )

                total_pages = int(
                    first.get("totalPages", 1)
                )

                pages = min(
                    total_pages,
                    int(max_pages)
                )

                for page in range(pages):
                    obj = (
                        first
                        if page == 0
                        else fetch(
                            session,
                            srn,
                            page
                        )
                    )

                    content = obj["content"]

                    if (
                        content
                        and all(
                            norm(
                                item.get(
                                    "manufacturerSrn"
                                )
                            ) != norm(srn)
                            for item in content
                        )
                    ):
                        raise RuntimeError(
                            "El servidor no está "
                            "aplicando el filtro SRN."
                        )

                    for item in content:
                        if (
                            norm(
                                item.get(
                                    "manufacturerSrn"
                                )
                            ) != norm(srn)
                        ):
                            continue

                        ref = norm(
                            item.get("reference")
                        )

                        if ref in lookup[srn]:
                            found.setdefault(
                                ref,
                                []
                            ).append(item)

                    if page < pages - 1:
                        time.sleep(1.1)

                completed = (
                    pages >= total_pages
                )

            except Exception as exc:
                error = (
                    f"{type(exc).__name__}: "
                    f"{exc}"
                )

            if error:
                st.session_state.errors[srn] = error

                st.error(
                    f"{srn}: {error}"
                )

                break

            now = datetime.now(
                timezone.utc
            ).isoformat(
                timespec="seconds"
            )

            for ref, assigned_srn in pairs:
                if assigned_srn != srn:
                    continue

                hits = found.get(
                    norm(ref),
                    []
                )

                if hits:
                    for item in hits:
                        st.session_state.results.append({
                            "MANUFACT_ITEMNUM": ref,
                            "SRN": srn,
                            "UDI-DI": str(
                                item.get("primaryDi")
                                or ""
                            ),
                            "Basic UDI-DI": str(
                                item.get("basicUdi")
                                or ""
                            ),
                            "Fabricante": str(
                                item.get(
                                    "manufacturerName"
                                ) or ""
                            ),
                            "Nombre comercial": str(
                                item.get("tradeName")
                                or ""
                            ),
                            "Estado":
                                "Coincidencia exacta "
                                "SRN + referencia",
                            "UUID": str(
                                item.get("uuid")
                                or ""
                            ),
                            "Fecha UTC": now
                        })
                else:
                    st.session_state.results.append({
                        "MANUFACT_ITEMNUM": ref,
                        "SRN": srn,
                        "UDI-DI": "",
                        "Basic UDI-DI": "",
                        "Fabricante": "",
                        "Nombre comercial": "",
                        "Estado": (
                            "Sin coincidencia"
                            if completed
                            else "Búsqueda incompleta"
                        ),
                        "UUID": "",
                        "Fecha UTC": now
                    })

            st.session_state.finished_srns.add(
                srn
            )

            progress.progress(
                (index + 1)
                / max(1, len(pending))
            )

    st.success(
        "Fabricantes procesados: "
        f"{len(st.session_state.finished_srns)}"
        f"/{len(lookup)}"
    )


# DESCARGAR RESULTADOS

st.subheader("3. Resultados")

if st.session_state.results:
    results = pd.DataFrame(
        st.session_state.results,
        columns=COLS
    )

    found_count = int(
        (
            results["Estado"]
            == "Coincidencia exacta "
               "SRN + referencia"
        ).sum()
    )

    st.metric(
        "Coincidencias verificadas",
        found_count
    )

    st.dataframe(
        results,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "⬇ Descargar Excel con UDI-DI",
        data=xlsx(results),
        file_name="UDI_Connect_resultados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

else:
    st.info(
        "Todavía no hay resultados. "
        "Prueba primero la conexión."
    )

st.caption(
    "Los datos dependen del servicio público "
    "de EUDAMED. No se generan UDI-DI "
    "sin coincidencias comprobadas."
)
