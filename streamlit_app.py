from collections import Counter
from datetime import datetime
from hashlib import sha256
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from ssco import (
    MAX_ROWS,
    PURCHASE_EXTENSIONS,
    REPORT_HEADERS,
    SUNAT_PAGE,
    SUNAT_XLSX,
    DataError,
    analyze,
    export_report,
    fetch_registry,
    normalize_ruc,
    read_purchases,
    read_registry,
    report_values,
)

LIMA = ZoneInfo("America/Lima")
OFFICIAL = "Descargar de SUNAT"
MANUAL = "Cargar Excel del padrón"
ALL = "Todos"
MATCH = "Coincidencia SSCO"
NO_MATCH = "Sin coincidencia"
PENDING = "Sin RUC consultable"
MAX_BATCH_BYTES = 45 * 1024**2

st.set_page_config(page_title="Consulta SSCO", page_icon="📋", layout="wide")
st.markdown(
    """
    <style>
    .block-container { max-width: 1130px; padding-top: 3.5rem; padding-bottom: 4rem; }
    [data-testid="stSidebarContent"] { border-right: 1px solid #dbe8df; }
    [data-testid="stMetricValue"] { font-variant-numeric: tabular-nums; }
    .ssco-banner {
        background: #143b34;
        color: #fff;
        border-radius: 1.1rem;
        padding: clamp(1.3rem, 4vw, 2.5rem);
        margin-bottom: .9rem;
    }
    .ssco-banner h1 {
        color: inherit;
        font-size: clamp(2rem, 4.5vw, 3rem);
        line-height: 1.08;
        letter-spacing: -.04em;
        margin: 0 0 .55rem;
        padding: 0;
    }
    .ssco-banner p { color: #dcf2e9; margin: 0; font-size: 1.04rem; }
    </style>
    <div class="ssco-banner">
        <h1>Consulta SSCO</h1>
        <p>Verifica si tus proveedores figuran en la relación oficial de SUNAT.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=4 * 60 * 60, show_spinner=False, max_entries=1)
def official_registry():
    return fetch_registry()


def today():
    return datetime.now(LIMA)


def format_date(value):
    return value.strftime("%d/%m/%Y") if value else "Sin fecha informada"


def show_registry(registry, consulted):
    st.caption(
        f"{registry.source} · Última publicación en el padrón {format_date(registry.latest_publication)} "
        f"· Consulta {consulted}"
    )


def upload_id(file):
    return file.file_id if file else None


with st.sidebar:
    st.header("Padrón SSCO")
    mode = st.radio("Fuente", [OFFICIAL, MANUAL])
    registry_file = st.file_uploader("Excel oficial del padrón", type=["xlsx"]) if mode == MANUAL else None
    if mode == OFFICIAL:
        st.caption("La app descarga la relación oficial de SUNAT al realizar la consulta.")
        if st.button("Volver a descargar en la próxima consulta"):
            official_registry.clear()
            st.session_state["registry_revision"] = st.session_state.get("registry_revision", 0) + 1
    st.link_button("Ver publicación de SUNAT", SUNAT_PAGE)
    if mode == MANUAL:
        st.link_button("Descargar Excel oficial", SUNAT_XLSX)

registry_key = (mode, upload_id(registry_file), st.session_state.get("registry_revision", 0))


def get_registry():
    if mode == MANUAL:
        if not registry_file:
            raise DataError("Carga el padrón oficial en el panel lateral.")
        return read_registry(registry_file.getvalue(), f"Excel cargado · {registry_file.name}")
    return official_registry()


def run_batch(files):
    if sum(file.size for file in files) > MAX_BATCH_BYTES:
        raise DataError("La carga supera los 45 MB en total. Divide los archivos en varias consultas.")
    purchases, files_report = [], []
    seen = set()
    name_counts = Counter(file.name for file in files)
    processed_names = Counter()
    for file in files:
        processed_names[file.name] += 1
        source = file.name
        if name_counts[file.name] > 1:
            source = f"{file.name} ({processed_names[file.name]})"
        data = file.getvalue()
        digest = sha256(data).digest()
        if digest in seen:
            files_report.append((source, "Copia idéntica omitida", 0))
            continue
        seen.add(digest)
        try:
            rows = read_purchases(data, file.name, source)
        except DataError as exc:
            files_report.append((source, str(exc), 0))
            continue
        if len(purchases) + len(rows) > MAX_ROWS:
            raise DataError(f"La carga supera los {MAX_ROWS:,} registros. Divide los archivos en varias consultas.")
        purchases.extend(rows)
        files_report.append((source, "Leído" if rows else "Sin comprobantes", len(rows)))
    if not purchases:
        return [], None, None, files_report
    registry = get_registry()
    consulted = today().strftime("%d/%m/%Y %H:%M")
    return analyze(purchases, registry), registry, consulted, files_report


def show_table(reviews):
    columns = REPORT_HEADERS[:15] + ("Archivo", "Fila")
    selections = [REPORT_HEADERS.index(name) for name in columns]
    rows = [
        [row[index] for index in selections]
        for row in map(report_values, reviews)
    ]
    table = pd.DataFrame(rows, columns=columns)
    for col in ("Total informado", "IGV informado"):
        table[col] = table[col].map(lambda value: float(value) if value is not None else None)
    st.dataframe(
        table,
        hide_index=True,
        width="stretch",
        column_config={
            "Emisión": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Publicación SSCO": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Resolución firme": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Total informado": st.column_config.NumberColumn(format="localized", step=0.01),
            "IGV informado": st.column_config.NumberColumn(format="localized", step=0.01),
        },
    )


def show_batch(result):
    reviews, registry, consulted, files_report = result
    failed = [(name, status) for name, status, _ in files_report
              if status not in {"Leído", "Copia idéntica omitida", "Sin comprobantes"}]
    if failed:
        st.warning(f"Quedaron {len(failed)} archivo(s) sin procesar. Revisa el detalle de la carga.")
    with st.expander("Detalle de la carga", expanded=bool(failed)):
        st.dataframe(pd.DataFrame(files_report, columns=["Archivo", "Estado", "Filas"]),
                     hide_index=True, width="stretch")
    if not reviews:
        st.error("No se encontraron comprobantes para cruzar en los archivos cargados.")
        return
    show_registry(registry, consulted)
    dates = [review.purchase.issued for review in reviews if review.purchase.issued]
    if dates:
        st.caption(f"Emisiones del {format_date(min(dates))} al {format_date(max(dates))}")
    matches = [review for review in reviews if review.subject]
    pending = [review for review in reviews if not review.purchase.ruc]
    providers = {review.purchase.ruc for review in matches}
    counts = st.columns(4)
    counts[0].metric("Filas leídas", len(reviews), border=True)
    counts[1].metric("Filas con coincidencia", len(matches), border=True)
    counts[2].metric("Proveedores SSCO", len(providers), border=True)
    counts[3].metric("Filas sin RUC consultable", len(pending), border=True)
    if matches:
        st.warning("Los proveedores que figuran en el padrón aparecen marcados para revisión contable.")
    elif len(pending) == len(reviews):
        st.warning("Las filas leídas carecen de un RUC consultable.")
    else:
        st.success("Los RUC consultados no figuran en este padrón SSCO.")
    if pending:
        st.caption("Las filas sin RUC consultable incluyen otros tipos de documento o identificadores incompletos.")
    if len(files_report) > 1:
        st.caption("Si el mismo comprobante aparece en formatos distintos, se muestra una fila por cada origen.")
    filter_options = [MATCH, NO_MATCH, PENDING, ALL] if matches else [ALL, MATCH, NO_MATCH, PENDING]
    view = st.radio("Mostrar", filter_options, horizontal=True, key="report_filter")
    selected = reviews if view == ALL else [review for review in reviews if review.status == view]
    if selected:
        show_table(selected)
    else:
        st.info("No hay filas con este estado.")
    st.download_button(
        "Descargar reporte Excel",
        data=lambda: export_report(reviews, registry, consulted, files_report),
        file_name=f"cruce_ssco_{today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click="ignore",
    )


batch_tab, manual_tab = st.tabs(["Cruzar compras", "Buscar por RUC"])

with batch_tab:
    st.subheader("Carga tus archivos de compras")
    files = st.file_uploader(
        "Registro contable, resultado de NewContaSis o propuesta SIRE",
        type=[ext[1:] for ext in PURCHASE_EXTENSIONS],
        accept_multiple_files=True,
        help="Admite COMPRAS.xlsx, resultado_*.xlsx y propuestas SUNAT en ZIP, TXT o CSV.",
    )
    batch_key = (tuple(upload_id(file) for file in files), registry_key)
    if st.session_state.get("batch_key") != batch_key:
        st.session_state.pop("batch_result", None)
        st.session_state.pop("report_filter", None)
        st.session_state["batch_key"] = batch_key
    if st.button("Cruzar con el padrón", type="primary", disabled=not files):
        st.session_state.pop("batch_result", None)
        st.session_state.pop("report_filter", None)
        try:
            with st.spinner("Leyendo archivos y verificando el padrón SSCO…"):
                st.session_state["batch_result"] = run_batch(files)
        except DataError as exc:
            st.error(str(exc))
    if "batch_result" in st.session_state:
        show_batch(st.session_state["batch_result"])

with manual_tab:
    st.subheader("Consulta un proveedor")
    with st.form("ruc_search"):
        requested = st.text_input("RUC del proveedor", placeholder="Ingresa los 11 dígitos", max_chars=24)
        search = st.form_submit_button("Consultar RUC", type="primary")
    if st.session_state.get("search_registry") != registry_key or search:
        st.session_state.pop("search_result", None)
    if search:
        ruc = normalize_ruc(requested)
        if not ruc:
            st.error("Ingresa un RUC de 11 dígitos.")
        else:
            try:
                with st.spinner("Consultando el padrón SSCO…"):
                    registry = get_registry()
                st.session_state["search_result"] = (
                    ruc, registry, today().strftime("%d/%m/%Y %H:%M")
                )
                st.session_state["search_registry"] = registry_key
            except DataError as exc:
                st.error(str(exc))
    if "search_result" in st.session_state:
        ruc, registry, consulted = st.session_state["search_result"]
        show_registry(registry, consulted)
        subjects = registry.subjects.get(ruc, ())
        if subjects:
            st.warning(f"El RUC {ruc} figura en el padrón SSCO.")
            st.dataframe(
                [{
                    "Razón social": subject.name,
                    "Resolución": subject.resolution,
                    "Fecha de publicación": format_date(subject.published),
                    "Resolución firme": format_date(subject.final),
                } for subject in subjects],
                hide_index=True, width="stretch",
            )
        else:
            st.success(f"El RUC {ruc} no figura en el padrón consultado.")

st.divider()
st.caption("Se admiten compras de cualquier año. El resultado corresponde al padrón consultado.")
