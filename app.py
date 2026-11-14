import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

# ------------------------------------------------------------
# ASISTENTE - JC
# Consulta basada en coincidencias con el Excel; no inventa datos.
# ------------------------------------------------------------

st.set_page_config(
    page_title="Asistente Territorial",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PRIMARY_FIELDS = ["ZONA", "COMUNA", "BARRIO", "CORREGIMIENTO", "VEREDA", "SECTOR"]
EXTRA_FIELDS = ["EDIFICIO"]
FIELD_ALIASES = {
    "ZONA": ["ZONA"],
    "COMUNA": ["COMUNA", "COMUNAS"],
    "BARRIO": ["BARRIO", "BARRIOS"],
    "CORREGIMIENTO": ["CORREGIMIENTO", "CORREGIMIENTOS"],
    "VEREDA": ["VEREDA", "VEREDAS"],
    "SECTOR": ["SECTOR", "SECTORES"],
    "EDIFICIO": ["EDIFICIO", "EDIFICIOS", "CONJUNTO", "UNIDAD"],
}

def normalize(value):
    if pd.isna(value):
        return ""
    value = str(value).strip().lower()
    value = "".join(
        c for c in unicodedata.normalize("NFD", value)
        if unicodedata.category(c) != "Mn"
    )
    value = re.sub(r"[^a-z0-9#]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()

def clean_cell(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    if value.lower() in {"nan", "none", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value

def resolve_columns(frame):
    mapping = {}
    normalized = {}
    for col in frame.columns:
        key = normalize(col).replace(" ", "")
        normalized.setdefault(key, col)
    for target, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            key = normalize(alias).replace(" ", "")
            if key in normalized:
                mapping[target] = normalized[key]
                break
    return mapping

def load_excel(uploaded_file=None):
    path = Path("Pasto.xlsx")
    if uploaded_file is not None:
        raw = pd.read_excel(uploaded_file, dtype=str)
        source_name = uploaded_file.name
    elif path.exists():
        raw = pd.read_excel(path, dtype=str)
        source_name = path.name
    else:
        return None, None, None

    mapping = resolve_columns(raw)
    if not any(field in mapping for field in PRIMARY_FIELDS):
        raise ValueError(
            "No se reconocieron las columnas territoriales. "
            "Verifica los encabezados del Excel."
        )

    data = pd.DataFrame()
    for field, original in mapping.items():
        data[field] = raw[original].map(clean_cell)

    available = [f for f in PRIMARY_FIELDS + EXTRA_FIELDS if f in data.columns]
    data = data[available]
    data = data.loc[~data.apply(lambda row: all(not str(v).strip() for v in row), axis=1)]
    return data.reset_index(drop=True), source_name, mapping

def unique_nonempty(values):
    found = {}
    for value in values:
        value = clean_cell(value)
        key = normalize(value)
        if key and key not in found:
            found[key] = value
    return sorted(found.values(), key=normalize)

def search_records(question, data):
    q = normalize(question)

    if not q:
        return data.iloc[0:0], "", 0

    def match_score(query, value):
        nq = normalize(query)
        nv = normalize(value)

        if not nq or not nv:
            return 0

        nq_alt = re.sub(r"^(comuna|comunas)\s*", "", nq).strip()
        nv_alt = re.sub(r"^(comuna|comunas)\s*", "", nv).strip()

        if nq in (nv, nv_alt) or nv in (nq, nq_alt):
            return 100
        if nq in nv or nv in nq:
            return 95
        if nq_alt and nq_alt in nv_alt:
            return 95

        return max(fuzz.ratio(nq, nv), fuzz.ratio(nq_alt, nv_alt) if nq_alt else 0)

    def find_matches(field, query, threshold=76):
        if field not in data.columns:
            return []
        matches = []
        for value in unique_nonempty(data[field].tolist()):
            score = match_score(query, value)
            if score >= threshold:
                matches.append((value, score))
        return matches

    def get_related(field, matches):
        if field not in data.columns or not matches:
            return data.iloc[0:0]
        normalized_values = {normalize(value) for value, score in matches}
        return data[data[field].map(normalize).isin(normalized_values)].drop_duplicates().reset_index(drop=True)

    if q in {"rural", "zona rural"}:
        matches = unique_nonempty(data["CORREGIMIENTO"].tolist()) if "CORREGIMIENTO" in data.columns else []
        if not matches:
            return data.iloc[0:0], "", 0
        result = pd.DataFrame({"CORREGIMIENTO": matches})
        return result, "Todos los corregimientos de Pasto", 100

    if q in {"urbana", "urbano", "zona urbana"}:
        matches = unique_nonempty(data["COMUNA"].tolist()) if "COMUNA" in data.columns else []
        if not matches:
            return data.iloc[0:0], "", 0
        result = pd.DataFrame({"COMUNA": matches})
        return result, "Todas las comunas de Pasto", 100

    comuna_matches = find_matches("COMUNA", q)
    if comuna_matches:
        rows = get_related("COMUNA", comuna_matches)
        nombres = ", ".join(value for value, score in comuna_matches)
        return rows, nombres, max(score for value, score in comuna_matches)

    corregimiento_matches = find_matches("CORREGIMIENTO", q)
    if corregimiento_matches:
        rows = get_related("CORREGIMIENTO", corregimiento_matches)
        nombres = ", ".join(value for value, score in corregimiento_matches)
        return rows, nombres, max(score for value, score in corregimiento_matches)

    barrio_matches = find_matches("BARRIO", q)
    if barrio_matches:
        rows = get_related("BARRIO", barrio_matches)
        nombres = ", ".join(value for value, score in barrio_matches)
        return rows, nombres, max(score for value, score in barrio_matches)

    all_matches = []
    for field in data.columns:
        matches = find_matches(field, q)
        for value, score in matches:
            all_matches.append((field, value, score))

    if not all_matches:
        tokens = [token for token in q.split() if len(token) >= 3]
        for field in data.columns:
            for value in unique_nonempty(data[field].tolist()):
                nv = normalize(value)
                if nv and any(token in nv for token in tokens):
                    all_matches.append((field, value, 78))

    if not all_matches:
        return data.iloc[0:0], "", 0

    result_frames = []
    for field, value, score in all_matches:
        matched_rows = data[data[field].map(normalize) == normalize(value)]
        result_frames.append(matched_rows)

    rows = pd.concat(result_frames, ignore_index=True).drop_duplicates().reset_index(drop=True)
    names = unique_nonempty([value for field, value, score in all_matches])
    matched_name = ", ".join(names[:10])

    if len(names) > 10:
        matched_name += f" y {len(names) - 10} más"

    best_score = max(score for field, value, score in all_matches)
    return rows, matched_name, int(best_score)

def render_result(rows, matched_name, score):
    st.markdown(
        f'<div class="match-heading">Coincidencia encontrada: <b>{matched_name}</b>'
        f'<span class="match-score">{score}%</span></div>',
        unsafe_allow_html=True,
    )
    if len(rows) > 1:
        st.caption(f"{len(rows)} registros relacionados encontrados en el archivo. Los valores repetidos se agrupan.")

    grouped = {}
    for field in PRIMARY_FIELDS + EXTRA_FIELDS:
        if field in rows.columns:
            vals = unique_nonempty(rows[field].tolist())
            if vals:
                grouped[field] = vals

    st.markdown('<div class="result-card">', unsafe_allow_html=True)
    main = [(k, v) for k, v in grouped.items() if k in PRIMARY_FIELDS]
    extra = [(k, v) for k, v in grouped.items() if k in EXTRA_FIELDS]
    cols = st.columns(3)
    for i, (field, values) in enumerate(main):
        value_html = "<br>".join(values) if len(values) <= 8 else "<br>".join(values[:8]) + f"<br><span class='note'>y {len(values)-8} más…</span>"
        with cols[i % 3]:
            st.markdown(
                f'<div class="field"><div class="field-label">{field.title()}</div>'
                f'<div class="field-value">{value_html}</div></div>',
                unsafe_allow_html=True,
            )
    if extra:
        st.markdown('<div class="extra-title">Información complementaria</div>', unsafe_allow_html=True)
        extra_cols = st.columns(2)
        for i, (field, values) in enumerate(extra):
            value_html = "<br>".join(values) if len(values) <= 8 else "<br>".join(values[:8]) + f"<br><span class='note'>y {len(values)-8} más…</span>"
            with extra_cols[i % 2]:
                st.markdown(
                    f'<div class="field"><div class="field-label">{field.title()}</div>'
                    f'<div class="field-value">{value_html}</div></div>',
                    unsafe_allow_html=True,
                )
    st.markdown('</div>', unsafe_allow_html=True)

# Estilo visual centrado y limpio
st.markdown("""
<style>
#MainMenu, footer {visibility:hidden;}
.block-container {padding-top:2.5rem !important; max-width:900px;}

.header-container {
    text-align: center;
    margin-top: 15px;
    margin-bottom: 25px;
}
.map-icon {
    font-size: 55px;
    line-height: 1;
    margin-bottom: 10px;
}
.brand-title {
    font-size: 28px;
    font-weight: 800;
    color: #17365d;
    margin-bottom: 2px;
    letter-spacing: 0.5px;
}
.brand-sub {
    font-size: 15px;
    color: #687386;
    font-weight: 500;
}
.search-label {
    text-align: center;
    font-size: 15px;
    font-weight: 600;
    color: #222222;
    margin-top: 8px;
    margin-bottom: 20px;
}
div[data-testid="stTextInput"] input {
    border: none !important;
    box-shadow: none !important;
    font-size: 17px;
    background: transparent !important;
}
div[data-testid="stTextInput"] div {
    background: transparent !important;
    border: none !important;
}
div.stButton > button {
    border-radius: 10px;
    height: 48px;
    font-weight: 700;
    width: 100%;
}
.result-card {
    background: #fff;
    border: 1px solid #e5eaf0;
    border-radius: 15px;
    padding: 17px 18px;
    margin: 8px 0 15px;
    box-shadow: 0 3px 12px rgba(25,45,75,.045);
}
.field {
    padding: 8px 9px;
    margin: 4px 0;
    background: #f7f9fc;
    border-radius: 9px;
    min-height: 64px;
}
.field-label {
    font-size: 11px;
    color: #6c7788;
    text-transform: uppercase;
    letter-spacing: .6px;
    font-weight: 700;
}
.field-value {
    font-size: 15px;
    color: #172b4d;
    font-weight: 650;
    overflow-wrap: anywhere;
    margin-top: 4px;
}
.extra-title {
    font-weight: 750;
    color: #244b70;
    border-top: 1px solid #e9edf2;
    padding-top: 12px;
    margin-top: 12px;
}
.match-heading {
    font-size: 18px;
    color: #1d3557;
    margin: 16px 0 6px;
}
.match-score {
    float: right;
    color: #16734a;
    font-size: 12px;
    background: #e8f6ef;
    padding: 4px 9px;
    border-radius: 20px;
}
.note {
    font-size: 12px;
    color: #718096;
}
</style>
""", unsafe_allow_html=True)

# Encabezado centrado con el Pin y subtítulo Pasto - Nariño
st.markdown("""
<div class="header-container">
  <div class="map-icon">📍</div>
  <div class="brand-title">ASISTENTE TERRITORIAL</div>
  <div class="brand-sub">Pasto - Nariño</div>
</div>
""", unsafe_allow_html=True)

uploaded = None
try:
    df, source_name, column_map = load_excel(uploaded)
except Exception as exc:
    st.error(f"No fue posible leer el archivo Excel: {exc}")
    st.stop()

if df is None:
    st.info("Para comenzar, coloca el archivo Pasto.xlsx junto a app.py o cárgalo aquí.")
    st.stop()

# Campo de búsqueda y etiqueta ordenados
c1, c2 = st.columns([5, 1])
with c1:
    query = st.text_input(
        "Buscar",
        placeholder="Ejemplo: El Tejar, comuna 5, Catambuco...",
        label_visibility="collapsed",
        key="territorial_query",
    )
with c2:
    search_clicked = st.button("Buscar", type="primary", use_container_width=True)

st.markdown('<div class="search-label">¿Qué lugar o nombre deseas consultar?</div>', unsafe_allow_html=True)

if search_clicked or query:
    if query.strip():
        results, matched_name, score = search_records(query, df)
        if len(results):
            render_result(results, matched_name, score)
        else:
            st.warning("No encontré una coincidencia suficientemente clara en el Excel.")
            st.caption("Verifica la escritura o intenta con una parte más corta del nombre.")
    else:
        st.info("Escribe un nombre en el recuadro superior para iniciar la búsqueda.")
