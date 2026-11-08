import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
st.markdown("""<style>.block-container{padding-top: 4rem !important;}</style>""",unsafe_allow_html=True)
from rapidfuzz import fuzz, process

# ------------------------------------------------------------
# ASISTENTE - JC
# Consulta basada en coincidencias con el Excel; no inventa datos.
# ------------------------------------------------------------

st.set_page_config(
    page_title="Asistente JC",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PRIMARY_FIELDS = ["ZONA", "COMUNA", "BARRIO", "CORREGIMIENTO", "VEREDA", "SECTOR"]
EXTRA_FIELDS = ["EDIFICIO", "DIRECCION", "LOCAL", "TELEFONO"]
FIELD_ALIASES = {
    "ZONA": ["ZONA"],
    "COMUNA": ["COMUNA", "COMUNAS"],
    "BARRIO": ["BARRIO", "BARRIOS"],
    "CORREGIMIENTO": ["CORREGIMIENTO", "CORREGIMIENTOS"],
    "VEREDA": ["VEREDA", "VEREDAS"],
    "SECTOR": ["SECTOR", "SECTORES"],
    "EDIFICIO": ["EDIFICIO", "EDIFICIOS", "CONJUNTO", "UNIDAD"],
    "DIRECCION": ["DIRECCION", "DIRECCIÓN", "DIRECCION DEL PREDIO", "DIRECCIÓN DEL PREDIO"],
    "LOCAL": ["LOCAL", "LOCALES", "ESTABLECIMIENTO"],
    "TELEFONO": ["TELEFONO", "TELÉFONO", "CELULAR", "CONTACTO", "TELEFONOS", "TELÉFONOS"],
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
    # Evita mostrar números telefónicos como 3001234567.0
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

    # Conserva únicamente campos con contenido y elimina filas totalmente vacías.
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

    # Busca coincidencia por nombre en todos los campos disponibles.
    candidates = []
    for field in data.columns:
        for value in unique_nonempty(data[field].tolist()):
            nv = normalize(value)
            if not nv:
                continue
            if nv == q:
                score = 100
            elif nv in q:
                score = min(99, 92 + len(nv) / max(len(q), 1) * 7)
            elif q in nv and len(q) >= 3:
                score = 88
            else:
                score = fuzz.ratio(q, nv)
            if score >= 76:
                candidates.append((score, field, value, nv))

    if not candidates:
        # Si la consulta incluye una frase, intenta comparar sus palabras principales.
        tokens = [t for t in q.split() if len(t) >= 3]
        if tokens:
            for field in data.columns:
                for value in unique_nonempty(data[field].tolist()):
                    nv = normalize(value)
                    if nv and any(t in nv for t in tokens):
                        candidates.append((78, field, value, nv))

    if not candidates:
        return data.iloc[0:0], "", 0

    candidates.sort(key=lambda x: (x[0], len(x[3])), reverse=True)
    best_score, best_field, best_value, best_norm = candidates[0]

    # Coincidencia con el valor más relevante; evita mezclar registros de nombres distintos.
    mask = data[best_field].map(normalize) == best_norm
    rows = data.loc[mask].copy()

    # Si la búsqueda es una comuna y hay varias filas, se presentan sus registros asociados.
    return rows.drop_duplicates().reset_index(drop=True), best_value, int(best_score)

def display_fields(row):
    order = [f for f in PRIMARY_FIELDS + EXTRA_FIELDS if f in row.index]
    return [(field, clean_cell(row[field])) for field in order if clean_cell(row[field])]

def render_result(rows, matched_name, score):
    st.markdown(
        f'<div class="match-heading">Coincidencia encontrada: <b>{matched_name}</b>'
        f'<span class="match-score">{score}%</span></div>',
        unsafe_allow_html=True,
    )
    if len(rows) > 1:
        st.caption(f"{len(rows)} registros relacionados encontrados en el archivo. Los valores repetidos se agrupan.")

    # Agrupa valores únicos por campo para evitar decenas de tarjetas repetidas
    # al consultar una comuna o un corregimiento.
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

# Estilo visual: buscador superior y marca JC a la derecha.
st.markdown("""
<style>
#MainMenu, footer {visibility:hidden;}
.block-container {padding-top:1.5rem; max-width:1180px;}
.topline {display:flex; align-items:center; justify-content:space-between; gap:18px;
          border-bottom:1px solid #e7eaf0; padding-bottom:16px; margin-bottom:18px;}
.brand-title {font-size:27px; font-weight:800; color:#17365d; line-height:1.15;}
.brand-sub {font-size:13px; color:#687386; margin-top:5px;}
.jc-logo {width:62px; height:62px; border-radius:16px; display:flex; align-items:center;
          justify-content:center; background:linear-gradient(135deg,#17365d,#247ba0);
          color:white; font-size:24px; font-weight:900; letter-spacing:1px;
          box-shadow:0 5px 16px rgba(23,54,93,.18); flex-shrink:0;}
.search-label {font-size:15px; font-weight:700; color:#24364b; margin:10px 0 5px;}
div[data-testid="stTextInput"] input {height:52px; border-radius:12px; font-size:17px;}
div.stButton > button {border-radius:10px; height:48px; font-weight:700;}
.result-card {background:#fff; border:1px solid #e5eaf0; border-radius:15px;
              padding:17px 18px; margin:8px 0 15px; box-shadow:0 3px 12px rgba(25,45,75,.045);}
.field {padding:8px 9px; margin:4px 0; background:#f7f9fc; border-radius:9px; min-height:64px;}
.field-label {font-size:11px; color:#6c7788; text-transform:uppercase; letter-spacing:.6px; font-weight:700;}
.field-value {font-size:15px; color:#172b4d; font-weight:650; overflow-wrap:anywhere; margin-top:4px;}
.extra-title {font-weight:750; color:#244b70; border-top:1px solid #e9edf2; padding-top:12px; margin-top:12px;}
.match-heading {font-size:18px; color:#1d3557; margin:16px 0 6px;}
.match-score {float:right; color:#16734a; font-size:12px; background:#e8f6ef; padding:4px 9px; border-radius:20px;}
.note {font-size:12px; color:#718096;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="topline">
  <div>
    <div class="brand-title">Asistente Territorial de Pasto</div>
    <div class="brand-sub">Consulta organizada de información territorial por coincidencia de nombre</div>
  </div>
  <div class="jc-logo">JC</div>
</div>
""", unsafe_allow_html=True)

#uploaded = none
try:
    df, source_name, column_map = load_excel(uploaded)
except Exception as exc:
    st.error(f"No fue posible leer el archivo Excel: {exc}")
    st.stop()

if df is None:
    st.info("Para comenzar, coloca el archivo Pasto.xlsx junto a app.py o cárgalo aquí.")
    st.stop()

st.markdown('<div class="search-label">¿Qué lugar o nombre deseas consultar?</div>', unsafe_allow_html=True)
c1, c2 = st.columns([5, 1])
with c1:
    query = st.text_input(
        "Buscar",
        placeholder="Ejemplo: El Tejar, comuna 5, Catambuco, una dirección o un teléfono…",
        label_visibility="collapsed",
        key="territorial_query",
    )
with c2:
    search_clicked = st.button("Buscar", type="primary", use_container_width=True)

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

with st.expander("Ejemplos de búsqueda"):
    st.write("• El Tejar")
    st.write("• Comuna 5")
    st.write("• Catambuco")
    st.write("• Nombre de una vereda o sector")
    st.write("• Dirección, edificio, local o teléfono (si esos datos están incluidos en el Excel)")

with st.expander("Estado y cobertura de la base"):
    st.write(f"**Archivo consultado:** {source_name}")
    st.write(f"**Registros con información:** {len(df)}")
    available_primary = [f for f in PRIMARY_FIELDS if f in df.columns]
    available_extra = [f for f in EXTRA_FIELDS if f in df.columns]
    st.write("**Campos territoriales:** " + (", ".join(available_primary) if available_primary else "No identificados"))
    st.write("**Campos complementarios disponibles:** " + (", ".join(available_extra) if available_extra else "Ninguno"))
    missing_extra = [f for f in EXTRA_FIELDS if f not in df.columns]
    if missing_extra:
        st.warning(
            "El Excel actual no contiene columnas reconocibles para: "
            + ", ".join(missing_extra)
            + ". La aplicación las mostrará automáticamente cuando se agreguen al archivo."
        )
    st.caption("La aplicación muestra únicamente datos existentes en el Excel. No completa ni inventa información.")
