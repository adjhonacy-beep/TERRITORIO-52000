
import re
import html
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

# ------------------------------------------------------------
# ASISTENTE TERRITORIAL DE PASTO - JC
# Búsqueda basada exclusivamente en información del Excel.
# ------------------------------------------------------------

st.set_page_config(
    page_title="Asistente Territorial",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PRIMARY_FIELDS = [
    "ZONA", "COMUNA", "BARRIO",
    "CORREGIMIENTO", "VEREDA", "SECTOR"
]

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


# ------------------------------------------------------------
# NORMALIZACIÓN Y LIMPIEZA
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# CARGA DEL EXCEL
# CONSERVA TODAS LAS COLUMNAS ORIGINALES
# ------------------------------------------------------------

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

    # Conservar todas las columnas del archivo
    data = raw.copy()

    # Normalizar los nombres de las columnas reconocidas
    rename_map = {
        original: field
        for field, original in mapping.items()
    }

    data = data.rename(columns=rename_map)

    # Evitar columnas duplicadas por los cambios de nombre
    data = data.loc[:, ~data.columns.duplicated()]

    # Limpiar todas las celdas del Excel
    for col in data.columns:
        data[col] = data[col].map(clean_cell)

    # Eliminar filas completamente vacías
    data = data.loc[
        ~data.apply(
            lambda row: all(not str(v).strip() for v in row),
            axis=1
        )
    ]

    return (
        data.drop_duplicates().reset_index(drop=True),
        source_name,
        mapping
    )


# ------------------------------------------------------------
# VALORES ÚNICOS Y SIN CELDAS VACÍAS
# ------------------------------------------------------------

def unique_nonempty(values):

    found = {}

    for value in values:

        value = clean_cell(value)
        key = normalize(value)

        if key and key not in found:
            found[key] = value

    return sorted(found.values(), key=normalize)


# ------------------------------------------------------------
# MOTOR DE BÚSQUEDA
# ------------------------------------------------------------

def search_records(question, data):

    q = normalize(question)

    if not q:
        return data.iloc[0:0], "", 0

    # --------------------------------------------------------
    # BÚSQUEDA RURAL
    # --------------------------------------------------------

    if q in {
        "rural",
        "zona rural",
        "corregimientos",
        "corregimiento"
    }:

        mask = pd.Series(False, index=data.index)

        if "ZONA" in data.columns:
            mask |= data["ZONA"].map(
                normalize
            ).str.contains("rural", na=False)

        if "CORREGIMIENTO" in data.columns:
            mask |= data["CORREGIMIENTO"].map(
                lambda x: bool(clean_cell(x))
            )

        rows = data[mask].drop_duplicates().reset_index(drop=True)

        if rows.empty:
            return data.iloc[0:0], "", 0

        return rows, "Zona Rural / Todos los corregimientos de Pasto", 100

    # --------------------------------------------------------
    # BÚSQUEDA URBANA
    # --------------------------------------------------------

    if q in {
        "urbana",
        "urbano",
        "zona urbana",
        "comunas",
        "comuna"
    }:

        mask = pd.Series(False, index=data.index)

        if "ZONA" in data.columns:
            mask |= data["ZONA"].map(
                normalize
            ).str.contains("urban", na=False)

        if "COMUNA" in data.columns:
            mask |= data["COMUNA"].map(
                lambda x: bool(clean_cell(x))
            )

        rows = data[mask].drop_duplicates().reset_index(drop=True)

        if rows.empty:
            return data.iloc[0:0], "", 0

        return rows, "Zona Urbana / Todas las comunas de Pasto", 100

    # --------------------------------------------------------
    # COMPARACIÓN DE COINCIDENCIAS
    # --------------------------------------------------------

    def match_score(query, value):

        nq = normalize(query)
        nv = normalize(value)

        if not nq or not nv:
            return 0

        nq_alt = re.sub(
            r"^(comuna|comunas)\s*", "", nq
        ).strip()

        nv_alt = re.sub(
            r"^(comuna|comunas)\s*", "", nv
        ).strip()

        if nq in (nv, nv_alt) or nv in (nq, nq_alt):
            return 100

        if nq in nv or nv in nq:
            return 95

        if nq_alt and nq_alt in nv_alt:
            return 95

        return max(
            fuzz.ratio(nq, nv),
            fuzz.ratio(nq_alt, nv_alt) if nq_alt else 0
        )

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

        normalized_values = {
            normalize(value)
            for value, score in matches
        }

        return (
            data[
                data[field].map(normalize).isin(normalized_values)
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

    # --------------------------------------------------------
    # PRIORIDAD DE BÚSQUEDA TERRITORIAL
    # --------------------------------------------------------

    for priority_field in [
        "ZONA",
        "COMUNA",
        "CORREGIMIENTO",
        "BARRIO",
        "VEREDA",
        "SECTOR",
        "EDIFICIO"
    ]:

        matches = find_matches(priority_field, q)

        if matches:

            rows = get_related(priority_field, matches)

            nombres = ", ".join(
                value for value, score in matches[:10]
            )

            if len(matches) > 10:
                nombres += f" y {len(matches) - 10} más"

            return (
                rows,
                nombres,
                max(score for value, score in matches)
            )

    # --------------------------------------------------------
    # BÚSQUEDA AMPLIA EN TODAS LAS COLUMNAS
    # --------------------------------------------------------

    all_matches = []

    tokens = [
        token for token in q.split()
        if len(token) >= 3
    ]

    for field in data.columns:

        for value in unique_nonempty(data[field].tolist()):

            nv = normalize(value)

            if nv and any(token in nv for token in tokens):
                all_matches.append((field, value, 78))

    if not all_matches:
        return data.iloc[0:0], "", 0

    result_frames = []

    for field, value, score in all_matches:

        matched_rows = data[
            data[field].map(normalize) == normalize(value)
        ]

        result_frames.append(matched_rows)

    rows = (
        pd.concat(result_frames, ignore_index=True)
        .drop_duplicates()
        .reset_index(drop=True)
    )

    names = unique_nonempty([
        value for field, value, score in all_matches
    ])

    matched_name = ", ".join(names[:10])

    if len(names) > 10:
        matched_name += f" y {len(names) - 10} más"

    best_score = max(
        score for field, value, score in all_matches
    )

    return rows, matched_name, int(best_score)


# ------------------------------------------------------------
# PRESENTACIÓN DE VALORES
# ------------------------------------------------------------

def format_field_values(values, preview_limit=5):

    # Evitar repetidos y celdas vacías
    values = unique_nonempty(values)

    # Escapar contenido HTML proveniente del Excel
    safe_values = [
        html.escape(value)
        for value in values
    ]

    if len(safe_values) <= preview_limit:
        return "<br>".join(safe_values)

    visible = "<br>".join(
        safe_values[:preview_limit]
    )

    hidden = "<br>".join(
        safe_values[preview_limit:]
    )

    remaining = len(safe_values) - preview_limit

    return (
        f"{visible}"
        f"<details class='more-details'>"
        f"<summary class='more-btn'>"
        f"+ Ver {remaining} más"
        f"</summary>"
        f"<div class='more-content'>{hidden}</div>"
        f"</details>"
    )


# ------------------------------------------------------------
# PRESENTACIÓN DE RESULTADOS
# CONSOLIDA INFORMACIÓN DE TODAS LAS FILAS ENCONTRADAS
# ------------------------------------------------------------

def render_result(rows, matched_name, score):

    rows = rows.drop_duplicates().reset_index(drop=True)

    safe_name = html.escape(str(matched_name))

    st.markdown(
        f'<div class="match-heading">'
        f'Coincidencia encontrada: <b>{safe_name}</b>'
        f'<span class="match-score">{score}%</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        f"{len(rows)} filas relacionadas encontradas en el Excel. "
        "Se excluyen celdas vacías y se unifican los datos repetidos."
    )

    # Orden de presentación
    ordered_fields = (
        [f for f in PRIMARY_FIELDS if f in rows.columns]
        + [f for f in EXTRA_FIELDS if f in rows.columns]
        + [
            f for f in rows.columns
            if f not in PRIMARY_FIELDS + EXTRA_FIELDS
        ]
    )

    grouped = {}

    for field in ordered_fields:

        values = unique_nonempty(rows[field].tolist())

        if values:
            grouped[field] = values

    if not grouped:
        st.warning("No hay información disponible para mostrar.")
        return

    st.markdown(
        '<div class="result-card">',
        unsafe_allow_html=True
    )

    main_fields = [
        (k, v)
        for k, v in grouped.items()
        if k in PRIMARY_FIELDS
    ]

    extra_fields = [
        (k, v)
        for k, v in grouped.items()
        if k not in PRIMARY_FIELDS
    ]

    # Información territorial
    cols = st.columns(3)

    for i, (field, values) in enumerate(main_fields):

        value_html = format_field_values(
            values,
            preview_limit=5
        )

        with cols[i % 3]:

            st.markdown(
                f'<div class="field">'
                f'<div class="field-label">'
                f'{html.escape(str(field.title()))} '
                f'<span class="count-badge">({len(values)})</span>'
                f'</div>'
                f'<div class="field-value">{value_html}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    # Información complementaria
    if extra_fields:

        st.markdown(
            '<div class="extra-title">'
            'Información complementaria'
            '</div>',
            unsafe_allow_html=True,
        )

        extra_cols = st.columns(3)

        for i, (field, values) in enumerate(extra_fields):

            value_html = format_field_values(
                values,
                preview_limit=5
            )

            with extra_cols[i % 3]:

                st.markdown(
                    f'<div class="field">'
                    f'<div class="field-label">'
                    f'{html.escape(str(field.title()))} '
                    f'<span class="count-badge">({len(values)})</span>'
                    f'</div>'
                    f'<div class="field-value">{value_html}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    st.markdown('</div>', unsafe_allow_html=True)


# ------------------------------------------------------------
# DISEÑO VISUAL
# ------------------------------------------------------------

st.markdown("""
<style>

#MainMenu, footer {
    visibility: hidden;
}

.block-container {
    padding-top: 2.5rem !important;
    max-width: 900px;
}

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

div[data-testid="stFormSubmitButton"] button {
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
    padding: 10px 12px;
    margin: 5px 0;
    background: #f7f9fc;
    border-radius: 9px;
    min-height: 68px;
}

.field-label {
    font-size: 11px;
    color: #6c7788;
    text-transform: uppercase;
    letter-spacing: .6px;
    font-weight: 700;
}

.count-badge {
    color: #17365d;
    font-weight: 800;
}

.field-value {
    font-size: 15px;
    color: #172b4d;
    font-weight: 600;
    overflow-wrap: anywhere;
    margin-top: 5px;
    line-height: 1.45;
}

.more-details {
    margin-top: 6px;
}

.more-btn {
    display: inline-block;
    cursor: pointer;
    font-size: 12px;
    font-weight: 700;
    color: #17365d;
    background: #e3ecf7;
    padding: 3px 10px;
    border-radius: 12px;
    user-select: none;
    list-style: none;
    transition: background 0.2s;
}

.more-btn::-webkit-details-marker {
    display: none;
}

.more-btn:hover {
    background: #d0e0f3;
}

.more-details[open] .more-btn {
    background: #17365d;
    color: #ffffff;
    margin-bottom: 6px;
}

.more-content {
    margin-top: 4px;
    padding-top: 4px;
    border-top: 1px dashed #cbd5e1;
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

</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------
# ENCABEZADO
# ------------------------------------------------------------

st.markdown("""
<div class="header-container">
    <div class="map-icon">📍</div>
    <div class="brand-title">ASISTENTE TERRITORIAL</div>
    <div class="brand-sub">Pasto - Nariño</div>
</div>
""", unsafe_allow_html=True)


# ------------------------------------------------------------
# CARGA DE INFORMACIÓN
# ------------------------------------------------------------

uploaded = None

try:
    df, source_name, column_map = load_excel(uploaded)

except Exception as exc:

    st.error(
        f"No fue posible leer el archivo Excel: {exc}"
    )

    st.stop()


if df is None:

    st.info(
        "Para comenzar, coloca el archivo Pasto.xlsx "
        "junto a app.py."
    )

    st.stop()


# ------------------------------------------------------------
# BÚSQUEDA CON ENTER O BOTÓN
# ------------------------------------------------------------

st.markdown(
    '<div class="search-label">'
    '¿Qué lugar o nombre deseas consultar?'
    '</div>',
    unsafe_allow_html=True
)

with st.form(key="territorial_search_form"):

    c1, c2 = st.columns([5, 1])

    with c1:

        query = st.text_input(
            "Buscar",
            placeholder=(
                "Ejemplo: rural, comunas, El Tejar, comuna 5..."
            ),
            label_visibility="collapsed",
            key="territorial_query",
        )

    with c2:

        search_clicked = st.form_submit_button(
            "Buscar",
            type="primary",
            use_container_width=True
        )


# ------------------------------------------------------------
# EJECUCIÓN DE LA BÚSQUEDA
# ENTER O BOTÓN BUSCAR
# ------------------------------------------------------------

if search_clicked:

    if query.strip():

        results, matched_name, score = search_records(
            query,
            df
        )

        if len(results):

            render_result(
                results,
                matched_name,
                score
            )

        else:

            st.warning(
                "No encontré una coincidencia suficientemente "
                "clara en el Excel."
            )

            st.caption(
                "Verifica la escritura o intenta con una "
                "parte más corta del nombre."
            )

    else:

        st.info(
            "Escribe un nombre en el recuadro superior "
            "para iniciar la búsqueda."
        )