import re
import html
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------

st.set_page_config(
    page_title="Asistente JC",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PRIMARY_FIELDS = [
    "ZONA", "COMUNA", "BARRIO",
    "CORREGIMIENTO", "VEREDA", "SECTOR"
]

EXTRA_FIELDS = [
    "EDIFICIO", "DIRECCION", "LOCAL", "TELEFONO"
]

FIELD_ALIASES = {
    "ZONA": ["ZONA"],
    "COMUNA": ["COMUNA", "COMUNAS"],
    "BARRIO": ["BARRIO", "BARRIOS"],
    "CORREGIMIENTO": ["CORREGIMIENTO", "CORREGIMIENTOS"],
    "VEREDA": ["VEREDA", "VEREDAS"],
    "SECTOR": ["SECTOR", "SECTORES"],
    "EDIFICIO": ["EDIFICIO", "EDIFICIOS", "CONJUNTO", "UNIDAD"],
    "DIRECCION": [
        "DIRECCION", "DIRECCIÓN",
        "DIRECCION DEL PREDIO", "DIRECCIÓN DEL PREDIO"
    ],
    "LOCAL": ["LOCAL", "LOCALES", "ESTABLECIMIENTO"],
    "TELEFONO": [
        "TELEFONO", "TELÉFONO", "CELULAR",
        "CONTACTO", "TELEFONOS", "TELÉFONOS"
    ],
}

# ------------------------------------------------------------
# NORMALIZACIÓN DE INFORMACIÓN
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


def unique_nonempty(values):
    found = {}

    for value in values:
        value = clean_cell(value)
        key = normalize(value)

        if key and key not in found:
            found[key] = value

    return sorted(found.values(), key=normalize)


# ------------------------------------------------------------
# IDENTIFICACIÓN DE COLUMNAS
# ------------------------------------------------------------

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

    data = pd.DataFrame()

    for field, original in mapping.items():
        data[field] = raw[original].map(clean_cell)

    available = [
        f for f in PRIMARY_FIELDS + EXTRA_FIELDS
        if f in data.columns
    ]

    data = data[available]

    data = data.loc[
        ~data.apply(
            lambda row: all(not str(v).strip() for v in row),
            axis=1
        )
    ]

    return data.reset_index(drop=True), source_name, mapping


# ------------------------------------------------------------
# MOTOR DE BÚSQUEDA
# ------------------------------------------------------------

def search_records(question, data):

    q = normalize(question)

    if not q:
        return data.iloc[0:0], "", 0

    # Comparación de nombres y teléfonos.
    def match_score(query, value, field=""):

        nq = normalize(query)
        nv = normalize(value)

        if not nq or not nv:
            return 0

        # TELÉFONOS: solamente coincidencias reales.
        if field == "TELEFONO":

            qdigits = re.sub(r"\D", "", nq)
            vdigits = re.sub(r"\D", "", nv)

            if len(qdigits) < 7:
                return 0

            if qdigits == vdigits:
                return 100

            if qdigits in vdigits:
                return 95

            return 0

        # Permite buscar 5 y encontrar Comuna 5.
        nq_alt = re.sub(
            r"^comunas?\s*", "", nq
        ).strip()

        nv_alt = re.sub(
            r"^comunas?\s*", "", nv
        ).strip()

        # Coincidencia exacta.
        if nq == nv or nq == nv_alt or nq_alt == nv:
            return 100

        # Coincidencia por fragmento.
        if len(nq) >= 3 and nq in nv:
            return 96

        if len(nq_alt) >= 3 and nq_alt in nv_alt:
            return 96

        # Coincidencia aproximada.
        return max(
            fuzz.ratio(nq, nv),
            fuzz.token_set_ratio(nq, nv)
        )

    # Busca todos los valores coincidentes.
    def find_matches(field, query, threshold=82):

        if field not in data.columns:
            return []

        matches = []

        for value in unique_nonempty(data[field].tolist()):

            score = match_score(query, value, field)

            if score >= threshold:
                matches.append((value, score))

        # Exactas primero y aproximadas después.
        return sorted(
            matches,
            key=lambda item: (-item[1], normalize(item[0]))
        )

    # Recupera todos los registros relacionados.
    def get_related(field, matches):

        if field not in data.columns or not matches:
            return data.iloc[0:0]

        values = {
            normalize(value)
            for value, score in matches
        }

        return (
            data[
                data[field].map(normalize).isin(values)
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

    # --------------------------------------------------------
    # 1. BÚSQUEDA URBANA
    # --------------------------------------------------------

    if q in {"urbana", "urbano", "zona urbana"}:

        if "COMUNA" not in data.columns:
            return data.iloc[0:0], "", 0

        comunas = unique_nonempty(data["COMUNA"].tolist())

        if not comunas:
            return data.iloc[0:0], "", 0

        result = pd.DataFrame({"COMUNA": comunas})

        return result, "Todas las comunas de Pasto", 100

    # --------------------------------------------------------
    # 2. BÚSQUEDA RURAL
    # --------------------------------------------------------

    if q in {"rural", "zona rural"}:

        if "CORREGIMIENTO" not in data.columns:
            return data.iloc[0:0], "", 0

        rural_rows = data[
            data["CORREGIMIENTO"].map(normalize) != ""
        ].copy()

        if rural_rows.empty:
            return data.iloc[0:0], "", 0

        corregimientos = unique_nonempty(
            rural_rows["CORREGIMIENTO"].tolist()
        )

        return (
            rural_rows.drop_duplicates().reset_index(drop=True),
            f"Todos los corregimientos de Pasto ({len(corregimientos)})",
            100
        )

    # --------------------------------------------------------
    # 3. BÚSQUEDA POR COMUNA
    # --------------------------------------------------------

    comuna_matches = find_matches("COMUNA", q)

    if comuna_matches:

        rows = get_related("COMUNA", comuna_matches)

        nombres = ", ".join(
            value for value, score in comuna_matches
        )

        return (
            rows,
            nombres,
            max(score for value, score in comuna_matches)
        )

    # --------------------------------------------------------
    # 4. BÚSQUEDA POR CORREGIMIENTO
    # --------------------------------------------------------

    corregimiento_matches = find_matches("CORREGIMIENTO", q)

    if corregimiento_matches:

        rows = get_related(
            "CORREGIMIENTO",
            corregimiento_matches
        )

        nombres = ", ".join(
            value for value, score in corregimiento_matches
        )

        return (
            rows,
            nombres,
            max(score for value, score in corregimiento_matches)
        )

    # --------------------------------------------------------
    # 5. BÚSQUEDA POR BARRIO
    # --------------------------------------------------------

    barrio_matches = find_matches("BARRIO", q)

    if barrio_matches:

        rows = get_related("BARRIO", barrio_matches)

        nombres = ", ".join(
            value for value, score in barrio_matches
        )

        return (
            rows,
            nombres,
            max(score for value, score in barrio_matches)
        )

    # --------------------------------------------------------
    # 6. BÚSQUEDA GENERAL
    # --------------------------------------------------------

    all_matches = []

    for field in data.columns:

        matches = find_matches(field, q)

        for value, score in matches:
            all_matches.append((field, value, score))

    # Búsqueda por palabras cuando no hay coincidencias.
    if not all_matches:

        tokens = [
            token for token in q.split()
            if len(token) >= 3
        ]

        for field in data.columns:

            if field == "TELEFONO":
                continue

            for value in unique_nonempty(data[field].tolist()):

                nv = normalize(value)

                if nv and any(
                    token in nv.split()
                    for token in tokens
                ):
                    all_matches.append((field, value, 78))

    if not all_matches:
        return data.iloc[0:0], "", 0

    # Recupera todos los registros coincidentes.
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

    matched_name = ", ".join(names)

    best_score = max(
        score for field, value, score in all_matches
    )

    return rows, matched_name, int(best_score)


# ------------------------------------------------------------
# PRESENTACIÓN DE RESULTADOS
# ------------------------------------------------------------

def render_result(rows, matched_name, score):

    safe_name = html.escape(str(matched_name))

    st.markdown(
        f'<div class="match-heading">'
        f'Coincidencia encontrada: <b>{safe_name}</b>'
        f'<span class="match-score">{score}%</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        f"{len(rows)} registros relacionados encontrados. "
        "Se muestran todos los valores disponibles."
    )

    grouped = {}

    for field in PRIMARY_FIELDS + EXTRA_FIELDS:

        if field in rows.columns:

            values = unique_nonempty(rows[field].tolist())

            if values:
                grouped[field] = values

    st.markdown(
        '<div class="result-card">',
        unsafe_allow_html=True
    )

    # Campos territoriales.
    main = [
        (k, v)
        for k, v in grouped.items()
        if k in PRIMARY_FIELDS
    ]

    # Campos complementarios.
    extra = [
        (k, v)
        for k, v in grouped.items()
        if k in EXTRA_FIELDS
    ]

    cols = st.columns(3)

    for i, (field, values) in enumerate(main):

        value_html = "<br>".join(
            html.escape(value)
            for value in values
        )

        with cols[i % 3]:

            st.markdown(
                f'<div class="field">'
                f'<div class="field-label">{field.title()}</div>'
                f'<div class="field-value">{value_html}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    if extra:

        st.markdown(
            '<div class="extra-title">'
            'Información complementaria'
            '</div>',
            unsafe_allow_html=True,
        )

        extra_cols = st.columns(2)

        for i, (field, values) in enumerate(extra):

            value_html = "<br>".join(
                html.escape(value)
                for value in values
            )

            with extra_cols[i % 2]:

                st.markdown(
                    f'<div class="field">'
                    f'<div class="field-label">{field.title()}</div>'
                    f'<div class="field-value">{value_html}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    st.markdown('</div>', unsafe_allow_html=True)


# ------------------------------------------------------------
# DISEÑO VISUAL JC
# ------------------------------------------------------------

st.markdown("""
<style>

#MainMenu, footer {
    visibility: hidden;
}

.block-container {
    padding-top: 1.5rem;
    max-width: 1180px;
}

.topline {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
    border-bottom: 1px solid #e7eaf0;
    padding-bottom: 16px;
    margin-bottom: 18px;
}

.brand-title {
    font-size: 27px;
    font-weight: 800;
    color: #17365d;
    line-height: 1.15;
}

.brand-sub {
    font-size: 13px;
    color: #687386;
    margin-top: 5px;
}

.jc-logo {
    width: 62px;
    height: 62px;
    border-radius: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg,#17365d,#247ba0);
    color: white;
    font-size: 24px;
    font-weight: 900;
    letter-spacing: 1px;
    box-shadow: 0 5px 16px rgba(23,54,93,.18);
    flex-shrink: 0;
}

.search-label {
    font-size: 15px;
    font-weight: 700;
    color: #24364b;
    margin: 10px 0 5px;
}

div[data-testid="stTextInput"] input {
    height: 52px;
    border-radius: 12px;
    font-size: 17px;
}

div.stButton > button {
    border-radius: 10px;
    height: 48px;
    font-weight: 700;
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
    overflow-wrap: anywhere;
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
<div class="topline">

    <div>
        <div class="brand-title">
            Asistente Territorial de Pasto
        </div>

        <div class="brand-sub">
            Consulta organizada de información territorial
            por coincidencia de nombre
        </div>
    </div>

    <div class="jc-logo">JC</div>

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
# BUSCADOR SUPERIOR
# ------------------------------------------------------------

st.markdown(
    '<div class="search-label">'
    '¿Qué lugar o nombre deseas consultar?'
    '</div>',
    unsafe_allow_html=True
)

c1, c2 = st.columns([5, 1])

with c1:

    query = st.text_input(
        "Buscar",
        placeholder=(
            "Ejemplo: El Tejar, comuna 5, "
            "Catambuco, una dirección o teléfono…"
        ),
        label_visibility="collapsed",
        key="territorial_query",
    )

with c2:

    search_clicked = st.button(
        "Buscar",
        type="primary",
        use_container_width=True
    )


# ------------------------------------------------------------
# RESULTADOS
# ------------------------------------------------------------

if search_clicked or query:

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
                "No encontré una coincidencia "
                "suficientemente clara en el Excel."
            )

            st.caption(
                "Verifica la escritura o intenta "
                "con una parte más corta del nombre."
            )

    else:

        st.info(
            "Escribe un nombre en el recuadro "
            "superior para iniciar la búsqueda."
        )


# ------------------------------------------------------------
# EJEMPLOS DE BÚSQUEDA
# ------------------------------------------------------------

with st.expander("Ejemplos de búsqueda"):

    st.write("• Urbana")
    st.write("• Rural")
    st.write("• Comuna 5")
    st.write("• El Tejar")
    st.write("• Catambuco")
    st.write("• Nombre de una vereda o sector")
    st.write("• Dirección, edificio, local o teléfono")


# ------------------------------------------------------------
# ESTADO Y COBERTURA DE LA BASE
# ------------------------------------------------------------

with st.expander("Estado y cobertura de la base"):

    st.write(f"**Archivo consultado:** {source_name}")

    st.write(f"**Registros con información:** {len(df)}")

    available_primary = [
        f for f in PRIMARY_FIELDS
        if f in df.columns
    ]

    available_extra = [
        f for f in EXTRA_FIELDS
        if f in df.columns
    ]

    st.write(
        "**Campos territoriales:** "
        + (
            ", ".join(available_primary)
            if available_primary
            else "No identificados"
        )
    )

    st.write(
        "**Campos complementarios disponibles:** "
        + (
            ", ".join(available_extra)
            if available_extra
            else "Ninguno"
        )
    )

    missing_extra = [
        f for f in EXTRA_FIELDS
        if f not in df.columns
    ]

    if missing_extra:

        st.warning(
            "El Excel actual no contiene columnas "
            "reconocibles para: "
            + ", ".join(missing_extra)
            + ". Se mostrarán automáticamente "
            "cuando se agreguen al archivo."
        )

    st.caption(
        "La aplicación muestra únicamente datos "
        "existentes en el Excel. No completa ni "
        "inventa información."
    )
