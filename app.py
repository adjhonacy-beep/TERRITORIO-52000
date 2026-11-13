from pathlib import Path

code = r'''import re
import html
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

# ------------------------------------------------------------
# ASISTENTE TERRITORIAL DE PASTO - JC
# Consulta basada exclusivamente en los datos del archivo Excel.
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
EXTRA_FIELDS = ["EDIFICIO", "DIRECCION", "LOCAL", "TELEFONO"]

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
        "TELEFONO", "TELÉFONO", "CELULAR", "CONTACTO",
        "TELEFONOS", "TELÉFONOS"
    ],
}


# ------------------------------------------------------------
# NORMALIZACIÓN Y LIMPIEZA
# ------------------------------------------------------------

def normalize(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    value = str(value).strip().lower()
    value = "".join(
        char for char in unicodedata.normalize("NFD", value)
        if unicodedata.category(char) != "Mn"
    )
    value = re.sub(r"[^a-z0-9#]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def clean_cell(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    value = str(value).strip()

    if value.lower() in {"nan", "none", "null"}:
        return ""

    # Evita mostrar teléfonos o números como 3001234567.0
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]

    return value


def unique_nonempty(values):
    found = {}

    for value in values:
        cleaned = clean_cell(value)
        key = normalize(cleaned)

        if key and key not in found:
            found[key] = cleaned

    return sorted(found.values(), key=normalize)


# ------------------------------------------------------------
# IDENTIFICACIÓN DE COLUMNAS DEL EXCEL
# ------------------------------------------------------------

def resolve_columns(frame):
    mapping = {}
    normalized_headers = {}

    for column in frame.columns:
        key = normalize(column).replace(" ", "")
        normalized_headers.setdefault(key, column)

    for target, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            key = normalize(alias).replace(" ", "")
            if key in normalized_headers:
                mapping[target] = normalized_headers[key]
                break

    return mapping


# ------------------------------------------------------------
# CARGA DE LA BASE DE DATOS
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

    for field, original_column in mapping.items():
        data[field] = raw[original_column].map(clean_cell)

    available = [
        field for field in PRIMARY_FIELDS + EXTRA_FIELDS
        if field in data.columns
    ]
    data = data[available]

    # Elimina solamente filas completamente vacías.
    nonempty_rows = data.apply(
        lambda row: any(clean_cell(value) for value in row),
        axis=1
    )
    data = data.loc[nonempty_rows]

    return data.reset_index(drop=True), source_name, mapping


# ------------------------------------------------------------
# MOTOR DE COINCIDENCIAS
# ------------------------------------------------------------

def match_score(query, value, field=""):
    nq = normalize(query)
    nv = normalize(value)

    if not nq or not nv:
        return 0

    # Teléfonos: no se aplican coincidencias difusas.
    if field == "TELEFONO":
        query_digits = re.sub(r"\D", "", nq)
        value_digits = re.sub(r"\D", "", nv)

        if len(query_digits) < 7:
            return 0
        if query_digits == value_digits:
            return 100
        if len(query_digits) >= 7 and query_digits in value_digits:
            return 95
        return 0

    # Permite consultar "5" y encontrar "Comuna 5".
    query_without_comuna = re.sub(
        r"^comunas?\s*", "", nq
    ).strip()
    value_without_comuna = re.sub(
        r"^comunas?\s*", "", nv
    ).strip()

    if nq == nv or nq == value_without_comuna:
        return 100

    if query_without_comuna and query_without_comuna == value_without_comuna:
        return 100

    # Coincidencia por frase contenida.
    if len(nq) >= 3 and nq in nv:
        return 96
    if len(query_without_comuna) >= 3 and query_without_comuna in value_without_comuna:
        return 96

    # Coincidencia aproximada.
    return max(
        fuzz.ratio(nq, nv),
        fuzz.token_set_ratio(nq, nv)
    )


def find_matches(data, field, query, threshold=84):
    if field not in data.columns:
        return []

    values = unique_nonempty(data[field].tolist())
    exact = []
    approximate = []

    for value in values:
        score = match_score(query, value, field)

        if score == 100:
            exact.append((value, score))
        elif score >= threshold:
            approximate.append((value, score))

    # Si hay coincidencia exacta en ese campo, no mezcla
    # nombres aproximados con el resultado exacto.
    matches = exact if exact else approximate

    return sorted(
        matches,
        key=lambda item: (-item[1], normalize(item[0]))
    )


def get_related(data, field, matches):
    if field not in data.columns or not matches:
        return data.iloc[0:0].copy()

    matched_values = {
        normalize(value) for value, _score in matches
    }

    return (
        data[data[field].map(normalize).isin(matched_values)]
        .drop_duplicates()
        .reset_index(drop=True)
    )


def search_records(question, data):
    query = normalize(question)

    if not query:
        return data.iloc[0:0].copy(), "", 0

    # --------------------------------------------------------
    # 1. URBANA: mostrar todas las comunas.
    # --------------------------------------------------------
    if query in {"urbana", "urbano", "zona urbana"}:
        if "COMUNA" not in data.columns:
            return data.iloc[0:0].copy(), "", 0

        comunas = unique_nonempty(data["COMUNA"].tolist())
        if not comunas:
            return data.iloc[0:0].copy(), "", 0

        result = pd.DataFrame({"COMUNA": comunas})
        return result, "Todas las comunas de Pasto", 100

    # --------------------------------------------------------
    # 2. RURAL: mostrar registros de todos los corregimientos,
    # incluyendo veredas y sectores existentes en el Excel.
    # --------------------------------------------------------
    if query in {"rural", "zona rural"}:
        if "CORREGIMIENTO" not in data.columns:
            return data.iloc[0:0].copy(), "", 0

        rural_rows = data[
            data["CORREGIMIENTO"].map(normalize) != ""
        ].copy()

        if rural_rows.empty:
            return data.iloc[0:0].copy(), "", 0

        count = len(unique_nonempty(rural_rows["CORREGIMIENTO"]))
        return (
            rural_rows.drop_duplicates().reset_index(drop=True),
            f"Todos los corregimientos de Pasto ({count})",
            100
        )

    # --------------------------------------------------------
    # 3. BÚSQUEDA POR COMUNA: barrios, edificios y demás datos.
    # --------------------------------------------------------
    comuna_matches = find_matches(data, "COMUNA", query)

    if comuna_matches:
        rows = get_related(data, "COMUNA", comuna_matches)
        names = ", ".join(value for value, _score in comuna_matches)
        score = max(value for _name, value in comuna_matches)
        return rows, names, score

    # --------------------------------------------------------
    # 4. BÚSQUEDA POR CORREGIMIENTO: veredas y sectores.
    # --------------------------------------------------------
    corregimiento_matches = find_matches(
        data, "CORREGIMIENTO", query
    )

    if corregimiento_matches:
        rows = get_related(
            data, "CORREGIMIENTO", corregimiento_matches
        )
        names = ", ".join(
            value for value, _score in corregimiento_matches
        )
        score = max(value for _name, value in corregimiento_matches)
        return rows, names, score

    # --------------------------------------------------------
    # 5. BÚSQUEDA POR BARRIO: edificios y demás datos.
    # --------------------------------------------------------
    barrio_matches = find_matches(data, "BARRIO", query)

    if barrio_matches:
        rows = get_related(data, "BARRIO", barrio_matches)
        names = ", ".join(value for value, _score in barrio_matches)
        score = max(value for _name, value in barrio_matches)
        return rows, names, score

    # --------------------------------------------------------
    # 6. BÚSQUEDA GENERAL: todos los campos disponibles.
    # --------------------------------------------------------
    all_matches = []

    for field in data.columns:
        for value, score in find_matches(data, field, query):
            all_matches.append((field, value, score))

    # Si no hubo coincidencia por frase, intenta por palabras.
    if not all_matches:
        tokens = [token for token in query.split() if len(token) >= 3]

        for field in data.columns:
            if field == "TELEFONO":
                continue

            for value in unique_nonempty(data[field].tolist()):
                normalized_value = normalize(value)

                if normalized_value and any(
                    token in normalized_value.split()
                    for token in tokens
                ):
                    all_matches.append((field, value, 78))

    if not all_matches:
        return data.iloc[0:0].copy(), "", 0

    # Recupera las filas completas para conservar sus relaciones.
    result_frames = []

    for field, value, _score in all_matches:
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
        value for _field, value, _score in all_matches
    ])

    matched_name = ", ".join(names)
    best_score = max(score for _field, _value, score in all_matches)

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
        "Se muestran todos los valores únicos disponibles."
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

    main_fields = [
        (field, values)
        for field, values in grouped.items()
        if field in PRIMARY_FIELDS
    ]
    extra_fields = [
        (field, values)
        for field, values in grouped.items()
        if field in EXTRA_FIELDS
    ]

    if main_fields:
        columns = st.columns(3)

        for index, (field, values) in enumerate(main_fields):
            value_html = "<br>".join(
                html.escape(value) for value in values
            )

            with columns[index % 3]:
                st.markdown(
                    f'<div class="field">'
                    f'<div class="field-label">{field.title()}</div>'
                    f'<div class="field-value">{value_html}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    if extra_fields:
        st.markdown(
            '<div class="extra-title">'
            'Información complementaria'
            '</div>',
            unsafe_allow_html=True,
        )

        columns = st.columns(2)

        for index, (field, values) in enumerate(extra_fields):
            value_html = "<br>".join(
                html.escape(value) for value in values
            )

            with columns[index % 2]:
                st.markdown(
                    f'<div class="field">'
                    f'<div class="field-label">{field.title()}</div>'
                    f'<div class="field-value">{value_html}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    st.markdown('</div>', unsafe_allow_html=True)


# ------------------------------------------------------------
# DISEÑO VISUAL ORIGINAL JC
# ------------------------------------------------------------

st.markdown("""
<style>
#MainMenu, footer {visibility:hidden;}

.block-container {
    padding-top:1.5rem;
    max-width:1180px;
}

.topline {
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:18px;
    border-bottom:1px solid #e7eaf0;
    padding-bottom:16px;
    margin-bottom:18px;
}

.brand-title {
    font-size:27px;
    font-weight:800;
    color:#17365d;
    line-height:1.15;
}

.brand-sub {
    font-size:13px;
    color:#687386;
    margin-top:5px;
}

.jc-logo {
    width:62px;
    height:62px;
    border-radius:16px;
    display:flex;
    align-items:center;
    justify-content:center;
    background:linear-gradient(135deg,#17365d,#247ba0);
    color:white;
    font-size:24px;
    font-weight:900;
    letter-spacing:1px;
    box-shadow:0 5px 16px rgba(23,54,93,.18);
    flex-shrink:0;
}

.search-label {
    font-size:15px;
    font-weight:700;
    color:#24364b;
    margin:10px 0 5px;
}

div[data-testid="stTextInput"] input {
    height:52px;
    border-radius:12px;
    font-size:17px;
}

div.stButton > button {
    border-radius:10px;
    height:48px;
    font-weight:700;
}

.result-card {
    background:#fff;
    border:1px solid #e5eaf0;
    border-radius:15px;
    padding:17px 18px;
    margin:8px 0 15px;
    box-shadow:0 3px 12px rgba(25,45,75,.045);
}

.field {
    padding:8px 9px;
    margin:4px 0;
    background:#f7f9fc;
    border-radius:9px;
    min-height:64px;
}

.field-label {
    font-size:11px;
    color:#6c7788;
    text-transform:uppercase;
    letter-spacing:.6px;
    font-weight:700;
}

.field-value {
    font-size:15px;
    color:#172b4d;
    font-weight:650;
    overflow-wrap:anywhere;
    margin-top:4px;
}

.extra-title {
    font-weight:750;
    color:#244b70;
    border-top:1px solid #e9edf2;
    padding-top:12px;
    margin-top:12px;
}

.match-heading {
    font-size:18px;
    color:#1d3557;
    margin:16px 0 6px;
    overflow-wrap:anywhere;
}

.match-score {
    float:right;
    color:#16734a;
    font-size:12px;
    background:#e8f6ef;
    padding:4px 9px;
    border-radius:20px;
}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------
# ENCABEZADO
# ------------------------------------------------------------

st.markdown("""
<div class="topline">
  <div>
    <div class="brand-title">Asistente Territorial de Pasto</div>
    <div class="brand-sub">
      Consulta organizada de información territorial por coincidencia de nombre
    </div>
  </div>
  <div class="jc-logo">JC</div>
</div>
""", unsafe_allow_html=True)


# ------------------------------------------------------------
# LECTURA DEL EXCEL
# ------------------------------------------------------------

try:
    df, source_name, column_map = load_excel()
except Exception as exc:
    st.error(f"No fue posible leer el archivo Excel: {exc}")
    st.stop()

if df is None:
    st.error(
        "No se encontró el archivo Pasto.xlsx. "
        "Verifica que esté en la misma carpeta de app.py "
        "y en el repositorio de GitHub."
    )
    st.stop()


# ------------------------------------------------------------
# BUSCADOR
# ------------------------------------------------------------

st.markdown(
    '<div class="search-label">'
    '¿Qué lugar o nombre deseas consultar?'
    '</div>',
    unsafe_allow_html=True
)

column1, column2 = st.columns([5, 1])

with column1:
    query = st.text_input(
        "Buscar",
        placeholder=(
            "Ejemplo: Urbana, Rural, El Tejar, "
            "Comuna 5, Catambuco, dirección o teléfono…"
        ),
        label_visibility="collapsed",
        key="territorial_query",
    )

with column2:
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
        results, matched_name, score = search_records(query, df)

        if not results.empty:
            render_result(results, matched_name, score)
        else:
            st.warning(
                "No encontré una coincidencia suficientemente clara en el Excel."
            )
            st.caption(
                "Verifica la escritura o intenta con una parte más corta del nombre."
            )
    else:
        st.info(
            "Escribe un nombre en el recuadro superior para iniciar la búsqueda."
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
        field for field in PRIMARY_FIELDS if field in df.columns
    ]
    available_extra = [
        field for field in EXTRA_FIELDS if field in df.columns
    ]
from pathlib import Path

# 1. Define correctamente tu ruta y el código completo dentro de la variable 'code'
path = Path("mi_aplicacion.py")  # O la ruta que estés utilizando

code = """
    st.write(
        "**Campos territoriales:** "
        + (
            ", ".join(available_primary)
            if available_primary else "No identificados"
        )
    )
    st.write(
        "**Campos complementarios disponibles:** "
        + (
            ", ".join(available_extra)
            if available_extra else "Ninguno"
        )
    )

    missing_extra = [
        field for field in EXTRA_FIELDS if field not in df.columns
    ]

    if missing_extra:
        st.warning(
            "El Excel actual no contiene columnas reconocibles para: "
            + ", ".join(missing_extra)
            + ". Se mostrarán automáticamente cuando se agreguen al archivo."
        )

    st.caption(
        "La aplicación muestra únicamente datos existentes en el Excel. "
        "No completa ni inventa información."
    )
"""

# 2. Comprobación sintáctica antes de entregar o ejecutar el archivo
try:
    compile(code, str(path), "exec")
    print(f"Archivo comprobado exitosamente: {path}")
    print(f"Líneas de código: {len(code.splitlines())}")
except SyntaxError as e:
    print(f"Error de sintaxis encontrado: {e}")
