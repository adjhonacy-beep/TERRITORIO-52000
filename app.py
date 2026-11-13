import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Asistente Territorial de Pasto", page_icon="🗺️", layout="centered")
st.title("🗺️ Asistente Territorial de Pasto")
st.caption("Consulta barrios, comunas, corregimientos y veredas con base en tu Excel.")

DEFAULT_FILE = Path(__file__).parent / "Pasto.xlsx"

@st.cache_data
def load_data(file_bytes=None, file_name=None):
    if file_bytes is not None:
        df = pd.read_excel(file_bytes)
    else:
        df = pd.read_excel(DEFAULT_FILE)
    # Keep only the useful columns and normalize blanks.
    expected = ["ZONA", "COMUNA", "BARRIO", "CORREGIMIENTO", "VEREDA", "SECTOR"]
    for c in expected:
        if c not in df.columns:
            df[c] = None
    df = df[expected].copy()
    for c in expected:
        df[c] = df[c].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() else None)
    df = df.dropna(how="all", subset=expected)
    return df

def norm(value):
    value = str(value or "").lower().strip()
    value = "".join(ch for ch in unicodedata.normalize("NFD", value) if unicodedata.category(ch) != "Mn")
    value = re.sub(r"[^a-z0-9 ]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()

def unique_values(df, col):
    return sorted({v for v in df[col].dropna().astype(str).str.strip() if v}, key=str.casefold)

def find_value(query, values, threshold=82):
    q = norm(query)
    if not q:
        return None
    exact = {norm(v): v for v in values}
    if q in exact:
        return exact[q]
    # Look for a full name contained in the user's sentence.
    contained = [(v, norm(v)) for v in values if norm(v) and norm(v) in q]
    if contained:
        return max(contained, key=lambda item: len(item[1]))[0]
    # Fuzzy match only when it is reasonably close.
    match = process.extractOne(q, values, scorer=lambda a, b: fuzz.ratio(norm(a), norm(b)))
    if match and match[1] >= threshold:
        return match[0]
    return None

def answer_question(question, df):
    q = norm(question)
    barrios = unique_values(df, "BARRIO")
    comunas = unique_values(df, "COMUNA")
    corregimientos = unique_values(df, "CORREGIMIENTO")
    veredas = unique_values(df, "VEREDA")
    sectores = unique_values(df, "SECTOR")

    # Comuna -> list of barrios
    comuna = find_value(question, comunas)
    asks_list = any(word in q for word in ["barrios", "lista", "listado", "cuales son", "que barrios", "barrios de"])
    if comuna and (asks_list or "pertenecen" in q or "incluye" in q):
        items = unique_values(df[df["COMUNA"].map(lambda x: norm(x) == norm(comuna))], "BARRIO")
        if items:
            return f"**{comuna}** tiene **{len(items)} barrios registrados** en el Excel:\n\n" + "\n".join(f"- {x}" for x in items)
        return f"No encontré barrios asociados a {comuna} en el archivo."

    # Barrio -> comuna
    barrio = find_value(question, barrios)
    if barrio:
        rows = df[df["BARRIO"].map(lambda x: norm(x) == norm(barrio))]
        related = rows[["COMUNA", "ZONA"]].drop_duplicates()
        comunas_barrio = [x for x in related["COMUNA"].dropna().astype(str).unique()]
        zonas = [x for x in related["ZONA"].dropna().astype(str).unique()]
        result = f"El barrio **{barrio}** aparece en el Excel."
        if comunas_barrio:
            result += "\n\n**Comuna:** " + ", ".join(comunas_barrio)
        if zonas:
            result += "\n\n**Zona:** " + ", ".join(zonas)
        if len(comunas_barrio) > 1:
            result += "\n\n⚠️ El barrio aparece asociado a más de una comuna. Verifica los registros."
        return result

    # Corregimiento -> list of veredas/sectors
    corregimiento = find_value(question, corregimientos)
    if corregimiento:
        rows = df[df["CORREGIMIENTO"].map(lambda x: norm(x) == norm(corregimiento))]
        items = unique_values(rows, "VEREDA")
        sect = unique_values(rows, "SECTOR")
        if any(word in q for word in ["veredas", "sectores", "lista", "listado", "cuales son", "que tiene", "pertenecen"]):
            output = f"**{corregimiento}** tiene **{len(items)} veredas registradas**"
            if sect:
                output += f" y **{len(sect)} sectores registrados**"
            output += " en el Excel.\n\n"
            if items:
                output += "**Veredas:**\n" + "\n".join(f"- {x}" for x in items)
            if sect:
                output += "\n\n**Sectores:**\n" + "\n".join(f"- {x}" for x in sect)
            if not items and not sect:
                output += "No hay veredas ni sectores asociados en el archivo."
            return output
        return f"**{corregimiento}** aparece en la columna CORREGIMIENTO. Puedes preguntar: “¿Qué veredas tiene {corregimiento}?”"

    # Direct lookup of vereda or sector
    vereda = find_value(question, veredas)
    if vereda:
        rows = df[df["VEREDA"].map(lambda x: norm(x) == norm(vereda))]
        corr = unique_values(rows, "CORREGIMIENTO")
        out = f"La vereda **{vereda}** aparece en el Excel."
        if corr:
            out += "\n\n**Corregimiento:** " + ", ".join(corr)
        return out

    sector = find_value(question, sectores)
    if sector:
        rows = df[df["SECTOR"].map(lambda x: norm(x) == norm(sector))]
        corr = unique_values(rows, "CORREGIMIENTO")
        out = f"El sector **{sector}** aparece en el Excel."
        if corr:
            out += "\n\n**Corregimiento:** " + ", ".join(corr)
        return out

    if any(word in q for word in ["direccion", "dirección", "carrera", "calle", "diagonal", "transversal", "avenida", "cra", "cll", "cr "]) or re.search(r"\b\d+\s*(con|#|-)\s*\d+", q):
        return ("Tu Excel actual no contiene una columna de DIRECCIÓN ni coordenadas o rangos de vías. "
                "Por eso no puedo determinar de forma confiable el barrio y la comuna a partir de una dirección. "
                "Para habilitarlo, agrega una tabla de direcciones con columnas como DIRECCIÓN, BARRIO y COMUNA, "
                "o incorpora una capa cartográfica validada con los límites territoriales.")

    # If no match, suggest closest names
    all_names = barrios + comunas + corregimientos + veredas + sectores
    suggestions = process.extract(question, all_names, scorer=lambda a, b: fuzz.ratio(norm(a), norm(b)), limit=3)
    suggestions = [name for name, score, _ in suggestions if score >= 45]
    msg = "No encontré una coincidencia clara en el Excel."
    if suggestions:
        msg += "\n\n¿Quizás te refieres a alguno de estos nombres?\n" + "\n".join(f"- {s}" for s in suggestions)
    msg += "\n\nPuedes preguntar, por ejemplo:\n- ¿A qué comuna pertenece el barrio El Tejar?\n- ¿Qué barrios tiene la comuna 5?\n- ¿Qué veredas tiene el corregimiento Catambuco?"
    return msg

st.sidebar.header("Datos")
st.sidebar.write("Archivo base: `Pasto.xlsx`")
st.sidebar.caption("Puedes cargar una versión actualizada del Excel aquí. Los cambios se aplican a esta sesión.")

uploaded = st.sidebar.file_uploader("Actualizar Excel", type=["xlsx"])
try:
    if uploaded is not None:
        df = load_data(uploaded, uploaded.name)
    else:
        if not DEFAULT_FILE.exists():
            st.error("No se encontró Pasto.xlsx. Coloca el archivo en la misma carpeta que app.py o cárgalo en el panel lateral.")
            st.stop()
        df = load_data()
except Exception as exc:
    st.error(f"No pude leer el Excel: {exc}")
    st.stop()

st.sidebar.metric("Filas leídas", len(df))
st.sidebar.metric("Barrios distintos", len(unique_values(df, "BARRIO")))
st.sidebar.metric("Comunas", len(unique_values(df, "COMUNA")))
st.sidebar.metric("Corregimientos", len(unique_values(df, "CORREGIMIENTO")))

if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "¡Hola! Pregúntame por un barrio, comuna, corregimiento, vereda o sector del archivo Pasto.xlsx."}
    ]

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

question = st.chat_input("Escribe tu consulta territorial...")
if question:
    st.session_state.messages.append({"role": "user", "content": question})
    with st.chat_message("user"):
        st.markdown(question)
    result = answer_question(question, df)
    st.session_state.messages.append({"role": "assistant", "content": result})
    with st.chat_message("assistant"):
        st.markdown(result)

with st.expander("Ver una muestra de la base"):
    st.dataframe(df.head(15), use_container_width=True)
