from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

# Base de datos simulada o adaptada a tu estructura en Python
BASE_TERRITORIAL = {
    "EL ARROYO": {
        "zona": ["Rural"],
        "nivel2": ["ANGANOY"],
        "nivel3": ["ANGANOY", "SAN JUAN DE ANGANOY", "CHACHAGÜÍ ALTO"],
        "nivel4": ["EL ARROYO", "EL ARROYO BAJO", "LA QUEBRADA"]
    },
    "TAMASAGRA": {
        "zona": ["Urbano"],
        "nivel2": ["COMUNA 6"],
        "nivel3": ["TAMASAGRA I", "TAMASAGRA II", "NUEVA ARANDA"],
        "nivel4": ["SECTOR PARQUE", "MANZANA 14", "SECTOR PRINCIPAL"]
    }
}

@app.route("/")
def index():
    # Renderiza la página principal por defecto con "EL ARROYO"
    return render_template("index.html")

@app.route("/buscar", methods=["GET"])
def buscar():
    query = request.args.get("q", "").strip().upper()
    registro = BASE_TERRITORIAL.get(query)
    
    if not registro:
        return jsonify({"encontrado": False, "mensaje": f"No se encontraron coincidencias para '{query}'"})
    
    # Determina dinámicamente los encabezados según la Zona
    es_urbano = registro["zona"][0].lower() == "urbano"
    encabezados = ["ZONA", "COMUNA", "BARRIO", "SECTOR"] if es_urbano else ["ZONA", "CORREGIMIENTO", "VEREDA", "SECTOR"]
    
    return jsonify({
        "encontrado": True,
        "encabezados": encabezados,
        "datos": [
            registro["zona"],
            registro["nivel2"],
            registro["nivel3"],
            registro["nivel4"]
        ]
    })

if __name__ == "__main__":
    app.run(debug=True)
