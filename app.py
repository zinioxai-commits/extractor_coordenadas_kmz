import streamlit as st
import zipfile
import re
import io
from pyproj import Transformer
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuración de página ───────────────────────────────────────
st.set_page_config(
    page_title="Extractor KMZ → UTM",
    page_icon="📍",
    layout="centered",
)

# ── CSS personalizado ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

.stApp {
    background-color: #0f1117;
    color: #e8e8e8;
}

/* Header principal */
.header-block {
    border-left: 4px solid #00d4aa;
    padding: 1rem 1.5rem;
    margin-bottom: 2rem;
    background: linear-gradient(90deg, rgba(0,212,170,0.07) 0%, transparent 100%);
}
.header-block h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: #00d4aa;
    margin: 0 0 0.2rem 0;
    letter-spacing: -0.5px;
}
.header-block p {
    color: #888;
    font-size: 0.9rem;
    margin: 0;
    font-weight: 300;
}

/* Zona selector */
.zona-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #00d4aa;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 0.3rem;
}

/* Métrica cards */
.metric-row {
    display: flex;
    gap: 1rem;
    margin: 1.5rem 0;
}
.metric-card {
    flex: 1;
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    border-top: 2px solid #00d4aa;
}
.metric-card .val {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.8rem;
    font-weight: 600;
    color: #00d4aa;
    line-height: 1;
}
.metric-card .lbl {
    font-size: 0.78rem;
    color: #666;
    margin-top: 0.3rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Tabla preview */
.preview-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #555;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin: 1.5rem 0 0.5rem 0;
}

/* Upload zone */
[data-testid="stFileUploader"] {
    border: 2px dashed #2a2d3a !important;
    border-radius: 10px !important;
    background: #13161f !important;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #00d4aa !important;
}

/* Selectbox */
[data-testid="stSelectbox"] > div > div {
    background: #1a1d27 !important;
    border-color: #2a2d3a !important;
    color: #e8e8e8 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.9rem !important;
}

/* Botón download */
[data-testid="stDownloadButton"] button {
    background: #00d4aa !important;
    color: #0f1117 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 0.6rem 2rem !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.5px !important;
    width: 100%;
    transition: opacity 0.2s !important;
}
[data-testid="stDownloadButton"] button:hover {
    opacity: 0.85 !important;
}

/* Success / warning alerts */
[data-testid="stAlert"] {
    border-radius: 6px !important;
    font-size: 0.88rem !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    overflow: hidden;
}

/* Divider */
hr {
    border-color: #1e2130 !important;
    margin: 2rem 0 !important;
}

/* Footer */
.footer {
    text-align: center;
    color: #333;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid #1e2130;
}
</style>
""", unsafe_allow_html=True)


# ── Funciones ─────────────────────────────────────────────────────
def leer_kml(archivo_bytes: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(archivo_bytes)) as kmz:
        kml_file = next(f for f in kmz.namelist() if f.endswith(".kml"))
        return kmz.read(kml_file).decode("utf-8")


def extraer_puntos(kml_texto: str) -> list[dict]:
    placemarks = re.findall(r"<Placemark>(.*?)</Placemark>", kml_texto, re.DOTALL)
    puntos = []
    for i, pm in enumerate(placemarks):
        nombre = re.search(r"<n>(.*?)</n>", pm)
        desc   = re.search(r"<description>(.*?)</description>", pm)
        coords = re.search(r"<coordinates>(.*?)</coordinates>", pm, re.DOTALL)
        if coords:
            partes = coords.group(1).strip().split(",")
            puntos.append({
                "N":           i + 1,
                "Nombre":      nombre.group(1).strip() if nombre else str(i + 1),
                "Descripcion": desc.group(1).strip()   if desc   else "",
                "Longitud":    round(float(partes[0]), 7),
                "Latitud":     round(float(partes[1]), 7),
            })
    return puntos


def convertir_utm(puntos: list[dict], epsg: str) -> list[dict]:
    t = Transformer.from_crs("EPSG:4326", epsg, always_xy=True)
    for p in puntos:
        e, n = t.transform(p["Longitud"], p["Latitud"])
        p["Este_UTM"]  = round(e, 2)
        p["Norte_UTM"] = round(n, 2)
    return puntos


def generar_xlsx(puntos: list[dict], zona: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Coordenadas"

    h_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    h_fill  = PatternFill("solid", start_color="1a6e5e")
    h_align = Alignment(horizontal="center", vertical="center")
    c_align = Alignment(horizontal="center")
    borde   = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    fill_par   = PatternFill("solid", start_color="E8F8F5")
    fill_impar = PatternFill("solid", start_color="FFFFFF")

    cols = ["N", "Nombre", "Descripción", "Longitud", "Latitud",
            f"Este UTM ({zona})", f"Norte UTM ({zona})"]

    for col, texto in enumerate(cols, 1):
        c = ws.cell(row=1, column=col, value=texto)
        c.font = h_font; c.fill = h_fill
        c.alignment = h_align; c.border = borde
    ws.row_dimensions[1].height = 22

    for fila, p in enumerate(puntos, 2):
        vals = [p["N"], p["Nombre"], p["Descripcion"],
                p["Longitud"], p["Latitud"], p["Este_UTM"], p["Norte_UTM"]]
        fill = fill_par if fila % 2 == 0 else fill_impar
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=fila, column=col, value=val)
            c.border = borde; c.alignment = c_align; c.fill = fill
            if col == 4: c.number_format = "0.0000000"
            if col == 5: c.number_format = "0.0000000"
            if col in (6, 7): c.number_format = "#,##0.00"

    anchos = [6, 12, 22, 16, 16, 20, 20]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Zonas UTM disponibles ─────────────────────────────────────────
ZONAS = {
    "17L  — Zona 17 Sur  (EPSG:32717)": ("EPSG:32717", "17L"),
    "18L  — Zona 18 Sur  (EPSG:32718)": ("EPSG:32718", "18L"),
    "19L  — Zona 19 Sur  (EPSG:32719)": ("EPSG:32719", "19L"),
    "17N  — Zona 17 Norte (EPSG:32617)": ("EPSG:32617", "17N"),
    "18N  — Zona 18 Norte (EPSG:32618)": ("EPSG:32618", "18N"),
    "19N  — Zona 19 Norte (EPSG:32619)": ("EPSG:32619", "19N"),
}


# ── UI ────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-block">
    <h1>📍 KMZ → UTM</h1>
    <p>Extrae coordenadas de archivos KMZ y convierte a sistema UTM · Exporta a Excel</p>
</div>
""", unsafe_allow_html=True)

# Zona UTM
st.markdown('<div class="zona-label">Zona UTM de destino</div>', unsafe_allow_html=True)
zona_key = st.selectbox("", list(ZONAS.keys()),
                        index=2, label_visibility="collapsed")
epsg_sel, nombre_zona = ZONAS[zona_key]

st.markdown("<br>", unsafe_allow_html=True)

# Upload
archivo = st.file_uploader("Sube tu archivo KMZ", type=["kmz"],
                            help="Arrastra el archivo .kmz aquí")

if archivo:
    try:
        kml_texto = leer_kml(archivo.read())
        puntos    = extraer_puntos(kml_texto)

        if not puntos:
            st.warning("No se encontraron puntos en el archivo KMZ.")
        else:
            puntos = convertir_utm(puntos, epsg_sel)

            # Métricas
            lons = [p["Longitud"] for p in puntos]
            lats = [p["Latitud"]  for p in puntos]
            estes = [p["Este_UTM"] for p in puntos]

            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-card">
                    <div class="val">{len(puntos)}</div>
                    <div class="lbl">Puntos extraídos</div>
                </div>
                <div class="metric-card">
                    <div class="val">{nombre_zona}</div>
                    <div class="lbl">Zona UTM</div>
                </div>
                <div class="metric-card">
                    <div class="val">{round(min(lats),4)}</div>
                    <div class="lbl">Lat mínima</div>
                </div>
                <div class="metric-card">
                    <div class="val">{round(max(lats),4)}</div>
                    <div class="lbl">Lat máxima</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Preview tabla
            st.markdown('<div class="preview-title">Vista previa — primeros 10 puntos</div>',
                        unsafe_allow_html=True)

            import pandas as pd
            df = pd.DataFrame(puntos)
            st.dataframe(df.head(10), use_container_width=True, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Descarga Excel
            xlsx_bytes = generar_xlsx(puntos, nombre_zona)
            nombre_salida = archivo.name.replace(".kmz", f"_UTM_{nombre_zona}.xlsx")

            st.download_button(
                label=f"⬇  Descargar Excel — {len(puntos)} puntos ({nombre_zona})",
                data=xlsx_bytes,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

else:
    st.markdown("""
    <div style="text-align:center; padding: 3rem 1rem; color: #333;
                font-family: 'IBM Plex Mono', monospace; font-size: 0.85rem;">
        ↑ Sube un archivo .kmz para comenzar
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    Soporta WGS84 → UTM · Zonas 17-19 N/S · Exporta .xlsx con formato
</div>
""", unsafe_allow_html=True)
