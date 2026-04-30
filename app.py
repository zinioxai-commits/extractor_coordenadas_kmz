import streamlit as st
import zipfile
import re
import io
import pandas as pd
from pyproj import Transformer
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Redsetel · Extractor KMZ",
    page_icon="📡",
    layout="centered"
)

# ── Paleta Redsetel: rojo #CC1E27, azul oscuro #0D2650, blanco #FFFFFF ──
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif;
}

.stApp {
    background-color: #f4f6f9;
    color: #1a1a2e;
}

/* ── Topbar ── */
.topbar {
    background: linear-gradient(135deg, #0D2650 0%, #1a3a6e 100%);
    padding: 1.2rem 2rem;
    border-radius: 12px;
    display: flex;
    align-items: center;
    gap: 1.5rem;
    margin-bottom: 1.8rem;
    box-shadow: 0 4px 20px rgba(13,38,80,0.25);
    border-bottom: 3px solid #CC1E27;
}
.topbar .brand {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    letter-spacing: 1px;
    line-height: 1;
}
.topbar .brand span.red  { color: #CC1E27; }
.topbar .brand span.white { color: #FFFFFF; }
.topbar .subtitle {
    color: rgba(255,255,255,0.55);
    font-size: 0.82rem;
    font-weight: 300;
    margin-top: 0.15rem;
    letter-spacing: 0.3px;
}
.topbar .divider {
    width: 2px; height: 42px;
    background: rgba(255,255,255,0.15);
    border-radius: 2px;
}
.topbar .tool-title {
    color: #FFFFFF;
    font-size: 1rem;
    font-weight: 600;
    line-height: 1.2;
}
.topbar .tool-desc {
    color: rgba(255,255,255,0.45);
    font-size: 0.78rem;
    font-weight: 300;
    margin-top: 0.1rem;
}

/* ── Tarjeta contenedora ── */
.card {
    background: #FFFFFF;
    border-radius: 10px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    border-top: 3px solid #CC1E27;
}
.card-blue {
    border-top-color: #0D2650;
}
.section-label {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.72rem;
    font-weight: 700;
    color: #0D2650;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin-bottom: 0.6rem;
}

/* ── Métricas ── */
.metric-row { display: flex; gap: 0.8rem; margin: 0.5rem 0; }
.metric-card {
    flex: 1;
    background: #f8f9fc;
    border: 1px solid #e4e8f0;
    border-radius: 8px;
    padding: 0.9rem 1rem;
    border-left: 3px solid #CC1E27;
    text-align: center;
}
.metric-card.blue { border-left-color: #0D2650; }
.metric-card .val {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    color: #CC1E27;
    line-height: 1;
}
.metric-card.blue .val { color: #0D2650; }
.metric-card .lbl {
    font-size: 0.7rem;
    color: #888;
    margin-top: 0.25rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    font-weight: 600;
}

/* ── Upload area ── */
[data-testid="stFileUploader"] {
    border: 2px dashed #c8d0e0 !important;
    border-radius: 10px !important;
    background: #f8f9fc !important;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #CC1E27 !important;
}

/* ── Selectbox ── */
[data-testid="stSelectbox"] > div > div {
    background: #f8f9fc !important;
    border-color: #c8d0e0 !important;
    color: #0D2650 !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
}

/* ── Botón descarga ── */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #CC1E27 0%, #a8151d 100%) !important;
    color: #FFFFFF !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    letter-spacing: 1px !important;
    border: none !important;
    border-radius: 8px !important;
    width: 100% !important;
    padding: 0.7rem 2rem !important;
    box-shadow: 0 4px 14px rgba(204,30,39,0.3) !important;
    transition: opacity 0.2s !important;
}
[data-testid="stDownloadButton"] button:hover {
    opacity: 0.88 !important;
}

/* ── Preview title ── */
.preview-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.72rem;
    font-weight: 700;
    color: #0D2650;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin: 1.2rem 0 0.4rem 0;
}

/* ── Footer ── */
.footer {
    text-align: center;
    padding: 1.2rem;
    margin-top: 2rem;
    background: #0D2650;
    border-radius: 8px;
    color: rgba(255,255,255,0.4);
    font-size: 0.75rem;
    letter-spacing: 0.3px;
}
.footer strong { color: rgba(255,255,255,0.7); }

/* ── Badge vacío ── */
.empty-state {
    text-align: center;
    padding: 3rem 1rem;
    color: #bbc4d4;
    font-size: 0.9rem;
}
.empty-state .icon { font-size: 2.5rem; margin-bottom: 0.5rem; }
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
        nombre = re.search(r"<name>(.*?)</name>", pm)
        desc   = re.search(r"<description>(.*?)</description>", pm, re.DOTALL)
        custom = re.search(r'<mwm:customName>.*?<mwm:lang[^>]*>(.*?)</mwm:lang>', pm, re.DOTALL)
        coords = re.search(r"<coordinates>(.*?)</coordinates>", pm, re.DOTALL)
        if coords:
            partes      = coords.group(1).strip().split(",")
            desc_raw    = desc.group(1).strip() if desc else ""
            desc_limpia = desc_raw.split("\n")[0].strip()
            puntos.append({
                "N":           i + 1,
                "Nombre":      nombre.group(1).strip() if nombre else str(i + 1),
                "Item":        custom.group(1).strip()  if custom else "",
                "Descripcion": desc_limpia,
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
    h_fill  = PatternFill("solid", start_color="0D2650")   # azul Redsetel
    h_align = Alignment(horizontal="center", vertical="center")
    c_align = Alignment(horizontal="center")
    borde   = Border(
        left=Side(style="thin"),  right=Side(style="thin"),
        top=Side(style="thin"),   bottom=Side(style="thin"),
    )
    fill_par   = PatternFill("solid", start_color="EEF2F8")
    fill_impar = PatternFill("solid", start_color="FFFFFF")

    cols = ["N", "Nombre", "Item", "Descripción",
            "Longitud", "Latitud",
            f"Este UTM ({zona})", f"Norte UTM ({zona})"]

    for col, texto in enumerate(cols, 1):
        c = ws.cell(row=1, column=col, value=texto)
        c.font = h_font; c.fill = h_fill
        c.alignment = h_align; c.border = borde
    ws.row_dimensions[1].height = 22

    for fila, p in enumerate(puntos, 2):
        vals = [p["N"], p["Nombre"], p["Item"], p["Descripcion"],
                p["Longitud"], p["Latitud"], p["Este_UTM"], p["Norte_UTM"]]
        fill = fill_par if fila % 2 == 0 else fill_impar
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=fila, column=col, value=val)
            c.border = borde; c.alignment = c_align; c.fill = fill
            if col == 5: c.number_format = "0.0000000"
            if col == 6: c.number_format = "0.0000000"
            if col in (7, 8): c.number_format = "#,##0.00"

    anchos = [6, 10, 16, 20, 16, 16, 20, 20]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Zonas UTM ─────────────────────────────────────────────────────
ZONAS = {
    "17L  — Zona 17 Sur  (EPSG:32717)": ("EPSG:32717", "17L"),
    "18L  — Zona 18 Sur  (EPSG:32718)": ("EPSG:32718", "18L"),
    "19L  — Zona 19 Sur  (EPSG:32719)": ("EPSG:32719", "19L"),
    "17N  — Zona 17 Norte (EPSG:32617)": ("EPSG:32617", "17N"),
    "18N  — Zona 18 Norte (EPSG:32618)": ("EPSG:32618", "18N"),
    "19N  — Zona 19 Norte (EPSG:32619)": ("EPSG:32619", "19N"),
}


# ── HEADER ────────────────────────────────────────────────────────
st.markdown("""
<div class="topbar">
    <div>
        <div class="brand">
            <span class="red">red</span><span class="white">setel</span>
        </div>
        <div class="subtitle">RED DE SERVICIOS Y TELECOMUNICACIONES PERÚ</div>
    </div>
    <div class="divider"></div>
    <div>
        <div class="tool-title">📡 Extractor KMZ → UTM</div>
        <div class="tool-desc">Coordenadas geográficas · Conversión · Exportación Excel</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── CONFIGURACIÓN ─────────────────────────────────────────────────
st.markdown('<div class="card card-blue">', unsafe_allow_html=True)
st.markdown('<div class="section-label">⚙ Configuración — Zona UTM</div>', unsafe_allow_html=True)
zona_key = st.selectbox("", list(ZONAS.keys()), index=2, label_visibility="collapsed")
epsg_sel, nombre_zona = ZONAS[zona_key]
st.markdown('</div>', unsafe_allow_html=True)

# ── UPLOAD ────────────────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">📂 Archivo de entrada</div>', unsafe_allow_html=True)
archivo = st.file_uploader("Arrastra tu archivo .kmz aquí", type=["kmz"],
                            label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

# ── RESULTADO ─────────────────────────────────────────────────────
if archivo:
    try:
        kml_texto = leer_kml(archivo.read())
        puntos    = extraer_puntos(kml_texto)

        if not puntos:
            st.warning("No se encontraron puntos en el archivo KMZ.")
        else:
            puntos   = convertir_utm(puntos, epsg_sel)
            con_item = sum(1 for p in puntos if p["Item"])
            sin_item = len(puntos) - con_item

            # Métricas
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-card">
                    <div class="val">{len(puntos)}</div>
                    <div class="lbl">Puntos totales</div>
                </div>
                <div class="metric-card">
                    <div class="val">{con_item}</div>
                    <div class="lbl">Con item</div>
                </div>
                <div class="metric-card blue">
                    <div class="val">{sin_item}</div>
                    <div class="lbl">Sin item</div>
                </div>
                <div class="metric-card blue">
                    <div class="val">{nombre_zona}</div>
                    <div class="lbl">Zona UTM</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Vista previa
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="section-label">🔍 Vista previa — primeros 10 puntos</div>',
                        unsafe_allow_html=True)
            df = pd.DataFrame(puntos)
            st.dataframe(df.head(10), use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Descarga
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="section-label">⬇ Exportar</div>', unsafe_allow_html=True)
            xlsx_bytes    = generar_xlsx(puntos, nombre_zona)
            nombre_salida = archivo.name.replace(".kmz", f"_UTM_{nombre_zona}.xlsx")
            st.download_button(
                label=f"DESCARGAR EXCEL  ·  {len(puntos)} PUNTOS  ·  ZONA {nombre_zona}",
                data=xlsx_bytes,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

else:
    st.markdown("""
    <div class="card">
        <div class="empty-state">
            <div class="icon">📂</div>
            Sube un archivo <strong>.kmz</strong> para comenzar
        </div>
    </div>
    """, unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────
st.markdown(f"""
<div class="footer">
    <strong>Redsetel</strong> · Red de Servicios y Telecomunicaciones Perú<br>
    Herramienta interna · Extracción KMZ · Conversión WGS84 → UTM · Zonas 17–19 N/S
</div>
""", unsafe_allow_html=True)
