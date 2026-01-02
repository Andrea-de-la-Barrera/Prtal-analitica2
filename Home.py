import streamlit as st
from pathlib import Path

# =========================
# Configuración base
# =========================
st.set_page_config(
    page_title="Portal de Analítica",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Bootstrap icons
st.markdown(
    '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css">',
    unsafe_allow_html=True
)

# =========================
# Cargar CSS GLOBAL (DISEÑO)
# =========================
css_file = Path(__file__).parent / "styles" / "global.css"
if css_file.exists():
    st.markdown(f"<style>{css_file.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)

# =========================
# SIDEBAR TOGGLE (FUNCIONALIDAD)
# =========================
if "sidebar_open" not in st.session_state:
    st.session_state.sidebar_open = True

# Botón (solo lógica)
col_btn, col_space = st.columns([0.06, 0.94])
with col_btn:
    label = "«" if st.session_state.sidebar_open else "☰"
    if st.button(label, help="Mostrar / Ocultar menú"):
        st.session_state.sidebar_open = not st.session_state.sidebar_open
        st.rerun()

# 🟦 IMPORTANTE: asegurar que el sidebar EXISTE (a veces sin contenido no lo dibuja bien)
with st.sidebar:
    st.markdown("")  # no agrega nada visible, solo asegura render

# ✅ CSS "FUERTE" para abrir/cerrar de verdad
if st.session_state.sidebar_open:
    st.markdown(
        """
        <style>
        /* Forzar sidebar visible */
        [data-testid="stSidebar"]{
            display: block !important;
            visibility: visible !important;
            transform: translateX(0) !important;
            width: 21rem !important;           /* ancho sidebar */
            min-width: 21rem !important;
            max-width: 21rem !important;
        }
        [data-testid="stSidebar"] > div:first-child{
            width: 21rem !important;
        }
        [data-testid="stSidebarNav"]{
            display: block !important;
            visibility: visible !important;
            transform: translateX(0) !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <style>
        /* Ocultar sidebar completo */
        [data-testid="stSidebar"]{
            display: none !important;
            visibility: hidden !important;
            width: 0 !important;
        }
        [data-testid="stSidebarNav"]{
            display: none !important;
            visibility: hidden !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# =========================
# HERO
# =========================
st.markdown(
    """
    <div class="hero-wrap">
        <div class="pill"><span class="pill-dot"></span> EQ. BTS (UI)</div>
        <div class="hero-title">Portal de <span class="accent">Analítica</span></div>
        <div class="hero-sub">
            Ecosistema inteligente para procesamiento, diagnóstico y predicción de datos estratégicos.<br>
            Selecciona un módulo, sube tu CSV y descarga reportes en segundos.
        </div>
    </div>
    <div class="sep"></div>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="section-title">Nuestras Soluciones Modulares</div>', unsafe_allow_html=True)

# =========================
# Definición de módulos
# =========================
modules = [
    {
        "title": "Descriptiva",
        "desc": "KPIs, filtros, tendencias y categorías principales para entender tu histórico.",
        "icon": "bi-list-check",
        "path": "pages/1_Descriptiva.py",
    },
    {
        "title": "Predictiva",
        "desc": "Modelado estadístico y proyecciones de tendencias (regresión lineal simple).",
        "icon": "bi-graph-up",
        "path": "pages/2_Predictiva.py",
    },
    {
        "title": "Prescriptiva",
        "desc": "Recomendaciones y plan de ajuste por categoría para cumplir un presupuesto.",
        "icon": "bi-lightbulb",
        "path": "pages/3_Prescriptiva.py",
    },
    {
        "title": "Calidad de Datos",
        "desc": "Auditoría: nulos, duplicados, outliers, score y reporte descargable.",
        "icon": "bi-shield-check",
        "path": "pages/4_Calidad_Datos.py",
    },
    {
        "title": "Minería / Segmentación",
        "desc": "Clustering para descubrir patrones y segmentar comportamientos en tus datos.",
        "icon": "bi-cpu",
        "path": "pages/5_Mineria_Segmentacion.py",
    },
]

# =========================
# Render cards
# =========================
def render_card(m):
    st.markdown(
        f"""
        <div class="module-card">
            <div>
                <div class="icon-badge"><i class="bi {m['icon']}"></i></div>
                <h3>{m['title']}</h3>
                <p>{m['desc']}</p>
                <div class="filehint">{m['path']}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.page_link(m["path"], label="Acceder", use_container_width=True)

# Fila 1 (3)
c1, c2, c3 = st.columns(3)
with c1: render_card(modules[0])
with c2: render_card(modules[1])
with c3: render_card(modules[2])

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# Fila 2 (2)
c4, c5 = st.columns(2)
with c4: render_card(modules[3])
with c5: render_card(modules[4])

st.caption("© 2025 Portal de Analítica | UI dark (Streamlit multipage)")
