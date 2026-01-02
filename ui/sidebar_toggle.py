import streamlit as st

def sidebar_toggle():
    # Estado
    if "sidebar_open" not in st.session_state:
        st.session_state.sidebar_open = True

    # Botón (mismo diseño que Home)
    col_btn, col_space = st.columns([0.06, 0.94])
    with col_btn:
        label = "«" if st.session_state.sidebar_open else "☰"
        if st.button(label, help="Mostrar / Ocultar menú"):
            st.session_state.sidebar_open = not st.session_state.sidebar_open
            st.rerun()

    # Asegura render del sidebar
    with st.sidebar:
        st.markdown("")

    # CSS fuerte: abrir/cerrar
    if st.session_state.sidebar_open:
        st.markdown(
            """
            <style>
            [data-testid="stSidebar"]{
                display: block !important;
                visibility: visible !important;
                transform: translateX(0) !important;
                width: 21rem !important;
                min-width: 21rem !important;
                max-width: 21rem !important;
            }
            [data-testid="stSidebar"] > div:first-child{ width: 21rem !important; }
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
