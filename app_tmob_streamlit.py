# -*- coding: utf-8 -*-
"""
TMOB – Importador Web
Aplicación principal con navegación a Incidencias y Peticiones/WO.
"""
import streamlit as st

st.set_page_config(
    page_title="TMOB – Importador Web",
    page_icon="📥",
    layout="wide",
)

# ============================================================
# AUTENTICACIÓN
# ============================================================

if "tmob_authenticated" not in st.session_state:
    st.session_state.tmob_authenticated = False

if not st.session_state.tmob_authenticated:

    st.title("📥 TMOB – Importador Web")
    st.caption("Acceso al sistema de importación TMOB")

    code = st.text_input(
        "Código de acceso",
        type="password"
    )

    if st.button(
        "Acceder",
        type="primary",
        use_container_width=True
    ):
        configured_code = st.secrets.get(
            "TMOB_ACCESS_CODE",
            ""
        )

        if configured_code and code == configured_code:
            st.session_state.tmob_authenticated = True
            st.rerun()
        else:
            st.error("Código de acceso incorrecto.")

    st.stop()


# ============================================================
# PÁGINAS
# ============================================================

incidencias = st.Page(
    "01_Incidencias.py",
    title="Incidencias",
    icon="📋"
)

peticiones = st.Page(
    "02_Peticiones_WO.py",
    title="Peticiones / WO",
    icon="📦"
)


# ============================================================
# NAVEGACIÓN
# ============================================================

pg = st.navigation(
    {
        "TMOB – Importador Web": [
            incidencias,
            peticiones
        ]
    }
)

pg.run()
