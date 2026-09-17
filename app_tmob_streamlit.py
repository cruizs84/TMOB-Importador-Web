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

if "tmob_authenticated" not in st.session_state:
    st.session_state.tmob_authenticated = False

if not st.session_state.tmob_authenticated:
    st.title("📥 TMOB – Importador Web")
    st.subheader("Acceso")
    code = st.text_input("Código de acceso", type="password")

    if st.button("Acceder", type="primary"):
        configured = st.secrets.get("TMOB_ACCESS_CODE", "")
        if configured and code == configured:
            st.session_state.tmob_authenticated = True
            st.rerun()
        else:
            st.error("Código de acceso incorrecto.")
    st.stop()

st.title("📥 TMOB – Importador Web")
st.caption("Selecciona el tipo de carga que quieres realizar.")
st.divider()

st.subheader("¿Qué quieres importar?")

col1, col2 = st.columns(2)

with col1:
    st.markdown("## 📋 Incidencias")
    st.write("Importador estable de incidencias históricas/activas.")
    st.write("Tipo Jira: **Incidencia**")
    if st.button("Abrir Incidencias", type="primary", use_container_width=True):
        st.switch_page("01_Incidencias.py")

with col2:
    st.markdown("## 📦 Peticiones / WO")
    st.write("Importador de órdenes de trabajo.")
    st.write("Tipo Jira: **Tarea por incidencia o petición**")
    if st.button("Abrir Peticiones / WO", type="primary", use_container_width=True):
        st.switch_page("02_Peticiones_WO.py")

st.divider()

st.info(
    "Los dos módulos utilizan la misma conexión Jira y el mismo código de acceso. "
    "El catálogo Empresa se mantiene en Jira; el importador de Peticiones no crea empresas."
)

with st.expander("⚙️ Configuración técnica"):
    st.write("Proyecto Jira: **TMOB**")
    st.write("URL Jira: **https://suport-secom.atlassian.net**")
    st.write("Secrets compartidos: TMOB_ACCESS_CODE, JIRA_URL, JIRA_EMAIL, JIRA_API_TOKEN")
