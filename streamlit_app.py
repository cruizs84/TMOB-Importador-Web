
import io
import os
import requests
import pandas as pd
import streamlit as st
from requests.auth import HTTPBasicAuth

PROJECT_KEY = "TMOB"
ISSUE_TYPE = "Incidencia"

FIELD_ID_ORIGEN = "customfield_10402"
FIELD_REQ = "customfield_10296"
FIELD_AMBITO = "customfield_10295"
FIELD_GRUPO = "customfield_10330"
FIELD_MOTIVO = "customfield_10332"
FIELD_FECHA_INICIO = "customfield_10472"
FIELD_EMPRESA = "customfield_10369"

AMBITO_PROVISIONAL = "Pendiente de asignación"

GROUP_MAP = {
    "MTC": "ATM – MTC",
    "SAT.Equips de camp": "SAT.Equips de camp",
    "SAT.Software": "SAT.Software",
    "App.Indra": "App.Indra",
    "App.Fujitsu.SIT": "App.Fujitsu.SIT",
    "Logística Soportes": "Logistica de soportes",
    "Service Desk": "Service Desk",
    "Seguretat": "Seguretat",
    "SEGURIDAD Y COMUNICACIONES": "Seguridad y comunicaciones SPOC",
    "Smarting": "Smarting",
}

def clean(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None

def load_excel(upload):
    df = pd.read_excel(upload, sheet_name="Report", header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={
        "ID de la incidencia*+": "ID incidencia origen",
        "ID de petición de servicio": "REQ incidencia",
        "Empresa*+": "Empresa",
        "Prioridad*": "Prioridad",
        "Fecha de envío": "Fecha y hora de inicio",
        "Fecha de última modificación": "Fecha última modificación origen",
        "Resumen*": "Resumen",
        "Grupo asignado*+": "Grupo externo resolutor origen",
        "Estado*": "Estado origen",
        "Status_Reason_Hidden": "Motivo de Pendiente",
    })
    df = df[df["ID incidencia origen"].notna()].copy()
    df["ID incidencia origen"] = df["ID incidencia origen"].astype(str).str.strip()
    df["Grupo externo resolutor"] = (
        df["Grupo externo resolutor origen"].astype("string").str.strip().map(GROUP_MAP)
    )
    df["Fecha y hora de inicio"] = pd.to_datetime(
        df["Fecha y hora de inicio"], errors="coerce"
    )
    df["Motivo de Pendiente"] = df["Motivo de Pendiente"].astype("string").str.strip()
    return df

def jira_session(base_url, email, token):
    base_url = base_url.rstrip("/")
    s = requests.Session()
    s.auth = HTTPBasicAuth(email, token)
    s.headers.update({"Accept": "application/json", "Content-Type": "application/json"})
    return base_url, s

def get_create_metadata(base_url, session):
    # Jira Cloud REST v3 endpoint for create metadata.
    url = f"{base_url}/rest/api/3/issue/createmeta/{PROJECT_KEY}/issuetypes"
    r = session.get(url, params={"maxResults": 100}, timeout=30)
    if not r.ok:
        raise RuntimeError(f"No se pudo leer metadata de creación: HTTP {r.status_code} {r.text[:1000]}")
    return r.json()

def get_field_options(base_url, session, field_id):
    # Used only if needed; the create metadata endpoint is preferred.
    url = f"{base_url}/rest/api/3/field/{field_id}/context/option"
    r = session.get(url, params={"maxResults": 100}, timeout=30)
    if not r.ok:
        return []
    return r.json().get("values", [])

def esc_jql(v):
    return str(v).replace("\\", "\\\\").replace('"', '\\"')

def exists_issue(base_url, session, id_origen):
    jql = (
        f'project = {PROJECT_KEY} '
        f'AND issuetype = "{ISSUE_TYPE}" '
        f'AND "{FIELD_ID_ORIGEN}" = "{esc_jql(id_origen)}"'
    )
    # POST search/jql is used to avoid URL length issues.
    url = f"{base_url}/rest/api/3/search/jql"
    r = session.post(
        url,
        json={"jql": jql, "maxResults": 2, "fields": ["summary", FIELD_ID_ORIGEN]},
        timeout=30,
    )
    if not r.ok:
        raise RuntimeError(f"Error buscando {id_origen}: HTTP {r.status_code} {r.text[:1000]}")
    issues = r.json().get("issues", [])
    return issues[0]["key"] if issues else None

def iso_datetime(v):
    if pd.isna(v):
        return None
    dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return None
    # The source is local Barcelona time; keep the local clock explicitly.
    return dt.to_pydatetime().strftime("%Y-%m-%dT%H:%M:%S+02:00")

def option_by_name(options, wanted):
    wanted_norm = wanted.strip().casefold()
    for o in options:
        if str(o.get("value", "")).strip().casefold() == wanted_norm:
            return {"id": str(o["id"])}
    return None

def create_issue(base_url, session, row, catalog):
    id_origen = clean(row["ID incidencia origen"])
    req = clean(row["REQ incidencia"])
    summary = clean(row["Resumen"])
    company = clean(row["Empresa"])
    priority = clean(row["Prioridad"])
    group = clean(row["Grupo externo resolutor"])
    reason = clean(row["Motivo de Pendiente"])
    start = iso_datetime(row["Fecha y hora de inicio"])

    if not id_origen or not summary or not req or not start:
        raise ValueError(f"{id_origen}: faltan campos obligatorios de origen.")

    ambit = option_by_name(catalog["ambito"], AMBITO_PROVISIONAL)
    if not ambit:
        raise ValueError(
            f'No existe la opción de Ámbito "{AMBITO_PROVISIONAL}" en Jira. '
            "Añádela al catálogo antes de ejecutar la prueba real."
        )

    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"name": ISSUE_TYPE},
        "summary": summary,
        FIELD_ID_ORIGEN: id_origen,
        FIELD_REQ: req,
        FIELD_AMBITO: ambit,
        FIELD_FECHA_INICIO: start,
    }

    if company:
        company_opt = option_by_name(catalog["empresa"], company)
        if not company_opt:
            raise ValueError(f"{id_origen}: Empresa no encontrada en Jira: {company}")
        fields[FIELD_EMPRESA] = company_opt

    if priority:
        fields["priority"] = {"name": priority}

    if group:
        group_opt = option_by_name(catalog["grupo"], group)
        if not group_opt:
            raise ValueError(f"{id_origen}: Grupo externo no encontrado en Jira: {group}")
        fields[FIELD_GRUPO] = group_opt

    if reason:
        reason_opt = option_by_name(catalog["motivo"], reason)
        if not reason_opt:
            # Do not invent or silently translate a reason.
            raise ValueError(f"{id_origen}: Motivo de Pendiente no encontrado en Jira: {reason}")
        fields[FIELD_MOTIVO] = reason_opt

    # assignee deliberately omitted.
    r = session.post(f"{base_url}/rest/api/3/issue", json={"fields": fields}, timeout=30)
    if not r.ok:
        raise RuntimeError(f"{id_origen}: error creando Jira HTTP {r.status_code}: {r.text[:1500]}")
    return r.json().get("key")

def build_catalog(base_url, session):
    # Get issue type id first.
    r = session.get(f"{base_url}/rest/api/3/issue/createmeta/{PROJECT_KEY}/issuetypes",
                     params={"maxResults": 100}, timeout=30)
    if not r.ok:
        raise RuntimeError(f"Error leyendo tipos de incidencia: HTTP {r.status_code} {r.text[:1000]}")
    data = r.json()
    types = data.get("issueTypes", data.get("values", []))
    inc = next((x for x in types if x.get("name") == ISSUE_TYPE), None)
    if not inc:
        raise RuntimeError("No se encontró el tipo Incidencia en TMOB.")

    issue_type_id = inc["id"]
    r = session.get(
        f"{base_url}/rest/api/3/issue/createmeta/{PROJECT_KEY}/issuetypes/{issue_type_id}",
        params={"maxResults": 100}, timeout=30
    )
    if not r.ok:
        raise RuntimeError(f"Error leyendo campos de Incidencia: HTTP {r.status_code} {r.text[:1000]}")
    fields = r.json().get("fields", [])

    def vals(field_id):
        f = next((x for x in fields if x.get("fieldId") == field_id or x.get("key") == field_id), None)
        return f.get("allowedValues", []) if f else []

    return {
        "ambito": vals(FIELD_AMBITO),
        "empresa": vals(FIELD_EMPRESA),
        "grupo": vals(FIELD_GRUPO),
        "motivo": vals(FIELD_MOTIVO),
    }

st.set_page_config(page_title="TMOB · Importador de incidencias", page_icon="🛠️", layout="wide")
st.title("🛠️ TMOB · Importador web de incidencias")
st.caption("Excel → Jira Cloud · control de duplicados · creación inicial en CREADA")

with st.sidebar:
    st.header("Conexión Jira")
    base_url = st.text_input("Jira URL", value="https://support-secom.atlassian.net")
    email = st.text_input("Email Jira")
    token = st.text_input("API Token", type="password")
    st.info("El token se usa en la sesión de la aplicación y no se escribe en el Excel de resultados.")

upload = st.file_uploader("Sube INC Activas.xlsx", type=["xlsx"])

if upload:
    try:
        df = load_excel(upload)
        st.success(f"Excel cargado: {len(df)} incidencias con ID origen.")
        st.dataframe(
            df[["ID incidencia origen", "REQ incidencia", "Resumen", "Empresa",
                "Prioridad", "Grupo externo resolutor", "Estado origen",
                "Motivo de Pendiente"]].head(20),
            use_container_width=True
        )

        limit = st.number_input("Número de incidencias para la prueba", min_value=1, max_value=len(df), value=min(5, len(df)))
        test_df = df.head(int(limit)).copy()

        st.subheader("Comportamiento de importación")
        st.write("**Nueva incidencia:** CREADA · Ámbito = Pendiente de asignación · Persona asignada = vacía")
        st.write("**Si ID incidencia origen ya existe:** OMITIR")
        st.write("**Status_Reason_Hidden:** se copia a Motivo de Pendiente cuando el valor existe.")

        if st.button("🔎 Comprobar duplicados (sin crear)", type="secondary"):
            if not email or not token:
                st.error("Introduce Email Jira y API Token.")
            else:
                try:
                    base, sess = jira_session(base_url, email, token)
                    rows = []
                    for _, row in test_df.iterrows():
                        id0 = clean(row["ID incidencia origen"])
                        existing = exists_issue(base, sess, id0)
                        rows.append({
                            "ID incidencia origen": id0,
                            "Resultado": "OMITIR" if existing else "CREAR",
                            "Jira existente": existing or "",
                        })
                    st.dataframe(pd.DataFrame(rows), use_container_width=True)
                except Exception as e:
                    st.error(str(e))

        st.warning("La creación REAL escribe en Jira. Haz primero la comprobación de duplicados.")
        if st.button("🚀 EJECUTAR PRUEBA REAL", type="primary"):
            if not email or not token:
                st.error("Introduce Email Jira y API Token.")
            else:
                try:
                    base, sess = jira_session(base_url, email, token)
                    catalog = build_catalog(base, sess)
                    results = []

                    for _, row in test_df.iterrows():
                        id0 = clean(row["ID incidencia origen"])
                        try:
                            existing = exists_issue(base, sess, id0)
                            if existing:
                                results.append({
                                    "ID incidencia origen": id0,
                                    "Resultado": "OMITIDA",
                                    "Jira Key": existing,
                                    "Estado inicial": "",
                                    "Error": "",
                                })
                                continue

                            key = create_issue(base, sess, row, catalog)
                            results.append({
                                "ID incidencia origen": id0,
                                "Resultado": "CREADA",
                                "Jira Key": key,
                                "Estado inicial": "CREADA",
                                "Error": "",
                            })
                        except Exception as e:
                            results.append({
                                "ID incidencia origen": id0,
                                "Resultado": "ERROR",
                                "Jira Key": "",
                                "Estado inicial": "",
                                "Error": str(e),
                            })

                    out = pd.DataFrame(results)
                    st.dataframe(out, use_container_width=True)
                    st.download_button(
                        "⬇️ Descargar resultado",
                        out.to_csv(index=False).encode("utf-8-sig"),
                        "TMOB_resultado_prueba.csv",
                        "text/csv"
                    )
                except Exception as e:
                    st.error(str(e))
    except Exception as e:
        st.error(f"Error cargando o procesando el Excel: {e}")
else:
    st.info("Carga el Excel INC Activas.xlsx para empezar.")
