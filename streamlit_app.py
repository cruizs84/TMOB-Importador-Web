import io
import hmac
import requests
import pandas as pd
import streamlit as st
from requests.auth import HTTPBasicAuth

PROJECT_KEY = "TMOB"
ISSUE_TYPE = "Incidencia"
ISSUE_TYPE_ID = "10223"

FIELD_ID_ORIGEN = "customfield_10402"
FIELD_REQ = "customfield_10296"
FIELD_AMBITO = "customfield_10295"
FIELD_GRUPO = "customfield_10330"
FIELD_MOTIVO = "customfield_10332"
FIELD_FECHA_INICIO = "customfield_10472"
FIELD_EMPRESA = "customfield_10369"

AMBITO_PROVISIONAL = "Pendiente de asignación"
AMBITO_PROVISIONAL_ID = "10372"

# Option IDs verified in Jira TMOB.
GROUP_OPTIONS = {
    "MTC": "10149",
    "SAT.Equips de camp": "10168",
    "SAT.Software": "10169",
    "App.Indra": "10164",
    "App.Fujitsu.SIT": "10163",
    "Logística Soportes": "10151",
    "Service Desk": "10159",
    "Seguretat": "10158",
    "SEGURIDAD Y COMUNICACIONES": "10173",
    "Smarting": "10176",
}

# Jira's current Empresa catalog matches the active report values after normalization.
# Aliases cover known spelling/case differences without guessing new companies.
EMPRESA_ALIASES = {
    "Autocares Prat": "Autocares PRAT",
    "Autocars Julià": "Autocars Julià",
    "CIRCUITOS RURI": "CIRCUITOS RURI",
    "Cintoi Bus, SL": "CINTOI BUS",
    "CTSA-Mataró Bus": "CTSA-Mataró BUS",
    "Fujitsu": "FUJITSU",
    "Smarting": "SMARTING",
    "TMB Sistemes Distribuïts": "TMB Sistemas Distribuits",
    "Transports Generals d'Olesa": "Transports Generals d'Olesa",
    "TRAM Baix": "TRAM BAIX",
    "TUS, S. Coop. CL (Sabadell)": "TUS, S. Coop. CL (Sabadell)",
    "UTE Monbus Port": "UTE Monbus Port",
}

# Source reason is intentionally not translated. If Jira lacks an exact option,
# the import reports the row as ERROR so the catalog can be corrected first.


def clean(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def norm(v):
    return str(v or "").strip().casefold()


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
        "Nivel 1 de categorización operacional+": "N1 operacional",
        "Nivel 2 de categorización operacional": "N2 operacional",
        "Nivel 3 de categorización operacional": "N3 operacional",
        "Nivel 2 de categorización de producto": "N2 producto",
    })
    required = ["ID incidencia origen", "REQ incidencia", "Empresa", "Prioridad",
                "Fecha y hora de inicio", "Resumen"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Faltan columnas obligatorias en el Excel: " + ", ".join(missing))

    df = df[df["ID incidencia origen"].notna()].copy()
    df["ID incidencia origen"] = df["ID incidencia origen"].astype(str).str.strip()
    df["Grupo externo resolutor"] = (
        df.get("Grupo externo resolutor origen", pd.Series(index=df.index, dtype="string"))
        .astype("string").str.strip()
    )
    df["Fecha y hora de inicio"] = pd.to_datetime(df["Fecha y hora de inicio"], errors="coerce")
    df["Motivo de Pendiente"] = (
        df.get("Motivo de Pendiente", pd.Series(index=df.index, dtype="string"))
        .astype("string").str.strip()
    )
    return df


def jira_session(base_url):
    email = st.secrets.get("JIRA_EMAIL", "").strip()
    token = st.secrets.get("JIRA_API_TOKEN", "").strip()
    if not email or not token:
        raise RuntimeError("Faltan JIRA_EMAIL o JIRA_API_TOKEN en Streamlit Secrets.")
    base_url = (st.secrets.get("JIRA_URL", base_url) or base_url).rstrip("/")
    s = requests.Session()
    s.auth = HTTPBasicAuth(email, token)
    s.headers.update({"Accept": "application/json", "Content-Type": "application/json"})
    return base_url, s


def esc_jql(v):
    return str(v).replace("\\", "\\\\").replace('"', '\\"')


def exists_issue(base_url, session, id_origen):
    jql = (
        f'project = {PROJECT_KEY} AND issuetype = "{ISSUE_TYPE}" '
        f'AND "{FIELD_ID_ORIGEN}" = "{esc_jql(id_origen)}"'
    )
    r = session.post(
        f"{base_url}/rest/api/3/search/jql",
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
    return dt.to_pydatetime().strftime("%Y-%m-%dT%H:%M:%S+02:00")


def option_by_name(options, wanted):
    wanted_norm = norm(wanted)
    for o in options:
        if norm(o.get("value")) == wanted_norm:
            return {"id": str(o["id"])}
    return None


def get_allowed_options(base_url, session, field_id):
    # Current Jira REST metadata endpoint, scoped to TMOB/Incidencia.
    url = f"{base_url}/rest/api/3/issue/createmeta"
    params = {
        "projectKeys": PROJECT_KEY,
        "issuetypeIds": ISSUE_TYPE_ID,
        "expand": "projects.issuetypes.fields",
        "maxResults": 100,
    }
    r = session.get(url, params=params, timeout=30)
    if not r.ok:
        raise RuntimeError(f"Error leyendo metadata de Jira: HTTP {r.status_code} {r.text[:1000]}")
    data = r.json()
    projects = data.get("projects", [])
    if not projects:
        raise RuntimeError("Jira no devolvió metadata para TMOB/Incidencia.")
    issue_types = projects[0].get("issuetypes", [])
    issue_type = next((x for x in issue_types if str(x.get("id")) == ISSUE_TYPE_ID), None)
    if not issue_type:
        raise RuntimeError("Jira no devolvió metadata para el tipo Incidencia (10223).")
    field = issue_type.get("fields", {}).get(field_id, {})
    return field.get("allowedValues", [])


def build_catalog(base_url, session):
    return {
        "ambito": get_allowed_options(base_url, session, FIELD_AMBITO),
        "empresa": get_allowed_options(base_url, session, FIELD_EMPRESA),
        "grupo": get_allowed_options(base_url, session, FIELD_GRUPO),
        "motivo": get_allowed_options(base_url, session, FIELD_MOTIVO),
    }


def resolve_empresa(catalog, source_name):
    wanted = EMPRESA_ALIASES.get(source_name, source_name)
    return option_by_name(catalog["empresa"], wanted)


def resolve_group(catalog, source_name):
    if not source_name:
        return None
    mapped = GROUP_OPTIONS.get(source_name)
    if not mapped:
        raise ValueError(f"Grupo externo sin mapeo definido: {source_name}")
    for o in catalog["grupo"]:
        if str(o.get("id")) == mapped:
            return {"id": mapped}
    raise ValueError(f"Grupo externo no disponible en Jira: {source_name}")


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
        # Known ID is used as a safety fallback after the option was verified in Jira.
        if any(str(o.get("id")) == AMBITO_PROVISIONAL_ID for o in catalog["ambito"]):
            ambit = {"id": AMBITO_PROVISIONAL_ID}
        else:
            raise ValueError(f'No existe la opción de Ámbito "{AMBITO_PROVISIONAL}" en Jira.')

    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"id": ISSUE_TYPE_ID},
        "summary": summary,
        FIELD_ID_ORIGEN: id_origen,
        FIELD_REQ: req,
        FIELD_AMBITO: ambit,
        FIELD_FECHA_INICIO: start,
    }

    if company:
        company_opt = resolve_empresa(catalog, company)
        if not company_opt:
            raise ValueError(f"{id_origen}: Empresa no encontrada en Jira: {company}")
        fields[FIELD_EMPRESA] = company_opt

    if priority:
        fields["priority"] = {"name": priority}

    if group:
        fields[FIELD_GRUPO] = resolve_group(catalog, group)

    if reason:
        reason_opt = option_by_name(catalog["motivo"], reason)
        if not reason_opt:
            raise ValueError(
                f"{id_origen}: Motivo de Pendiente no encontrado exactamente en Jira: {reason}"
            )
        fields[FIELD_MOTIVO] = reason_opt

    # Assignee deliberately omitted: new incidents remain unassigned.
    r = session.post(f"{base_url}/rest/api/3/issue", json={"fields": fields}, timeout=30)
    if not r.ok:
        raise RuntimeError(f"{id_origen}: error creando Jira HTTP {r.status_code}: {r.text[:1500]}")
    return r.json().get("key")


def authenticate():
    if st.session_state.get("authenticated"):
        return True
    access_code = str(st.secrets.get("TMOB_ACCESS_CODE", ""))
    if not access_code:
        st.error("La aplicación no está configurada: falta TMOB_ACCESS_CODE en Streamlit Secrets.")
        return False

    st.title("🔐 TMOB · Acceso")
    st.caption("Introduce el código de acceso facilitado por el coordinador.")
    code = st.text_input("Código de acceso", type="password")
    if st.button("Entrar", type="primary"):
        if hmac.compare_digest(code, access_code):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Código incorrecto.")
    return False


st.set_page_config(page_title="TMOB · Importador de incidencias", page_icon="🛠️", layout="wide")

if not authenticate():
    st.stop()

with st.sidebar:
    st.header("TMOB · Importador")
    st.success("Acceso autorizado")
    if st.button("Cerrar sesión"):
        st.session_state["authenticated"] = False
        st.rerun()
    st.caption("Las credenciales Jira están almacenadas en Streamlit Secrets y no se muestran a los usuarios.")

st.title("🛠️ TMOB · Importador web de incidencias")
st.caption("Excel → Jira Cloud · control de duplicados · creación inicial en CREADA")

upload = st.file_uploader("Sube INC Activas.xlsx", type=["xlsx"])

if upload:
    try:
        df = load_excel(upload)
        st.success(f"Excel cargado: {len(df)} incidencias con ID origen.")
        preview_cols = [c for c in ["ID incidencia origen", "REQ incidencia", "Resumen", "Empresa",
                                     "Prioridad", "Grupo externo resolutor", "Estado origen",
                                     "Motivo de Pendiente"] if c in df.columns]
        st.dataframe(df[preview_cols].head(20), use_container_width=True)

        max_rows = min(5, len(df))
        limit = st.number_input(
            "Número de incidencias para la prueba",
            min_value=1, max_value=len(df), value=max_rows
        )
        test_df = df.head(int(limit)).copy()

        st.subheader("Comportamiento de importación")
        st.write("**Nueva incidencia:** CREADA · Ámbito = Pendiente de asignación · Persona asignada = vacía")
        st.write("**Si ID incidencia origen ya existe:** OMITIR")
        st.write("**Status_Reason_Hidden:** se copia a Motivo de Pendiente cuando existe una opción idéntica en Jira.")

        try:
            base, sess = jira_session("https://support-secom.atlassian.net")
        except Exception as e:
            st.error(str(e))
            st.stop()

        if st.button("🔎 Comprobar conexión y duplicados (sin crear)", type="secondary"):
            try:
                # Also validates that the new provisional Ámbito exists in Jira.
                catalog = build_catalog(base, sess)
                if not option_by_name(catalog["ambito"], AMBITO_PROVISIONAL):
                    raise RuntimeError(f'No se encontró Ámbito "{AMBITO_PROVISIONAL}" en Jira.')
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
                st.session_state["preflight_ok"] = True
                st.success("Precomprobación correcta. Puedes ejecutar la prueba real.")
            except Exception as e:
                st.session_state["preflight_ok"] = False
                st.error(str(e))

        st.warning("La creación REAL escribe en Jira. Haz primero la comprobación de conexión y duplicados.")
        if st.button("🚀 EJECUTAR PRUEBA REAL", type="primary"):
            if not st.session_state.get("preflight_ok"):
                st.error("Primero ejecuta 'Comprobar conexión y duplicados (sin crear)'.")
            else:
                try:
                    catalog = build_catalog(base, sess)
                    results = []
                    progress = st.progress(0)
                    total = len(test_df)
                    for idx, (_, row) in enumerate(test_df.iterrows(), start=1):
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
                            else:
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
                        progress.progress(idx / total)
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
