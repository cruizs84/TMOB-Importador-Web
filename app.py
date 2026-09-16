
import io
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="TMOB - Importador Web", page_icon="📥", layout="wide")

# ============================================================
# CONFIGURACIÓN
# ============================================================

JIRA_API_BASE = "/rest/api/3"
PROJECT_KEY = "TMOB"
ISSUE_TYPE_NAME = "Incidencia"
TIMEZONE = ZoneInfo("Europe/Madrid")

# Grupo externo resolutor:
# Las claves son los valores del Excel. Los valores son los IDs
# reales de las opciones de Jira. NO se usa el texto directamente.
GROUP_EXTERNAL_MAP = {
    "MTC": "10149",
    "ATM – MTC": "10149",
    "ATM - MTC": "10149",
    "SAT.Equips de camp": "10168",
    "SAT.Software": "10169",
    "App.Indra": "10164",
    "App.Fujitsu.SIT": "10163",
    "Logística Soportes": "10151",
    "Logistica Soportes": "10151",
    "Logistica de soportes": "10151",
    "Service Desk": "10159",
    "Seguretat": "10158",
    "SEGURIDAD Y COMUNICACIONES": "10173",
    "Seguridad y comunicaciones SPOC": "10173",
    "Smarting": "10176",
}

# Empresa: claves = valores que pueden aparecer en el Excel.
# Los valores son los IDs reales de Jira.
COMPANY_MAP = {
    "ATM": "10221",
    "FGC": "10222",
    "RENFE": "10223",
    "TMB Sistemas Distribuits": "10224",
    "TRAM Baix": "10225",
    "TRAM BAIX": "10225",
    "TRAM BAIX": "10225",
    "SOLER I SAURET": "10226",
    "SOC Mobilitat": "10227",
    "FUJITSU": "10228",
    "LAC": "10229",
    "INDRA": "10230",
    "TUSGSAL": "10231",
    "Alsina Graells de Auto Trans": "10232",
    "AUTOCORB": "10233",
    "SAGALES": "10234",
    "Transports Generals d'Olesa": "10235",
    "Autocars Julià": "10236",
    "Empresa Casas, SA": "10237",
    "UTE Monbus Port": "10238",
    "Sarbus (Marfina Bus, SA)": "10239",
    "TUS, S. Coop. CL (Sabadell)": "10240",
    "AVANZA": "10241",
    "TMESA": "10242",
    "SMARTING": "10243",
    "CINTOI BUS": "10244",
    "EMPRESA PLANA, S.L.": "10245",
    "MASATS": "10246",
    "TMB Metro": "10247",
    "TB": "10248",
    "TRAM Besòs": "10249",
    "TRAM": "10250",
    "FONT": "10251",
    "AMB": "10304",
    "AMTU": "10305",
    "Autocares Prat": "10306",
    "Autocares PRAT": "10306",
    "Autocares Rovira": "10307",
    "Autocars Vendrell": "10308",
    "Autos Castellbisbal": "10309",
    "BUS CASTELLVI": "10310",
    "CIRCUITOS RURI": "10311",
    "CTSA-Mataró BUS": "10312",
    "HIFE": "10313",
    "Hispano Llacunense": "10314",
    "Mohn": "10315",
    "MOHN": "10315",
    "MOHN (Grup Baixbús)": "10315",
    "Montferri": "10316",
    "ROSABNBUS (Grup Baixbus)": "10317",
    "RubiBus": "10318",
    "TCC": "10319",
    "TEISA": "10320",
    "Transport MIR": "10321",
    "Urbà Vilafranca": "10322",
}

# Motivo de pendiente: aquí sí exigimos equivalencia exacta.
# El caso "Incidencia original pendiente" del Excel NO se traduce
# silenciosamente: requiere una equivalencia explícita.
PENDING_REASON_MAP = {
    "Acc. neces. de proveedor ext.": "10177",
    "Acción del cliente requerida": "10178",
    "Autorización de registro": "10181",
    "Cambio en la infraestructura": "10182",
    "Futura mejora": "10184",
    # Jira usa "Incidencia principal pendiente". El origen puede
    # contener "Incidencia original pendiente"; se deja fuera
    # deliberadamente para que el importador lo marque como error
    # de mapeo en lugar de inventar una equivalencia.
}

PRIORITY_MAP = {
    "Crítica": "10000",
    "Critica": "10000",
    "Alta": "10001",
    "Media": "10002",
    "Baja": "10003",
}

REQUIRED_COLUMNS = [
    "ID de la incidencia*+",
    "ID de petición de servicio",
    "Empresa*+",
    "Prioridad*",
    "Fecha de envío",
    "Fecha de última modificación",
    "Resumen*",
    "Grupo asignado*+",
    "Estado*",
    "Status_Reason_Hidden",
]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def get_settings():
    try:
        return (
            st.secrets["JIRA_URL"].rstrip("/"),
            st.secrets["JIRA_EMAIL"],
            st.secrets["JIRA_API_TOKEN"],
            st.secrets.get("TMOB_ACCESS_CODE", ""),
        )
    except Exception as exc:
        st.error(f"No se han podido leer los Secrets de Streamlit: {exc}")
        st.stop()


def jira_session():
    jira_url, email, token, _ = get_settings()
    s = requests.Session()
    s.auth = (email, token)
    s.headers.update({
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    return jira_url, s


def jira_get(path, params=None):
    jira_url, s = jira_session()
    return s.get(jira_url + path, params=params, timeout=30)


def jira_post(path, payload):
    jira_url, s = jira_session()
    return s.post(jira_url + path, json=payload, timeout=60)


def normalize_excel(df):
    # El fichero definitivo usa la primera fila como cabecera.
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Faltan columnas obligatorias en la hoja Report: "
            + ", ".join(missing)
        )

    df = df.copy()
    id_col = "ID de la incidencia*+"

    # Solo se consideran registros cuyo ID empieza por INC.
    df[id_col] = df[id_col].map(clean_text)
    df = df[df[id_col].str.match(r"^INC\d+$", na=False)].copy()
    df.reset_index(drop=True, inplace=True)
    return df


def parse_datetime_madrid(value):
    if pd.isna(value) or clean_text(value) == "":
        return None

    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None

    # Excel normalmente entrega Timestamp sin zona.
    # Se interpreta como hora local Europe/Madrid y se serializa
    # con el offset correspondiente (+01:00 o +02:00).
    if getattr(ts, "tzinfo", None) is None:
        ts = ts.tz_localize(TIMEZONE)
    else:
        ts = ts.tz_convert(TIMEZONE)

    return ts.isoformat()


def resolve_option(mapping, raw_value):
    raw = clean_text(raw_value)
    if raw == "":
        return None, None
    option_id = mapping.get(raw)
    return option_id, raw


def check_connection():
    results = {}

    r = jira_get("/rest/api/3/myself")
    results["myself"] = (r.status_code, r.text[:500])

    r = jira_get(f"/rest/api/3/project/{PROJECT_KEY}")
    results["project"] = (r.status_code, r.text[:500])

    return results


def search_duplicate(source_id):
    # Buscar por ID incidencia origen. Se escapa el texto para JQL.
    safe = source_id.replace("\\", "\\\\").replace('"', '\\"')
    jql = f'project = {PROJECT_KEY} AND "ID incidencia origen" ~ "\\"{safe}\\""'

    # Endpoint moderno de Jira Cloud.
    r = jira_post("/rest/api/3/search/jql", {
        "jql": jql,
        "maxResults": 20,
        "fields": ["summary", "status", "customfield_10402"],
    })

    if r.status_code == 200:
        data = r.json()
        return data.get("issues", [])

    # Fallback para instalaciones donde el endpoint moderno no esté disponible.
    r2 = jira_get("/rest/api/3/search", {
        "jql": jql,
        "maxResults": 20,
        "fields": "summary,status,customfield_10402",
    })
    if r2.status_code == 200:
        return r2.json().get("issues", [])

    raise RuntimeError(
        f"Error buscando duplicados HTTP {r.status_code}: {r.text[:1000]}"
    )


def validate_row(row):
    source_id = clean_text(row["ID de la incidencia*+"])
    req = clean_text(row["ID de petición de servicio"])
    company_raw = clean_text(row["Empresa*+"])
    priority_raw = clean_text(row["Prioridad*"])
    summary = clean_text(row["Resumen*"])
    group_raw = clean_text(row["Grupo asignado*+"])
    reason_raw = clean_text(row["Status_Reason_Hidden"])

    errors = []

    if not source_id:
        errors.append("ID incidencia origen vacío")
    if not req:
        errors.append("REQ incidencia vacío")
    if not summary:
        errors.append("Resumen vacío")

    company_id, _ = resolve_option(COMPANY_MAP, company_raw)
    if company_raw and not company_id:
        errors.append(f"Empresa sin mapeo explícito: {company_raw}")

    priority_id, _ = resolve_option(PRIORITY_MAP, priority_raw)
    if not priority_id:
        errors.append(f"Prioridad sin mapeo: {priority_raw}")

    group_id, _ = resolve_option(GROUP_EXTERNAL_MAP, group_raw)
    if group_raw and not group_id:
        errors.append(f"Grupo externo resolutor sin mapeo explícito: {group_raw}")

    reason_id, _ = resolve_option(PENDING_REASON_MAP, reason_raw)
    if reason_raw and not reason_id:
        errors.append(
            f"Motivo de pendiente sin mapeo explícito: {reason_raw}"
        )

    start_iso = parse_datetime_madrid(row["Fecha de envío"])
    if not start_iso:
        errors.append("Fecha de envío inválida o vacía")

    return {
        "source_id": source_id,
        "req": req,
        "company_raw": company_raw,
        "company_id": company_id,
        "priority_raw": priority_raw,
        "priority_id": priority_id,
        "summary": summary,
        "group_raw": group_raw,
        "group_id": group_id,
        "reason_raw": reason_raw,
        "reason_id": reason_id,
        "start_iso": start_iso,
        "errors": errors,
    }


def create_issue(v):
    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"name": ISSUE_TYPE_NAME},
        "summary": v["summary"],
        "description": "",
        "priority": {"id": v["priority_id"]},
        "customfield_10295": {"id": "10372"},  # Pendiente de asignación
        "customfield_10296": v["req"],
        "customfield_10402": v["source_id"],
        "customfield_10472": v["start_iso"],
    }

    if v["company_id"]:
        fields["customfield_10369"] = {"id": v["company_id"]}

    if v["group_id"]:
        fields["customfield_10330"] = {"id": v["group_id"]}

    if v["reason_id"]:
        fields["customfield_10332"] = {"id": v["reason_id"]}

    payload = {"fields": fields}
    r = jira_post("/rest/api/3/issue", payload)

    if r.status_code not in (200, 201):
        raise RuntimeError(f"HTTP {r.status_code}: {r.text[:2000]}")

    return r.json()


# ============================================================
# INTERFAZ
# ============================================================

st.title("📥 TMOB – Importador Web")
st.caption("Importación controlada de incidencias históricas a Jira Cloud")

# Acceso común
_, _, _, access_code = get_settings()

if access_code:
    entered = st.text_input("Código de acceso", type="password")
    if entered != access_code:
        st.info("Introduce el código de acceso para continuar.")
        st.stop()

st.success("Aplicación disponible.")

uploaded = st.file_uploader(
    "Carga el Excel definitivo",
    type=["xlsx"],
    help="Se leerá la hoja Report con la primera fila como cabecera."
)

if uploaded is None:
    st.info("Carga el fichero XLSX para comenzar.")
    st.stop()

try:
    df_raw = pd.read_excel(uploaded, sheet_name="Report", header=0)
    df = normalize_excel(df_raw)
except Exception as exc:
    st.error(f"Error leyendo el Excel: {exc}")
    st.stop()

st.success(f"Excel cargado: {len(df)} incidencias con ID origen válido.")

n = st.number_input(
    "Número de incidencias a procesar en esta prueba",
    min_value=1,
    max_value=len(df),
    value=min(5, len(df)),
    step=1,
)

test_df = df.iloc[:int(n)].copy()

st.subheader("Previsualización")
st.dataframe(
    test_df[
        [
            "ID de la incidencia*+",
            "ID de petición de servicio",
            "Empresa*+",
            "Prioridad*",
            "Resumen*",
            "Grupo asignado*+",
            "Status_Reason_Hidden",
        ]
    ],
    use_container_width=True,
)

st.divider()

if st.button("🔎 Comprobar conexión y duplicados (sin crear)", type="primary"):
    try:
        conn = check_connection()

        if conn["myself"][0] != 200:
            st.error(f"Jira /myself HTTP {conn['myself'][0]}: {conn['myself'][1]}")
            st.stop()

        if conn["project"][0] != 200:
            st.error(
                f"Jira proyecto {PROJECT_KEY} HTTP {conn['project'][0]}: "
                f"{conn['project'][1]}"
            )
            st.stop()

        st.success(f"Conexión Jira OK · proyecto {PROJECT_KEY}")

        rows = []
        for _, row in test_df.iterrows():
            v = validate_row(row)

            jira_exists = []
            duplicate_error = None
            try:
                jira_exists = search_duplicate(v["source_id"])
            except Exception as exc:
                duplicate_error = str(exc)

            if v["errors"]:
                result = "ERROR DE MAPEO"
                validation = " | ".join(v["errors"])
            elif duplicate_error:
                result = "ERROR"
                validation = duplicate_error
            elif jira_exists:
                result = "OMITIR"
                validation = "Ya existe en Jira"
            else:
                result = "CREAR"
                validation = "OK"

            rows.append({
                "ID incidencia origen": v["source_id"],
                "Resultado": result,
                "Jira existente": ", ".join(
                    i.get("key", "") for i in jira_exists
                ),
                "Validación": validation,
                "Grupo Excel": v["group_raw"],
                "Grupo Jira ID": v["group_id"] or "",
                "Empresa Excel": v["company_raw"],
                "Empresa Jira ID": v["company_id"] or "",
            })

        st.session_state["precheck"] = pd.DataFrame(rows)

    except Exception as exc:
        st.error(f"Error en la precomprobación: {exc}")

if "precheck" in st.session_state:
    pre = st.session_state["precheck"]
    st.subheader("Resultado de la precomprobación")
    st.dataframe(pre, use_container_width=True)

    errors = int((pre["Resultado"] == "ERROR").sum())
    mapping_errors = int((pre["Resultado"] == "ERROR DE MAPEO").sum())
    duplicates = int((pre["Resultado"] == "OMITIR").sum())
    creates = int((pre["Resultado"] == "CREAR").sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CREAR", creates)
    c2.metric("OMITIR", duplicates)
    c3.metric("ERRORES MAPEO", mapping_errors)
    c4.metric("ERRORES", errors)

    if errors == 0 and mapping_errors == 0:
        st.success("Precomprobación correcta. Puedes ejecutar la prueba real.")
    else:
        st.warning(
            "Hay registros que no deben crearse hasta corregir el mapeo o el error."
        )

    if st.button(
        "🚀 EJECUTAR PRUEBA REAL",
        type="secondary",
        disabled=(errors > 0 or mapping_errors > 0 or creates == 0),
    ):
        results = []

        for _, item in pre.iterrows():
            source_id = item["ID incidencia origen"]

            if item["Resultado"] == "OMITIR":
                results.append({
                    "ID incidencia origen": source_id,
                    "Resultado": "OMITIR",
                    "Jira Key": "",
                    "Estado inicial": "",
                    "Error": "Ya existía en Jira",
                })
                continue

            if item["Resultado"] != "CREAR":
                results.append({
                    "ID incidencia origen": source_id,
                    "Resultado": item["Resultado"],
                    "Jira Key": "",
                    "Estado inicial": "",
                    "Error": item["Validación"],
                })
                continue

            source_row = test_df[
                test_df["ID de la incidencia*+"].astype(str).str.strip()
                == str(source_id).strip()
            ]

            if source_row.empty:
                results.append({
                    "ID incidencia origen": source_id,
                    "Resultado": "ERROR",
                    "Jira Key": "",
                    "Estado inicial": "",
                    "Error": "No se encontró la fila original del Excel.",
                })
                continue

            v = validate_row(source_row.iloc[0])

            try:
                created = create_issue(v)
                results.append({
                    "ID incidencia origen": source_id,
                    "Resultado": "CREADO",
                    "Jira Key": created.get("key", ""),
                    "Estado inicial": "CREADA",
                    "Error": "",
                })
            except Exception as exc:
                results.append({
                    "ID incidencia origen": source_id,
                    "Resultado": "ERROR",
                    "Jira Key": "",
                    "Estado inicial": "",
                    "Error": f"{source_id}: error creando Jira {exc}",
                })

        st.session_state["creation_results"] = pd.DataFrame(results)

if "creation_results" in st.session_state:
    st.subheader("Resultado de la carga")
    st.dataframe(st.session_state["creation_results"], use_container_width=True)
