# -*- coding: utf-8 -*-
"""
TMOB – Importador Web integrado

Un único punto de entrada Streamlit para:
- Incidencias -> Jira: Incidencia (10223)
- Peticiones / WO -> Jira: Petición (10224)

Secrets esperados:
    TMOB_ACCESS_CODE
    JIRA_URL
    JIRA_EMAIL
    JIRA_API_TOKEN
"""

import re
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="TMOB - Importador Web",
    page_icon="📥",
    layout="wide",
)

JIRA_API_BASE = "/rest/api/3"
PROJECT_KEY = "TMOB"
TIMEZONE = ZoneInfo("Europe/Madrid")
EXPECTED_JIRA_URL = "https://suport-secom.atlassian.net"


# ============================================================
# CATÁLOGOS JIRA
# ============================================================

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

COMPANY_MAP = {
    "ATM": "10221",
    "FGC": "10222",
    "RENFE": "10223",
    "TMB Sistemes Distribuïts": "10224",
    "TRAM Baix": "10225",
    "TRAM BAIX": "10225",
    "SOLER I SAURET": "10226",
    "SOC Mobilitat": "10227",
    "Fujitsu": "10228",
    "LAC": "10229",
    "INDRA": "10230",
    "TUSGSAL": "10231",
    "Alsina Graells de Auto Trans": "10232",
    "Autocorb": "10233",
    "SAGALES": "10234",
    "Transports Generals d'Olesa": "10235",
    "Autocars Julià": "10236",
    "Empresa Casas, SA": "10237",
    "UTE Monbus Port": "10238",
    "Sarbus (Marfina Bus, SA)": "10239",
    "TUS, S. Coop. CL (Sabadell)": "10240",
    "AVANZA": "10241",
    "TMESA": "10242",
    "Smarting": "10243",
    "Cintoi Bus, SL": "10244",
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

PENDING_REASON_MAP = {
    "Acc. neces. de proveedor ext.": "10177",
    "Acción del cliente requerida": "10178",
    "Autorización de registro": "10181",
    "Cambio en la infraestructura": "10182",
    "Futura mejora": "10184",
    "Incidencia original pendiente": "10338",
}

PRIORITY_MAP = {
    "Crítica": "10000",
    "Critica": "10000",
    "Alta": "10001",
    "Media": "10002",
    "Baja": "10003",
}


# ============================================================
# CONFIGURACIÓN POR TIPO
# ============================================================

INC_CONFIG = {
    "title": "📋 Incidencias",
    "issue_type": "Incidencia",
    "issue_type_id": "10223",
    "id_column": "ID de la incidencia*+",
    "req_column": "ID de petición de servicio",
    "valid_id_regex": r"^INC\d+$",
    "header_row": 0,
    "caption": "Importación controlada de incidencias a Jira Cloud.",
    "required_columns": [
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
    ],
    "preview_columns": [
        "ID de la incidencia*+",
        "ID de petición de servicio",
        "Empresa*+",
        "Prioridad*",
        "Resumen*",
        "Grupo asignado*+",
        "Status_Reason_Hidden",
    ],
    "requires_group": True,
    "requires_pending_reason": True,
}

WO_CONFIG = {
    "title": "📦 Peticiones / WO",
    "issue_type": "Petición",
    "issue_type_id": "10224",
    "id_column": "ID de orden de trabajo+",
    "req_column": "ID de petición asociada",
    "priority_column": "Prioridad",
    "valid_id_regex": r"^WO\d+$",
    "header_row": 1,
    "caption": "Importación controlada de WO tratadas como Peticiones en Jira Cloud.",
    "required_columns": [
        "ID de orden de trabajo+",
        "ID de petición asociada",
        "Empresa*+",
        "Prioridad*",
        "Fecha de envío",
        "Fecha de última modificación",
        "Resumen*",
        "Estado*",
    ],
    "preview_columns": [
        "ID de orden de trabajo+",
        "ID de petición asociada",
        "Empresa*+",
        "Prioridad*",
        "Fecha de envío",
        "Resumen*",
        "Estado*",
    ],
    "requires_group": False,
    "requires_pending_reason": False,
}


# ============================================================
# UTILIDADES
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def get_settings():
    try:
        jira_url = st.secrets["JIRA_URL"].rstrip("/")
        email = st.secrets["JIRA_EMAIL"]
        token = st.secrets["JIRA_API_TOKEN"]
        access_code = st.secrets.get("TMOB_ACCESS_CODE", "")
        return jira_url, email, token, access_code
    except Exception as exc:
        st.error(f"No se han podido leer los Secrets de Streamlit: {exc}")
        st.stop()


def jira_session():
    jira_url, email, token, _ = get_settings()
    session = requests.Session()
    session.auth = (email, token)
    session.headers.update({
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    return jira_url, session


def jira_get(path, params=None):
    jira_url, session = jira_session()
    return session.get(jira_url + path, params=params, timeout=30)


def jira_post(path, payload):
    jira_url, session = jira_session()
    return session.post(jira_url + path, json=payload, timeout=60)


def parse_datetime_madrid(value):
    if pd.isna(value) or clean_text(value) == "":
        return None

    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None

    if getattr(ts, "tzinfo", None) is None:
        ts = ts.tz_localize(TIMEZONE)
    else:
        ts = ts.tz_convert(TIMEZONE)

    return ts.isoformat()


def resolve_option(mapping, raw_value):
    raw = clean_text(raw_value)
    if not raw:
        return None, raw
    return mapping.get(raw), raw


# ============================================================
# JIRA
# ============================================================

def check_connection():
    results = {}
    r = jira_get(f"{JIRA_API_BASE}/myself")
    results["myself"] = (r.status_code, r.text[:500])

    r = jira_get(f"{JIRA_API_BASE}/project/{PROJECT_KEY}")
    results["project"] = (r.status_code, r.text[:500])
    return results


def search_duplicate(source_id):
    safe = source_id.replace("\\", "\\\\").replace('"', '\\"')
    jql = (
        f'project = {PROJECT_KEY} '
        f'AND "ID incidencia origen" ~ "\\"{safe}\\""'
    )

    r = jira_post(
        f"{JIRA_API_BASE}/search/jql",
        {
            "jql": jql,
            "maxResults": 20,
            "fields": ["summary", "status", "customfield_10402"],
        },
    )

    if r.status_code == 200:
        return r.json().get("issues", [])

    r2 = jira_get(
        f"{JIRA_API_BASE}/search",
        {
            "jql": jql,
            "maxResults": 20,
            "fields": "summary,status,customfield_10402",
        },
    )

    if r2.status_code == 200:
        return r2.json().get("issues", [])

    raise RuntimeError(
        f"Error buscando duplicados HTTP {r.status_code}: {r.text[:1000]}"
    )


# ============================================================
# EXCEL
# ============================================================

def normalize_excel(df, config):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing = [
        col for col in config["required_columns"]
        if col not in df.columns
    ]
    if missing:
        raise ValueError(
            "Faltan columnas obligatorias en la hoja Report: "
            + ", ".join(missing)
        )

    id_column = config["id_column"]
    df[id_column] = df[id_column].map(clean_text)

    df = df[
        df[id_column].str.match(config["valid_id_regex"], na=False)
    ].copy()

    df.reset_index(drop=True, inplace=True)
    return df


# ============================================================
# VALIDACIÓN
# ============================================================

def validate_row(row, config):
    source_id = clean_text(row[config["id_column"]])
    req = clean_text(row[config["req_column"]])
    company_raw = clean_text(row["Empresa*+"])
    priority_raw = clean_text(row["Prioridad*"])
    summary = clean_text(row["Resumen*"])

    errors = []

    if not source_id:
        errors.append("ID origen vacío")

    if not req:
        errors.append("REQ / petición asociada vacío")

    if not summary:
        errors.append("Resumen vacío")

    company_id, _ = resolve_option(COMPANY_MAP, company_raw)
    if company_raw and not company_id:
        errors.append(f"Empresa sin mapeo explícito: {company_raw}")

    priority_id, _ = resolve_option(PRIORITY_MAP, priority_raw)
    if not priority_id:
        errors.append(f"Prioridad sin mapeo: {priority_raw}")

    start_iso = parse_datetime_madrid(row["Fecha de envío"])
    if not start_iso:
        errors.append("Fecha de envío inválida o vacía")

    group_raw = ""
    group_id = None
    reason_raw = ""
    reason_id = None

    if config["requires_group"]:
        group_raw = clean_text(row["Grupo asignado*+"])
        group_id, _ = resolve_option(GROUP_EXTERNAL_MAP, group_raw)
        if group_raw and not group_id:
            errors.append(
                "Grupo externo resolutor sin mapeo explícito: " + group_raw
            )

    if config["requires_pending_reason"]:
        reason_raw = clean_text(row["Status_Reason_Hidden"])
        reason_id, _ = resolve_option(PENDING_REASON_MAP, reason_raw)
        if reason_raw and not reason_id:
            errors.append(
                "Motivo de pendiente sin mapeo explícito: " + reason_raw
            )

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


# ============================================================
# CREACIÓN JIRA
# ============================================================

def create_issue(v, config):
    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"id": config["issue_type_id"]},
        "summary": v["summary"],
        "priority": {"id": v["priority_id"]},

        # Ámbito = Pendiente de asignación.
        "customfield_10295": {"id": "10372"},

        # REQ incidencia.
        "customfield_10296": v["req"],

        # ID de incidencia origen. Para WO contiene el ID WO.
        "customfield_10402": v["source_id"],

        # Fecha y hora de inicio.
        "customfield_10472": v["start_iso"],
    }

    if v["company_id"]:
        fields["customfield_10369"] = {"id": v["company_id"]}

    if v["group_id"]:
        fields["customfield_10330"] = {"id": v["group_id"]}

    if v["reason_id"]:
        fields["customfield_10332"] = {"id": v["reason_id"]}

    response = jira_post(
        f"{JIRA_API_BASE}/issue",
        {"fields": fields},
    )

    if response.status_code not in (200, 201):
        raise RuntimeError(
            f"HTTP {response.status_code}: {response.text[:2000]}"
        )

    return response.json()


# ============================================================
# PRECOMPROBACIÓN
# ============================================================

def run_precheck(test_df, config):
    conn = check_connection()

    if conn["myself"][0] != 200:
        raise RuntimeError(
            f"Jira /myself HTTP {conn['myself'][0]}: {conn['myself'][1]}"
        )

    if conn["project"][0] != 200:
        raise RuntimeError(
            f"Jira proyecto {PROJECT_KEY} HTTP {conn['project'][0]}: "
            f"{conn['project'][1]}"
        )

    rows = []

    for _, row in test_df.iterrows():
        validation = validate_row(row, config)

        jira_exists = []
        duplicate_error = None

        try:
            jira_exists = search_duplicate(validation["source_id"])
        except Exception as exc:
            duplicate_error = str(exc)

        if validation["errors"]:
            result = "ERROR DE MAPEO"
            validation_text = " | ".join(validation["errors"])
        elif duplicate_error:
            result = "ERROR"
            validation_text = duplicate_error
        elif jira_exists:
            result = "OMITIR"
            validation_text = "Ya existe en Jira"
        else:
            result = "CREAR"
            validation_text = "OK"

        rows.append({
            "ID origen": validation["source_id"],
            "Resultado": result,
            "Jira existente": ", ".join(
                issue.get("key", "") for issue in jira_exists
            ),
            "Validación": validation_text,
            "REQ": validation["req"],
            "Empresa Excel": validation["company_raw"],
            "Empresa Jira ID": validation["company_id"] or "",
        })

    return pd.DataFrame(rows)


# ============================================================
# EJECUCIÓN
# ============================================================

def run_creation(pre, test_df, config):
    results = []

    for _, item in pre.iterrows():
        source_id = item["ID origen"]

        if item["Resultado"] == "OMITIR":
            results.append({
                "ID origen": source_id,
                "Resultado": "OMITIR",
                "Jira Key": "",
                "Tipo Jira": "",
                "Estado inicial": "",
                "Error": "Ya existía en Jira",
            })
            continue

        if item["Resultado"] != "CREAR":
            results.append({
                "ID origen": source_id,
                "Resultado": item["Resultado"],
                "Jira Key": "",
                "Tipo Jira": config["issue_type"],
                "Estado inicial": "",
                "Error": item["Validación"],
            })
            continue

        source_row = test_df[
            test_df[config["id_column"]].astype(str).str.strip()
            == str(source_id).strip()
        ]

        if source_row.empty:
            results.append({
                "ID origen": source_id,
                "Resultado": "ERROR",
                "Jira Key": "",
                "Tipo Jira": config["issue_type"],
                "Estado inicial": "",
                "Error": "No se encontró la fila original del Excel.",
            })
            continue

        validation = validate_row(source_row.iloc[0], config)

        try:
            created = create_issue(validation, config)
            results.append({
                "ID origen": source_id,
                "Resultado": "CREADO",
                "Jira Key": created.get("key", ""),
                "Tipo Jira": config["issue_type"],
                "Estado inicial": "CREADA",
                "Error": "",
            })
        except Exception as exc:
            results.append({
                "ID origen": source_id,
                "Resultado": "ERROR",
                "Jira Key": "",
                "Tipo Jira": config["issue_type"],
                "Estado inicial": "",
                "Error": f"{source_id}: {exc}",
            })

    return pd.DataFrame(results)


# ============================================================
# INTERFAZ
# ============================================================

_, _, _, access_code = get_settings()

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "import_mode" not in st.session_state:
    st.session_state.import_mode = None

if not st.session_state.authenticated:
    st.title("📥 TMOB – Importador Web")
    st.caption("Acceso al sistema de importación TMOB")

    entered = st.text_input("Código de acceso", type="password")

    if st.button("Acceder", type="primary", use_container_width=True):
        if access_code and entered == access_code:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Código de acceso incorrecto.")

    st.stop()


# ============================================================
# MENÚ
# ============================================================

if st.session_state.import_mode is None:
    st.title("📥 TMOB – Importador Web")
    st.caption("Selecciona el tipo de información que quieres cargar.")
    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📋 Incidencias")
        st.write("Crear incidencias desde un Excel.")
        st.write("Tipo Jira: **Incidencia**")

        if st.button(
            "Abrir Incidencias",
            type="primary",
            use_container_width=True,
        ):
            st.session_state.import_mode = "incidencias"
            st.session_state.pop("precheck", None)
            st.session_state.pop("creation_results", None)
            st.rerun()

    with col2:
        st.subheader("📦 Peticiones / WO")
        st.write("Las WO se importan como **Petición** en Jira.")
        st.write("Tipo Jira: **Petición**")

        if st.button(
            "Abrir Peticiones / WO",
            type="primary",
            use_container_width=True,
        ):
            st.session_state.import_mode = "peticiones"
            st.session_state.pop("precheck", None)
            st.session_state.pop("creation_results", None)
            st.rerun()

    st.stop()


# ============================================================
# IMPORTADOR SELECCIONADO
# ============================================================

config = (
    INC_CONFIG
    if st.session_state.import_mode == "incidencias"
    else WO_CONFIG
)

st.title(config["title"])
st.caption(config["caption"])

if st.session_state.import_mode == "peticiones":
    st.info(
        "Las WO se crearán como Jira **Petición (10224)**. "
        "El estado inicial será **CREADA** y el campo Ámbito será "
        "**Pendiente de asignación**."
    )

if st.button("⬅️ Volver al menú"):
    st.session_state.import_mode = None
    st.session_state.pop("precheck", None)
    st.session_state.pop("creation_results", None)
    st.rerun()

st.divider()

uploaded = st.file_uploader(
    "Carga el Excel",
    type=["xlsx"],
    help=(
        "Incidencias: cabecera en la primera fila. "
        "WO: la primera fila contiene 'Wo_Activas' y la segunda fila "
        "contiene las cabeceras."
    ),
    key=f"upload_{st.session_state.import_mode}",
)

if uploaded is None:
    st.info("Carga el fichero XLSX para comenzar.")
    st.stop()

try:
    df_raw = pd.read_excel(
        uploaded,
        sheet_name="Report",
        header=config["header_row"],
    )
    df = normalize_excel(df_raw, config)
except Exception as exc:
    st.error(f"Error leyendo el Excel: {exc}")
    st.stop()

st.success(
    f"Excel cargado: {len(df)} registros con ID origen válido."
)

if len(df) == 0:
    st.warning(
        f"No se encontraron registros válidos con el patrón "
        f"{config['valid_id_regex']}."
    )
    st.stop()

n = st.number_input(
    "Número de registros a procesar en esta prueba",
    min_value=1,
    max_value=len(df),
    value=min(5, len(df)),
    step=1,
    key=f"number_{st.session_state.import_mode}",
)

test_df = df.iloc[: int(n)].copy()

st.subheader("Previsualización")
st.dataframe(
    test_df[config["preview_columns"]],
    use_container_width=True,
)

st.divider()

if st.button(
    "🔎 Comprobar conexión y duplicados (sin crear)",
    type="primary",
    key=f"precheck_button_{st.session_state.import_mode}",
):
    try:
        pre = run_precheck(test_df, config)
        st.session_state["precheck"] = pre
        st.success(f"Conexión Jira OK · proyecto {PROJECT_KEY}")
    except Exception as exc:
        st.error(f"Error en la precomprobación: {exc}")

if "precheck" in st.session_state:
    pre = st.session_state["precheck"]

    st.subheader("Resultado de la precomprobación")
    st.dataframe(pre, use_container_width=True)

    errors = int((pre["Resultado"] == "ERROR").sum())
    mapping_errors = int(
        (pre["Resultado"] == "ERROR DE MAPEO").sum()
    )
    duplicates = int((pre["Resultado"] == "OMITIR").sum())
    creates = int((pre["Resultado"] == "CREAR").sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CREAR", creates)
    c2.metric("OMITIR", duplicates)
    c3.metric("ERRORES MAPEO", mapping_errors)
    c4.metric("ERRORES", errors)

    if errors == 0 and mapping_errors == 0:
        st.success(
            "Precomprobación correcta. Puedes ejecutar la prueba real."
        )
    else:
        st.warning(
            "Hay registros que no deben crearse hasta corregir "
            "el mapeo o el error."
        )

    if st.button(
        "🚀 EJECUTAR PRUEBA REAL",
        type="secondary",
        disabled=(
            errors > 0
            or mapping_errors > 0
            or creates == 0
        ),
        key=f"create_button_{st.session_state.import_mode}",
    ):
        results = run_creation(pre, test_df, config)
        st.session_state["creation_results"] = results

if "creation_results" in st.session_state:
    st.subheader("Resultado de la carga")
    st.dataframe(
        st.session_state["creation_results"],
        use_container_width=True,
    )
