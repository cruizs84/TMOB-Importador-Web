# -*- coding: utf-8 -*-
"""
TMOB – Importador Web de PETICIONES / ÓRDENES DE TRABAJO

Adaptación del importador estable de Incidencias.

FUENTE:
  Hoja Report
  Header real: fila 2 del Excel (header=1)

CAMPOS ORIGEN:
  ID de orden de trabajo+
  ID de petición asociada
  Empresa*+
  Prioridad
  Fecha de envío
  Fecha de última modificación
  Resumen*
  Estado*

DESTINO JIRA:
  Tipo: Tarea por incidencia o petición
  ID WO -> customfield_10402 (ID incidencia origen)
  REQ   -> customfield_10296 (REQ incidencia)
  Empresa -> customfield_10369 (SOLO opciones ya existentes en Jira)
  Fecha envío -> customfield_10472 (Fecha y hora de inicio)
  Ámbito -> customfield_10295 = 10372 (Pendiente de asignación)
  Prioridad -> prioridad Jira

NO se importa:
  - Fecha de última modificación (Jira gestiona Updated)
  - Estado de origen (la creación empieza en CREADA)
  - Grupo externo resolutor (no está en el Excel)
  - Motivo/Detalle de pendiente
  - Resolución

REGLA DE EMPRESAS:
  El programa NO crea, modifica ni amplía el catálogo de Empresas.
  Consulta las opciones existentes en Jira y usa únicamente esas.
  Se mantienen únicamente alias explícitos para diferencias conocidas
  de escritura.
"""

import pandas as pd
import requests
import streamlit as st
from zoneinfo import ZoneInfo


# ============================================================
# CONFIGURACIÓN
# ============================================================

PROJECT_KEY = "TMOB"
ISSUE_TYPE_NAME = "Tarea por incidencia o petición"

FIELD_AMBITO = "customfield_10295"
FIELD_REQ = "customfield_10296"
FIELD_ID_ORIGEN = "customfield_10402"
FIELD_FECHA_INICIO = "customfield_10472"
FIELD_EMPRESA = "customfield_10369"

AMBITO_PENDIENTE_ASIGNACION = "10372"

SOURCE_ID_COLUMN = "ID de orden de trabajo+"
REQ_COLUMN = "ID de petición asociada"
COMPANY_COLUMN = "Empresa*+"
PRIORITY_COLUMN = "Prioridad"
DATE_COLUMN = "Fecha de envío"
SUMMARY_COLUMN = "Resumen*"
STATUS_COLUMN = "Estado*"

REQUIRED_COLUMNS = [
    SOURCE_ID_COLUMN,
    REQ_COLUMN,
    COMPANY_COLUMN,
    PRIORITY_COLUMN,
    DATE_COLUMN,
    "Fecha de última modificación",
    SUMMARY_COLUMN,
    STATUS_COLUMN,
]

TZ = ZoneInfo("Europe/Madrid")


# ============================================================
# EMPRESAS
# ============================================================
# NO hay catálogo de empresas dentro del código.
# El catálogo se obtiene de Jira.
#
# Estos alias SOLO resuelven diferencias de escritura ya conocidas.
# La empresa destino debe existir previamente en Jira.

COMPANY_ALIASES = {
    "Autocares Prat": "Autocares PRAT",
    "MOHN": "Mohn",
    "MOHN (Grup Baixbús)": "Mohn",
}


# ============================================================
# PRIORIDADES
# ============================================================

PRIORITY_ALIASES = {
    "Crítica": "Crítica",
    "Critica": "Crítica",
    "Alta": "Alta",
    "Media": "Media",
    "Baja": "Baja",
}


# ============================================================
# UTILIDADES
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def jira_config():
    jira_url = st.secrets.get("JIRA_URL", "").strip().rstrip("/")
    email = st.secrets.get("JIRA_EMAIL", "").strip()
    token = st.secrets.get("JIRA_API_TOKEN", "").strip()

    if not jira_url or not email or not token:
        raise RuntimeError(
            "Faltan JIRA_URL, JIRA_EMAIL o JIRA_API_TOKEN en secrets."
        )

    return jira_url, email, token


def auth_headers():
    return {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


# ============================================================
# PRE-FLIGHT
# ============================================================

def preflight():
    jira_url, email, token = jira_config()

    if jira_url != "https://suport-secom.atlassian.net":
        raise RuntimeError(
            "JIRA_URL incorrecta. Debe ser "
            "https://suport-secom.atlassian.net"
        )

    r = requests.get(
        f"{jira_url}/rest/api/3/myself",
        auth=(email, token),
        headers={"Accept": "application/json"},
        timeout=30,
    )
    r.raise_for_status()

    r = requests.get(
        f"{jira_url}/rest/api/3/project/{PROJECT_KEY}",
        auth=(email, token),
        headers={"Accept": "application/json"},
        timeout=30,
    )
    r.raise_for_status()

    return True


# ============================================================
# CATÁLOGO EMPRESA DESDE JIRA
# ============================================================

@st.cache_data(ttl=300)
def load_company_catalog():
    jira_url, email, token = jira_config()

    url = (
        f"{jira_url}/rest/api/3/field/"
        f"{FIELD_EMPRESA}/context/option"
    )

    r = requests.get(
        url,
        auth=(email, token),
        headers={"Accept": "application/json"},
        params={"maxResults": 1000},
        timeout=30,
    )

    if r.status_code >= 400:
        raise RuntimeError(
            "No se pudo leer el catálogo Empresa de Jira "
            f"({r.status_code}): {r.text[:500]}"
        )

    values = r.json().get("values", [])

    catalog = {}

    for option in values:
        value = clean_text(option.get("value"))
        option_id = clean_text(option.get("id"))

        if value and option_id:
            catalog[value] = option_id

    if not catalog:
        raise RuntimeError(
            "Jira no devolvió opciones para el campo Empresa."
        )

    return catalog


def resolve_company(company_raw, catalog):
    value = clean_text(company_raw)

    # 1. Coincidencia exacta.
    if value in catalog:
        return value, catalog[value]

    # 2. Alias explícito.
    alias = COMPANY_ALIASES.get(value)

    if alias and alias in catalog:
        return alias, catalog[alias]

    # 3. BLOQUEAR: nunca crear empresa.
    return None, None


# ============================================================
# PRIORIDADES
# ============================================================

def resolve_priority(value):
    raw = clean_text(value)

    if raw in PRIORITY_ALIASES:
        return PRIORITY_ALIASES[raw]

    for source, target in PRIORITY_ALIASES.items():
        if raw.casefold() == source.casefold():
            return target

    return None


def get_priority_id(priority_name):
    jira_url, email, token = jira_config()

    r = requests.get(
        f"{jira_url}/rest/api/3/priority",
        auth=(email, token),
        headers={"Accept": "application/json"},
        timeout=30,
    )

    r.raise_for_status()

    for priority in r.json():
        if clean_text(priority.get("name")) == priority_name:
            return clean_text(priority.get("id"))

    return None


# ============================================================
# FECHAS
# ============================================================

def parse_date_to_iso(value):
    if pd.isna(value) or clean_text(value) == "":
        return None

    ts = pd.to_datetime(
        value,
        dayfirst=True,
        errors="coerce",
    )

    if pd.isna(ts):
        return None

    if ts.tzinfo is None:
        ts = ts.tz_localize(TZ)
    else:
        ts = ts.tz_convert(TZ)

    return ts.isoformat()


# ============================================================
# LECTURA DEL EXCEL
# ============================================================

def read_workbook(uploaded_file):
    # El fichero tiene una fila de título "Wo_Activas"
    # y la fila siguiente contiene las cabeceras reales.
    df = pd.read_excel(
        uploaded_file,
        sheet_name="Report",
        header=1,
    )

    df.columns = [str(c).strip() for c in df.columns]

    missing = [
        c for c in REQUIRED_COLUMNS
        if c not in df.columns
    ]

    if missing:
        raise RuntimeError(
            "Faltan columnas obligatorias en el Excel: "
            + ", ".join(missing)
        )

    df[SOURCE_ID_COLUMN] = df[SOURCE_ID_COLUMN].map(clean_text)

    # Solo WO válidas.
    df = df[
        df[SOURCE_ID_COLUMN].str.match(
            r"^WO\d+$",
            na=False,
        )
    ].copy()

    return df


# ============================================================
# DUPLICADOS
# ============================================================

def find_existing_ids(ids):
    if not ids:
        return set()

    jira_url, email, token = jira_config()
    existing = set()

    for start in range(0, len(ids), 100):

        batch = ids[start:start + 100]

        quoted = ",".join(
            f'"{x}"'
            for x in batch
        )

        # customfield_10402 -> cf[10402]
        jql = (
            f'project = {PROJECT_KEY} '
            f'AND cf[10402] in ({quoted})'
        )

        r = requests.post(
            f"{jira_url}/rest/api/3/search/jql",
            auth=(email, token),
            headers=auth_headers(),
            json={
                "jql": jql,
                "maxResults": 100,
                "fields": [FIELD_ID_ORIGEN],
            },
            timeout=30,
        )

        # Fallback para API search antigua.
        if r.status_code == 404:
            r = requests.post(
                f"{jira_url}/rest/api/3/search",
                auth=(email, token),
                headers=auth_headers(),
                json={
                    "jql": jql,
                    "maxResults": 100,
                    "fields": [FIELD_ID_ORIGEN],
                },
                timeout=30,
            )

        r.raise_for_status()

        for issue in r.json().get("issues", []):

            value = clean_text(
                issue.get("fields", {}).get(
                    FIELD_ID_ORIGEN
                )
            )

            if value:
                existing.add(value)

    return existing


# ============================================================
# VALIDACIÓN
# ============================================================

def validate_row(
    row,
    company_catalog,
    priority_ids,
    existing_ids,
):
    errors = []

    source_id = clean_text(
        row[SOURCE_ID_COLUMN]
    )

    req = clean_text(
        row[REQ_COLUMN]
    )

    company_raw = clean_text(
        row[COMPANY_COLUMN]
    )

    priority_raw = clean_text(
        row[PRIORITY_COLUMN]
    )

    summary = clean_text(
        row[SUMMARY_COLUMN]
    )

    start_iso = parse_date_to_iso(
        row[DATE_COLUMN]
    )

    if not source_id:
        errors.append("WO vacía")

    if not summary:
        errors.append("Resumen vacío")

    if not req:
        errors.append("REQ vacío")

    company_name, company_id = resolve_company(
        company_raw,
        company_catalog,
    )

    if not company_id:
        errors.append(
            f"Empresa no dada de alta en Jira: "
            f"'{company_raw}'"
        )

    priority_name = resolve_priority(
        priority_raw
    )

    if not priority_name:
        errors.append(
            f"Prioridad no reconocida: "
            f"'{priority_raw}'"
        )

    priority_id = (
        priority_ids.get(priority_name)
        if priority_name
        else None
    )

    if priority_name and not priority_id:
        errors.append(
            f"No se encontró ID Jira para prioridad: "
            f"'{priority_name}'"
        )

    if not start_iso:
        errors.append(
            "Fecha de envío inválida o vacía"
        )

    if source_id in existing_ids:
        errors.append(
            "Duplicado: la WO ya existe en Jira"
        )

    return {
        "source_id": source_id,
        "req": req,
        "company_raw": company_raw,
        "company_name": company_name,
        "company_id": company_id,
        "priority_raw": priority_raw,
        "priority_name": priority_name,
        "priority_id": priority_id,
        "summary": summary,
        "start_iso": start_iso,
        "source_status": clean_text(
            row[STATUS_COLUMN]
        ),
        "errors": errors,
    }


# ============================================================
# CREACIÓN EN JIRA
# ============================================================

def create_issue(v):
    jira_url, email, token = jira_config()

    fields = {
        "project": {
            "key": PROJECT_KEY
        },

        "issuetype": {
            "name": ISSUE_TYPE_NAME
        },

        "summary": v["summary"],

        "priority": {
            "id": v["priority_id"]
        },

        # Ámbito = Pendiente de asignación
        FIELD_AMBITO: {
            "id": AMBITO_PENDIENTE_ASIGNACION
        },

        # REQ asociada
        FIELD_REQ: v["req"],

        # WO original
        FIELD_ID_ORIGEN: v["source_id"],

        # Fecha de envío
        FIELD_FECHA_INICIO: v["start_iso"],

        # Empresa EXISTENTE en Jira
        FIELD_EMPRESA: {
            "id": v["company_id"]
        },
    }

    r = requests.post(
        f"{jira_url}/rest/api/3/issue",
        auth=(email, token),
        headers=auth_headers(),
        json={"fields": fields},
        timeout=30,
    )

    if r.status_code not in (200, 201):
        return (
            False,
            None,
            r.text[:1000],
        )

    return (
        True,
        r.json().get("key"),
        None,
    )


# ============================================================
# INTERFAZ
# ============================================================

st.set_page_config(
    page_title="TMOB – Importador Peticiones",
    page_icon="📥",
    layout="wide",
)

st.title(
    "📥 TMOB – Importador Web de Peticiones"
)

st.caption(
    "Importación controlada de órdenes de trabajo (WO) "
    "como Tarea por incidencia o petición"
)

access_code = st.secrets.get(
    "TMOB_ACCESS_CODE",
    "",
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False


if not st.session_state.authenticated:

    code = st.text_input(
        "Código de acceso",
        type="password",
    )

    if st.button("Acceder"):

        if code == access_code:
            st.session_state.authenticated = True
            st.rerun()

        else:
            st.error(
                "Código de acceso incorrecto."
            )

    st.stop()


with st.sidebar:

    st.header("Configuración")

    st.write(
        f"Proyecto: **{PROJECT_KEY}**"
    )

    st.write(
        f"Tipo Jira: **{ISSUE_TYPE_NAME}**"
    )

    st.write(
        "Estado inicial: **CREADA**"
    )

    st.write(
        "Ámbito inicial: "
        "**Pendiente de asignación**"
    )

    st.info(
        "Las empresas NO se crean desde el importador. "
        "Se utilizan exclusivamente las empresas "
        "ya dadas de alta en Jira."
    )


if st.button(
    "🔎 Comprobar conexión Jira"
):

    try:
        preflight()

        st.success(
            "Conexión Jira OK."
        )

    except Exception as e:

        st.error(
            f"Error de conexión: {e}"
        )


uploaded = st.file_uploader(
    "Selecciona el Excel de peticiones / WO",
    type=["xlsx", "xls"],
)


if uploaded:

    try:

        df = read_workbook(
            uploaded
        )

        st.success(
            f"Excel cargado: {len(df)} "
            "órdenes de trabajo con WO válida."
        )

        st.subheader(
            "Previsualización"
        )

        preview_cols = [
            SOURCE_ID_COLUMN,
            REQ_COLUMN,
            COMPANY_COLUMN,
            PRIORITY_COLUMN,
            SUMMARY_COLUMN,
            STATUS_COLUMN,
        ]

        st.dataframe(
            df[preview_cols].head(20),
            use_container_width=True,
        )

        st.subheader(
            "Configuración de prueba"
        )

        test_count = st.number_input(
            "Número de WO a procesar",
            min_value=1,
            max_value=len(df),
            value=min(5, len(df)),
            step=1,
        )

        if st.button(
            "🧪 Validar selección"
        ):

            with st.spinner(
                "Consultando catálogo Empresa, "
                "prioridades y duplicados..."
            ):

                catalog = (
                    load_company_catalog()
                )

                priority_ids = {}

                for priority in set(
                    PRIORITY_ALIASES.values()
                ):
                    priority_ids[priority] = (
                        get_priority_id(
                            priority
                        )
                    )

                selected = df.head(
                    int(test_count)
                ).copy()

                ids = selected[
                    SOURCE_ID_COLUMN
                ].tolist()

                existing = (
                    find_existing_ids(ids)
                )

                results = [
                    validate_row(
                        row,
                        catalog,
                        priority_ids,
                        existing,
                    )
                    for _, row in selected.iterrows()
                ]

                validation_df = pd.DataFrame(
                    [
                        {
                            "WO": r["source_id"],
                            "REQ": r["req"],
                            "Empresa origen": (
                                r["company_raw"]
                            ),
                            "Empresa Jira": (
                                r["company_name"]
                                or ""
                            ),
                            "Prioridad": (
                                r["priority_name"]
                                or ""
                            ),
                            "Resumen": (
                                r["summary"]
                            ),
                            "Estado origen": (
                                r["source_status"]
                            ),
                            "Resultado": (
                                "OK"
                                if not r["errors"]
                                else " | ".join(
                                    r["errors"]
                                )
                            ),
                        }
                        for r in results
                    ]
                )

                st.session_state.validation_results = (
                    results
                )

                st.session_state.validation_df = (
                    validation_df
                )

        if (
            "validation_df"
            in st.session_state
        ):

            st.subheader(
                "Resultado de validación"
            )

            st.dataframe(
                st.session_state.validation_df,
                use_container_width=True,
            )

            results = (
                st.session_state.validation_results
            )

            valid_count = sum(
                not r["errors"]
                for r in results
            )

            if valid_count == len(results):

                st.success(
                    f"Validación OK: "
                    f"{valid_count}/"
                    f"{len(results)} WO "
                    "listas para crear."
                )

                if st.button(
                    "🚀 Crear WO en Jira"
                ):

                    created = 0
                    failed = 0

                    progress = st.progress(
                        0
                    )

                    status_box = st.empty()

                    for i, result in enumerate(
                        results,
                        start=1,
                    ):

                        ok, key, error = (
                            create_issue(
                                result
                            )
                        )

                        if ok:

                            created += 1

                            status_box.success(
                                f"{result['source_id']} "
                                f"→ {key}"
                            )

                        else:

                            failed += 1

                            status_box.error(
                                f"{result['source_id']} "
                                f"→ ERROR: {error}"
                            )

                        progress.progress(
                            i / len(results)
                        )

                    st.info(
                        "Proceso terminado. "
                        f"Creadas: {created}. "
                        f"Errores: {failed}."
                    )

            else:

                st.warning(
                    f"Hay "
                    f"{len(results) - valid_count} "
                    "WO con errores. "
                    "La creación está bloqueada "
                    "hasta corregirlos."
                )

    except Exception as e:

        st.error(
            f"Error procesando el Excel: {e}"
        )


st.divider()

st.caption(
    "TMOB – Importador de Peticiones | "
    "No modifica el catálogo de Empresas de Jira."
)
