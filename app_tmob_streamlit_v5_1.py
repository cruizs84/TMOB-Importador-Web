import os
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth

# ============================================================
# TMOB - IMPORTADOR WEB
# Streamlit
# ============================================================

st.set_page_config(
    page_title="TMOB - Importador",
    page_icon="📥",
    layout="wide",
)

PROJECT_KEY = "TMOB"
PROJECT_ID = "10166"
ISSUE_TYPE = "Incidencia"
ISSUE_TYPE_ID = "10223"

FIELD_ID_ORIGEN = "customfield_10402"
FIELD_REQ = "customfield_10296"
FIELD_GRUPO = "customfield_10330"
FIELD_MOTIVO = "customfield_10332"
FIELD_AMBITO = "customfield_10295"
FIELD_EMPRESA = "customfield_10369"
FIELD_FECHA_INICIO = "customfield_10472"

AMBITO_PROVISIONAL = "Pendiente de asignación"

PRIORIDADES = ["Crítica", "Alta", "Media", "Baja"]

# ============================================================
# MAPEO EXPLICITO DE LAS 32 EMPRESAS DEL EXCEL -> JIRA
#
# No se deja el mapeo a una coincidencia automática.
# Cada valor de Empresa origen tiene aquí un destino Jira.
#
# Las equivalencias documentadas son:
#   Autocares Prat -> Autocares PRAT
#   variantes MOHN -> Mohn
#
# El resto conserva literalmente el nombre del Excel.
# Antes de crear incidencias, la aplicación comprueba que el
# destino exista realmente entre las opciones del campo Empresa.
# ============================================================

EMPRESA_MAP = {
    "FGC": "FGC",
    "SAGALES": "SAGALES",
    "RENFE": "RENFE",
    "TUSGSAL": "TUSGSAL",
    "LAC": "LAC",
    "SOC Mobilitat": "SOC Mobilitat",
    "TRAM Baix": "TRAM Baix",
    "SOLER I SAURET": "SOLER I SAURET",
    "TRAM": "TRAM",
    "TUS, S. Coop. CL (Sabadell)": "TUS, S. Coop. CL (Sabadell)",
    "Sarbus (Marfina Bus, SA)": "Sarbus (Marfina Bus, SA)",
    "TMB Sistemes Distribuïts": "TMB Sistemes Distribuïts",
    "AVANZA": "AVANZA",
    "Fujitsu": "Fujitsu",
    "UTE Monbus Port": "UTE Monbus Port",
    "INDRA": "INDRA",
    "Transports Generals d'Olesa": "Transports Generals d'Olesa",
    "Autocorb": "Autocorb",
    "Alsina Graells de Auto Trans": "Alsina Graells de Auto Trans",
    "AMB": "AMB",
    "Cintoi Bus, SL": "Cintoi Bus, SL",
    "TMESA": "TMESA",
    "TRAM Besòs": "TRAM Besòs",
    "Autocars Julià": "Autocars Julià",
    "EMPRESA PLANA, S.L.": "EMPRESA PLANA, S.L.",
    "Empresa Casas, SA": "Empresa Casas, SA",
    "MASATS": "MASATS",
    "Autos Castellbisbal": "Autos Castellbisbal",
    "FONT": "FONT",
    "Smarting": "Smarting",
    "TMB Metro": "TMB Metro",
    "Autocares Prat": "Autocares PRAT",
    "MOHN (Grup Baixbús)": "Mohn",
    "MOHN (Grup Baixbus)": "Mohn",
    "Mohn (Grup Baixbus)": "Mohn",
    "Mohn (Grup Baixbús)": "Mohn",
}

# Solo las 32 empresas que aparecen como categorías históricas.
EMPRESAS_HISTORICAS_32 = [
    "FGC",
    "SAGALES",
    "RENFE",
    "TUSGSAL",
    "LAC",
    "SOC Mobilitat",
    "TRAM Baix",
    "SOLER I SAURET",
    "TRAM",
    "TUS, S. Coop. CL (Sabadell)",
    "Sarbus (Marfina Bus, SA)",
    "TMB Sistemes Distribuïts",
    "AVANZA",
    "Fujitsu",
    "UTE Monbus Port",
    "INDRA",
    "Transports Generals d'Olesa",
    "Autocorb",
    "Alsina Graells de Auto Trans",
    "AMB",
    "Cintoi Bus, SL",
    "TMESA",
    "TRAM Besòs",
    "Autocars Julià",
    "EMPRESA PLANA, S.L.",
    "Empresa Casas, SA",
    "MASATS",
    "Autos Castellbisbal",
    "FONT",
    "Smarting",
    "TMB Metro",
    "Autocares Prat",
]

# Mantener compatibilidad con el resto del código.
EMPRESA_ALIASES = EMPRESA_MAP


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
    "Sistemes": "Sistemes",
}

MOTIVO_MAP = {
    "Incidencia original pendiente": "Incidencia principal pendiente",
}


def normalizar(v):
    if v is None or pd.isna(v):
        return ""
    s = str(v).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.split()).casefold()


def text(v):
    if v is None or pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def secrets_config():
    # La app usa JIRA_URL, que coincide con el nombre configurado
    # en Streamlit Secrets.
    url = str(st.secrets.get("JIRA_URL", "")).strip().rstrip("/")
    email = str(st.secrets.get("JIRA_EMAIL", "")).strip()
    token = str(st.secrets.get("JIRA_API_TOKEN", "")).strip()
    access_code = str(st.secrets.get("TMOB_ACCESS_CODE", "")).strip()

    if not url or not email or not token:
        return None, None, None, access_code, (
            "Faltan JIRA_URL, JIRA_EMAIL o JIRA_API_TOKEN en Secrets."
        )

    if "support-secom.atlassian.net" in url.lower():
        return None, None, None, access_code, (
            "JIRA_URL es incorrecta. Debe ser "
            "https://suport-secom.atlassian.net"
        )

    return url, email, token, access_code, None


def jira_session():
    url, email, token, _, error = secrets_config()
    if error:
        return None, error

    s = requests.Session()
    s.auth = HTTPBasicAuth(email, token)
    s.headers.update({
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "TMOB-Importador-Web/5.1",
    })
    return (url, s), None


def jira_get(url, session, params=None, timeout=(4, 10)):
    try:
        return session.get(url, params=params, timeout=timeout)
    except requests.exceptions.Timeout:
        raise RuntimeError("Jira ha superado el tiempo máximo de espera.")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Error de conexión con Jira: {e}")


def jira_post(url, session, payload, timeout=(4, 15)):
    try:
        return session.post(url, json=payload, timeout=timeout)
    except requests.exceptions.Timeout:
        raise RuntimeError("Jira ha superado el tiempo máximo de espera.")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Error de conexión con Jira: {e}")


def cargar_excel(uploaded):
    df = pd.read_excel(uploaded, sheet_name="Report", header=0)
    df.columns = [str(c).strip() for c in df.columns]

    required = [
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

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Faltan columnas: " + ", ".join(missing)
        )

    df = df.rename(columns={
        "ID de la incidencia*+": "ID incidencia origen",
        "ID de petición de servicio": "REQ incidencia",
        "Empresa*+": "Empresa origen",
        "Prioridad*": "Prioridad",
        "Fecha de envío": "Fecha y hora de inicio",
        "Fecha de última modificación": "Fecha última modificación origen",
        "Resumen*": "Resumen",
        "Grupo asignado*+": "Grupo externo resolutor origen",
        "Estado*": "Estado origen",
        "Status_Reason_Hidden": "Motivo de Pendiente origen",
    })

    df["ID incidencia origen"] = (
        df["ID incidencia origen"].astype("string").str.strip()
    )
    df = df[
        df["ID incidencia origen"].fillna("").str.match(
            r"^INC\d+$", na=False
        )
    ].copy()

    return df


def preflight():
    result, error = jira_session()
    if error:
        return False, error

    base, s = result

    r = jira_get(f"{base}/rest/api/3/myself", s)
    if not r.ok:
        return False, f"/myself -> HTTP {r.status_code}: {r.text[:500]}"

    r = jira_get(f"{base}/rest/api/3/project/{PROJECT_KEY}", s)
    if not r.ok:
        return False, (
            f"/project/{PROJECT_KEY} -> HTTP {r.status_code}: "
            f"{r.text[:500]}"
        )

    project = r.json()
    if str(project.get("id")) != PROJECT_ID:
        return False, (
            f"TMOB tiene ID {project.get('id')} y se esperaba {PROJECT_ID}."
        )

    r = jira_get(
        f"{base}/rest/api/3/issuetype/{ISSUE_TYPE_ID}", s
    )
    if not r.ok:
        return False, (
            f"/issuetype/{ISSUE_TYPE_ID} -> HTTP {r.status_code}: "
            f"{r.text[:500]}"
        )

    issue_type = r.json()
    if issue_type.get("name") != ISSUE_TYPE:
        return False, (
            f"El tipo {ISSUE_TYPE_ID} es "
            f"{issue_type.get('name')}, no {ISSUE_TYPE}."
        )

    return True, {
        "base": base,
        "session": s,
        "usuario": (
            jira_get(f"{base}/rest/api/3/myself", s).json()
            .get("displayName", "usuario Jira")
        ),
    }


def obtener_opciones_campo(field_id, base, s):
    """Obtiene opciones de un campo select sin usar createmeta primero."""
    url = f"{base}/rest/api/3/field/{field_id}/context"
    r = jira_get(
        url, s, params={"startAt": 0, "maxResults": 100}
    )

    if not r.ok:
        return {}

    contexts = r.json().get("values", [])
    catalog = {}

    for ctx in contexts:
        cid = ctx.get("id")
        if not cid:
            continue

        start = 0
        while True:
            rr = jira_get(
                f"{base}/rest/api/3/field/{field_id}/context/{cid}/option",
                s,
                params={
                    "startAt": start,
                    "maxResults": 100,
                    "onlyOptions": "true",
                },
            )

            if not rr.ok:
                break

            data = rr.json()
            vals = data.get("values", [])

            for item in vals:
                name = item.get("value") or item.get("name")
                if name and not item.get("disabled", False):
                    catalog[str(name).strip()] = item

            if data.get("isLast", True) or not vals:
                break
            start += len(vals)

    return catalog


def obtener_catalogos(info):
    base, s = info["base"], info["session"]

    empresa = obtener_opciones_campo(FIELD_EMPRESA, base, s)
    grupo = obtener_opciones_campo(FIELD_GRUPO, base, s)
    motivo = obtener_opciones_campo(FIELD_MOTIVO, base, s)

    return empresa, grupo, motivo


def resolver_empresa(v, catalog):
    src = text(v)

    if not src:
        return None, "VACIO"

    # 1. El origen debe estar definido en el mapa.
    if src not in EMPRESA_MAP:
        return None, "NO_EXISTE_EN_MAPEO_EXPLICITO"

    target = EMPRESA_MAP[src]

    # 2. El destino debe existir realmente en Jira.
    if target in catalog:
        if target == src:
            return target, "MAPEO_EXPLICITO_1:1"
        return target, f"MAPEO_EXPLICITO:{src} -> {target}"

    # 3. Permitimos coincidencia normalizada SOLO para localizar
    #    el mismo destino, nunca para inventar una equivalencia.
    matches = [
        name for name in catalog
        if normalizar(name) == normalizar(target)
    ]

    if len(matches) == 1:
        return matches[0], f"MAPEO_EXPLICITO_NORMALIZADO:{src} -> {matches[0]}"

    if len(matches) > 1:
        return None, (
            "DESTINO_AMBIGUO_EN_JIRA:"
            + " | ".join(matches)
        )

    return None, f"DESTINO_NO_EXISTE_EN_JIRA:{target}"


def resolver_grupo(v, catalog):
    src = text(v)
    if not src:
        return None, "VACIO"

    target = GROUP_MAP.get(src)
    if not target:
        return None, "NO_ENCONTRADO_EN_MAPA"

    if catalog and target not in catalog:
        # No se traduce silenciosamente a otro valor.
        return None, f"NO_EXISTE_EN_JIRA:{target}"

    return target, "EXPLICITO"


def resolver_motivo(v, catalog):
    src = text(v)
    if not src:
        return None, "VACIO"

    target = MOTIVO_MAP.get(src, src)

    if catalog and target not in catalog:
        return None, f"NO_EXISTE_EN_JIRA:{target}"

    estado = (
        "EQUIVALENCIA_APROBADA"
        if src in MOTIVO_MAP else "LITERAL"
    )
    return target, estado


def preparar(df, empresa_cat, grupo_cat, motivo_cat):
    out = df.copy()

    e = out["Empresa origen"].apply(
        lambda x: resolver_empresa(x, empresa_cat)
    )
    out["Empresa"] = e.apply(lambda x: x[0])
    out["Empresa mapeo"] = e.apply(lambda x: x[1])

    g = out["Grupo externo resolutor origen"].apply(
        lambda x: resolver_grupo(x, grupo_cat)
    )
    out["Grupo externo resolutor"] = g.apply(lambda x: x[0])
    out["Grupo mapeo"] = g.apply(lambda x: x[1])

    m = out["Motivo de Pendiente origen"].apply(
        lambda x: resolver_motivo(x, motivo_cat)
    )
    out["Motivo de Pendiente"] = m.apply(lambda x: x[0])
    out["Motivo mapeo"] = m.apply(lambda x: x[1])

    out["Estado Jira"] = (
        out["Estado origen"].astype("string").str.strip().map({
            "Asignado": "ASIGNADA",
            "En curso": "EN PROGRESO",
            "Pendiente": "PENDIENTE",
        })
    )

    out["Fecha y hora de inicio"] = pd.to_datetime(
        out["Fecha y hora de inicio"],
        errors="coerce",
        dayfirst=True,
    )

    errores = []

    for _, row in out.iterrows():
        ident = row["ID incidencia origen"]
        msgs = []

        if not row["Empresa"]:
            msgs.append(
                f"Empresa '{row['Empresa origen']}' "
                f"-> {row['Empresa mapeo']}"
            )

        if not row["Grupo externo resolutor"]:
            msgs.append(
                f"Grupo '{row['Grupo externo resolutor origen']}' "
                f"-> {row['Grupo mapeo']}"
            )

        if not row["Estado Jira"]:
            msgs.append(
                f"Estado origen no contemplado: '{row['Estado origen']}'"
            )

        if not text(row["Resumen"]):
            msgs.append("Resumen vacío")

        if text(row["Prioridad"]) not in PRIORIDADES:
            msgs.append(
                f"Prioridad no válida: '{row['Prioridad']}'"
            )

        if pd.isna(row["Fecha y hora de inicio"]):
            msgs.append("Fecha de envío no válida")

        if (
            text(row["Estado origen"]) == "Pendiente"
            and not text(row["Motivo de Pendiente origen"])
        ):
            msgs.append("Pendiente sin Status_Reason_Hidden")

        if msgs:
            errores.append((ident, " | ".join(msgs)))

    return out, errores


def buscar_duplicados(ids, info):
    if not ids:
        return {}

    base, s = info["base"], info["session"]
    values = ", ".join(
        '"' + str(x).replace('"', '\\"') + '"'
        for x in ids
    )
    jql = (
        f'project = {PROJECT_KEY} '
        f'AND "{FIELD_ID_ORIGEN}" in ({values})'
    )

    r = jira_post(
        f"{base}/rest/api/3/search/jql",
        s,
        {
            "jql": jql,
            "maxResults": len(ids),
            "fields": ["summary", FIELD_ID_ORIGEN],
        },
    )

    if not r.ok:
        raise RuntimeError(
            f"Búsqueda de duplicados -> HTTP {r.status_code}: "
            f"{r.text[:1000]}"
        )

    result = {}
    for issue in r.json().get("issues", []):
        origin = issue.get("fields", {}).get(FIELD_ID_ORIGEN)
        if origin:
            result[str(origin)] = issue.get("key")
    return result


def iso_date(v):
    if pd.isna(v):
        return None

    if isinstance(v, pd.Timestamp):
        dt = v.to_pydatetime()
    elif isinstance(v, datetime):
        dt = v
    else:
        dt = pd.to_datetime(
            v, errors="coerce", dayfirst=True
        ).to_pydatetime()

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo("Europe/Madrid"))
    else:
        dt = dt.astimezone(ZoneInfo("Europe/Madrid"))

    return dt.isoformat(timespec="seconds")


def create_payload(row):
    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"id": ISSUE_TYPE_ID},
        "summary": text(row["Resumen"]),
        FIELD_ID_ORIGEN: text(row["ID incidencia origen"]),
        FIELD_AMBITO: {"value": AMBITO_PROVISIONAL},
    }

    req = text(row["REQ incidencia"])
    if req:
        fields[FIELD_REQ] = req

    empresa = text(row["Empresa"])
    if empresa:
        fields[FIELD_EMPRESA] = {"value": empresa}

    prioridad = text(row["Prioridad"])
    if prioridad:
        fields["priority"] = {"name": prioridad}

    fecha = iso_date(row["Fecha y hora de inicio"])
    if fecha:
        fields[FIELD_FECHA_INICIO] = fecha

    grupo = text(row["Grupo externo resolutor"])
    if grupo:
        # Grupo externo resolutor es un catálogo de opciones.
        fields[FIELD_GRUPO] = {"value": grupo}

    if text(row["Estado origen"]) == "Pendiente":
        motivo = text(row["Motivo de Pendiente"])
        if motivo:
            fields[FIELD_MOTIVO] = {"value": motivo}

    return fields


def crear(row, info):
    base, s = info["base"], info["session"]
    r = jira_post(
        f"{base}/rest/api/3/issue",
        s,
        {"fields": create_payload(row)},
        timeout=(4, 20),
    )

    if not r.ok:
        return None, (
            f"HTTP {r.status_code}: {r.text[:1500]}"
        )

    return r.json().get("key"), None


# ============================================================
# INTERFAZ
# ============================================================

st.title("📥 TMOB — Importador de incidencias")
st.caption("Importación controlada desde Excel a Jira Cloud")

url, email, token, access_code, config_error = secrets_config()

if config_error:
    st.error(config_error)
    st.info(
        "En Streamlit → Settings → Secrets deben existir "
        "JIRA_URL, JIRA_EMAIL, JIRA_API_TOKEN y TMOB_ACCESS_CODE."
    )
    st.stop()

# Acceso común
if "tmob_auth" not in st.session_state:
    st.session_state.tmob_auth = False

if not st.session_state.tmob_auth:
    st.subheader("Acceso")
    code = st.text_input("Código de acceso", type="password")
    if st.button("Entrar", type="primary"):
        if code == access_code:
            st.session_state.tmob_auth = True
            st.rerun()
        else:
            st.error("Código incorrecto.")
    st.stop()

st.success("Acceso autorizado")

with st.sidebar:
    st.subheader("Configuración")
    st.code(url, language=None)
    st.write("Proyecto: TMOB")
    st.write("Tipo: Incidencia")
    st.write("Ámbito de importación: Pendiente de asignación")

st.header("1. Cargar Excel")
uploaded = st.file_uploader(
    "Selecciona INC activas (13).xlsx",
    type=["xlsx"],
)

if not uploaded:
    st.info("Carga el Excel definitivo para continuar.")
    st.stop()

try:
    df = cargar_excel(uploaded)
except Exception as e:
    st.error(f"Error al leer el Excel: {e}")
    st.stop()

st.success(
    f"Excel cargado: **{len(df)} incidencias con ID origen válido**."
)

st.dataframe(
    df[
        [
            "ID incidencia origen",
            "REQ incidencia",
            "Empresa origen",
            "Prioridad",
            "Resumen",
            "Grupo externo resolutor origen",
            "Estado origen",
            "Motivo de Pendiente origen",
        ]
    ].head(10),
    use_container_width=True,
)

st.header("2. Comprobar conexión")

if st.button("🔎 Comprobar conexión y catálogos", type="primary"):
    with st.status("Comprobando Jira...", expanded=True) as status:
        ok, result = preflight()

        if not ok:
            status.update(
                label="Error de conexión",
                state="error",
            )
            st.error(result)
            st.stop()

        st.write(f"Usuario Jira: {result['usuario']}")
        st.write("Proyecto TMOB: OK")
        st.write("Tipo Incidencia: OK")

        try:
            empresa_cat, grupo_cat, motivo_cat = obtener_catalogos(result)

            st.session_state.jira_info = result
            st.session_state.empresa_cat = empresa_cat
            st.session_state.grupo_cat = grupo_cat
            st.session_state.motivo_cat = motivo_cat

            st.write(f"Empresa: {len(empresa_cat)} opciones")
            st.write(f"Grupo externo resolutor: {len(grupo_cat)} opciones")
            st.write(f"Motivo pendiente: {len(motivo_cat)} opciones")

            status.update(
                label="Conexión y catálogos OK",
                state="complete",
            )
        except Exception as e:
            status.update(
                label="Error obteniendo catálogos",
                state="error",
            )
            st.error(str(e))

if "jira_info" not in st.session_state:
    st.warning(
        "Primero pulsa «Comprobar conexión y catálogos»."
    )
    st.stop()

st.header("3. Validación del mapeo")

prepared, errores = preparar(
    df,
    st.session_state.empresa_cat,
    st.session_state.grupo_cat,
    st.session_state.motivo_cat,
)

if errores:
    st.error(
        f"Hay {len(errores)} errores. La importación queda BLOQUEADA."
    )

    err_df = pd.DataFrame(
        errores,
        columns=["ID incidencia origen", "Error"],
    )
    st.dataframe(err_df, use_container_width=True)

    st.subheader("Empresas")
    st.dataframe(
        prepared[
            [
                "Empresa origen",
                "Empresa",
                "Empresa mapeo",
            ]
        ].drop_duplicates(),
        use_container_width=True,
    )

    st.subheader("Grupos")
    st.dataframe(
        prepared[
            [
                "Grupo externo resolutor origen",
                "Grupo externo resolutor",
                "Grupo mapeo",
            ]
        ].drop_duplicates(),
        use_container_width=True,
    )

    st.subheader("Motivos")
    st.dataframe(
        prepared[
            [
                "Motivo de Pendiente origen",
                "Motivo de Pendiente",
                "Motivo mapeo",
            ]
        ].drop_duplicates(),
        use_container_width=True,
    )

    st.stop()

st.success("Validación completa: 0 errores de mapeo/datos.")

# Control explícito: las empresas históricas deben estar presentes
# en el diccionario de mapeo.
empresas_detectadas = set(
    prepared["Empresa origen"].dropna().astype(str).str.strip()
)
empresas_sin_mapa = sorted(
    empresas_detectadas - set(EMPRESA_MAP.keys())
)

if empresas_sin_mapa:
    st.error(
        "Hay empresas del Excel que no tienen mapeo explícito: "
        + ", ".join(empresas_sin_mapa)
    )
    st.stop()

st.subheader("Mapa explícito de Empresas (Excel → Jira)")

mapa_visual = pd.DataFrame(
    [
        {
            "Empresa origen": k,
            "Empresa Jira": EMPRESA_MAP[k],
            "Tipo": (
                "1:1"
                if k == EMPRESA_MAP[k]
                else "EQUIVALENCIA"
            ),
        }
        for k in EMPRESAS_HISTORICAS_32
    ]
)

st.dataframe(
    mapa_visual,
    use_container_width=True,
    hide_index=True,
)

st.subheader("Empresas detectadas")
st.dataframe(
    prepared[
        [
            "Empresa origen",
            "Empresa",
            "Empresa mapeo",
        ]
    ].drop_duplicates().sort_values("Empresa origen"),
    use_container_width=True,
)

st.header("4. Prueba controlada")

limit = st.number_input(
    "Número de incidencias a procesar",
    min_value=1,
    max_value=len(prepared),
    value=min(5, len(prepared)),
    step=1,
)

test_df = prepared.head(int(limit)).copy()

st.write(
    f"Se han seleccionado **{len(test_df)}** incidencias "
    "empezando por la primera del Excel."
)

if st.button("🔍 Comprobar duplicados (sin crear)"):
    try:
        ids = test_df["ID incidencia origen"].tolist()
        duplicates = buscar_duplicados(ids, st.session_state.jira_info)

        st.session_state.duplicates = duplicates

        if duplicates:
            st.warning("Se han encontrado incidencias ya existentes.")
            for origin, key in duplicates.items():
                st.write(f"- {origin} → {key}")
        else:
            st.success(
                f"No se encontraron duplicados entre las {len(ids)} seleccionadas."
            )
    except Exception as e:
        st.error(str(e))

if "duplicates" in st.session_state:
    st.subheader("Resumen de la prueba")

    for _, row in test_df.iterrows():
        origin = row["ID incidencia origen"]
        if origin in st.session_state.duplicates:
            st.write(
                f"⏭️ {origin} → ya existe como "
                f"{st.session_state.duplicates[origin]}"
            )
        else:
            st.write(
                f"🆕 {origin} | {row['Prioridad']} | "
                f"{row['Empresa']} | {row['Grupo externo resolutor']}"
            )

    st.header("5. Creación")

    st.warning(
        "La creación REAL solo debe hacerse después de revisar "
        "el resultado anterior."
    )

    if st.button("🚀 EJECUTAR PRUEBA REAL", type="primary"):
        resultados = []

        progress = st.progress(0)
        status = st.empty()

        for i, (_, row) in enumerate(test_df.iterrows(), start=1):
            origin = row["ID incidencia origen"]
            status.write(f"Procesando {origin} ({i}/{len(test_df)})")

            if origin in st.session_state.duplicates:
                resultados.append({
                    "ID incidencia origen": origin,
                    "Resultado": "OMITIDA",
                    "Jira Key": st.session_state.duplicates[origin],
                    "Error": "",
                })
            else:
                key, error = crear(
                    row,
                    st.session_state.jira_info,
                )
                resultados.append({
                    "ID incidencia origen": origin,
                    "Resultado": "CREADA" if key else "ERROR",
                    "Jira Key": key or "",
                    "Error": error or "",
                })

            progress.progress(i / len(test_df))

        result_df = pd.DataFrame(resultados)
        st.session_state.last_result = result_df

        st.success("Prueba REAL finalizada.")
        st.dataframe(result_df, use_container_width=True)

        csv = result_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Descargar resultado CSV",
            csv,
            file_name="TMOB_resultado_prueba.csv",
            mime="text/csv",
        )

st.header("6. Carga completa")

st.write(
    f"Registros preparados para la carga completa: **{len(prepared)}**"
)

st.warning(
    "Esta opción creará todas las incidencias no duplicadas. "
    "Úsala solo después de validar la prueba de 5."
)

confirm = st.checkbox(
    "He revisado la prueba y quiero permitir la carga completa."
)

if confirm and st.button("⚠️ CARGAR LAS 376 INCIDENCIAS", type="primary"):
    ids = prepared["ID incidencia origen"].tolist()

    try:
        duplicates_all = buscar_duplicados(
            ids,
            st.session_state.jira_info,
        )
    except Exception as e:
        st.error(f"No se puede iniciar la carga: {e}")
        st.stop()

    resultados = []
    progress = st.progress(0)
    status = st.empty()

    total = len(prepared)

    for i, (_, row) in enumerate(prepared.iterrows(), start=1):
        origin = row["ID incidencia origen"]
        status.write(f"Procesando {origin} ({i}/{total})")

        if origin in duplicates_all:
            resultados.append({
                "ID incidencia origen": origin,
                "Resultado": "OMITIDA",
                "Jira Key": duplicates_all[origin],
                "Error": "",
            })
        else:
            key, error = crear(
                row,
                st.session_state.jira_info,
            )
            resultados.append({
                "ID incidencia origen": origin,
                "Resultado": "CREADA" if key else "ERROR",
                "Jira Key": key or "",
                "Error": error or "",
            })

        progress.progress(i / total)

    final_df = pd.DataFrame(resultados)
    st.session_state.last_result = final_df

    st.success("Carga completa finalizada.")
    st.dataframe(final_df, use_container_width=True)

    csv = final_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Descargar resultado de carga completa",
        csv,
        file_name="TMOB_resultado_carga_completa.csv",
        mime="text/csv",
    )
