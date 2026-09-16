import io
import os
import time
import unicodedata
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
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


def _xml_xls_to_dataframe(upload):
    """Read Excel 2003 XML (.xls exported by Remedy/AR System) without xlrd."""
    upload.seek(0)
    raw = upload.read()
    root = ET.fromstring(raw)
    SS = "urn:schemas-microsoft-com:office:spreadsheet"
    ns = {"ss": SS}
    worksheets = root.findall("ss:Worksheet", ns)
    report = next((ws for ws in worksheets if ws.attrib.get(f"{{{SS}}}Name") == "Report"), None)
    if report is None:
        raise ValueError('No se encontró la hoja "Report" en el XLS/XML.')
    rows = report.findall(".//ss:Table/ss:Row", ns)
    if len(rows) < 2:
        raise ValueError('La hoja "Report" no contiene cabecera y datos.')

    def row_values(row):
        values = []
        next_index = 1
        for cell in row.findall("ss:Cell", ns):
            index = int(cell.attrib.get(f"{{{SS}}}Index", next_index))
            while len(values) < index - 1:
                values.append(None)
            data = cell.find("ss:Data", ns)
            values.append(data.text if data is not None else None)
            next_index = index + 1
        return values

    raw_rows = [row_values(row) for row in rows]
    headers = raw_rows[1]
    width = len(headers)
    data = []
    for values in raw_rows[2:]:
        values = values + [None] * max(0, width - len(values))
        data.append(values[:width])
    return pd.DataFrame(data, columns=headers)


def load_excel(upload):
    name = (getattr(upload, "name", "") or "").lower()
    upload.seek(0)
    if name.endswith(".xls"):
        # The supplied Remedy export is Excel 2003 XML, not binary BIFF.
        df = _xml_xls_to_dataframe(upload)
    else:
        df = pd.read_excel(upload, sheet_name="Report", header=0)

    df.columns = [str(c).strip() for c in df.columns]
    rename = {
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
    }
    df = df.rename(columns=rename)

    expected = [
        "ID incidencia origen", "REQ incidencia", "Empresa", "Prioridad",
        "Fecha y hora de inicio", "Fecha última modificación origen", "Resumen",
        "Grupo externo resolutor origen", "Estado origen", "Motivo de Pendiente",
    ]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError("El Excel no coincide con la estructura esperada. Faltan: " + ", ".join(missing))

    # Only rows with a real source incident ID are importable. This deliberately
    # ignores blank formatting rows and the trailing export timestamp row.
    df["ID incidencia origen"] = df["ID incidencia origen"].astype("string").str.strip()
    df = df[df["ID incidencia origen"].str.fullmatch(r"INC\d+", na=False)].copy()
    if df["ID incidencia origen"].duplicated().any():
        duplicated = sorted(df.loc[df["ID incidencia origen"].duplicated(keep=False), "ID incidencia origen"].unique())
        raise ValueError("El Excel contiene IDs de incidencia duplicados: " + ", ".join(duplicated[:20]))

    for c in ["REQ incidencia", "Empresa", "Prioridad", "Resumen", "Grupo externo resolutor origen", "Estado origen", "Motivo de Pendiente"]:
        df[c] = df[c].apply(clean)

    df["Fecha y hora de inicio"] = pd.to_datetime(
        df["Fecha y hora de inicio"], errors="coerce", dayfirst=True
    )
    df["Fecha última modificación origen"] = pd.to_datetime(
        df["Fecha última modificación origen"], errors="coerce", dayfirst=True
    )
    df["Grupo externo resolutor"] = df["Grupo externo resolutor origen"]
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


def jira_preflight(base_url, session):
    """Diagnostic-only Jira connectivity check. No Jira issue is created/modified."""
    import socket
    from urllib.parse import urlparse

    parsed = urlparse(base_url)
    host = parsed.hostname or ""
    if not host:
        raise RuntimeError(f"JIRA_URL no es válida: {base_url}")

    diagnostics = []

    # 1) Basic HTTPS reachability to the site root.
    root_url = f"{parsed.scheme or 'https'}://{host}/"
    try:
        r = session.get(root_url, timeout=30, allow_redirects=True)
        diagnostics.append({
            "check": "Sitio Jira (raíz)",
            "url": r.url,
            "status": r.status_code,
            "headers": {k: v for k, v in r.headers.items() if k.lower() in {"server", "location", "x-arequestid", "x-transaction-id", "x-seraph-loginreason"}},
            "body": r.text[:500].replace("\n", " "),
        })
    except requests.RequestException as exc:
        diagnostics.append({"check": "Sitio Jira (raíz)", "url": root_url, "network_error": str(exc)})

    # 2) DNS resolution from the Streamlit runtime.
    try:
        ips = sorted({item[4][0] for item in socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)})
        diagnostics.append({"check": "DNS", "host": host, "ips": ips})
    except Exception as exc:
        diagnostics.append({"check": "DNS", "host": host, "network_error": str(exc)})

    # 3) Authenticated Jira identity endpoint.
    myself_url = f"{base_url}/rest/api/3/myself"
    try:
        r = session.get(myself_url, timeout=30, allow_redirects=False)
        diagnostics.append({
            "check": "Jira API /myself",
            "url": r.url,
            "status": r.status_code,
            "headers": {k: v for k, v in r.headers.items() if k.lower() in {"server", "location", "x-arequestid", "x-transaction-id", "x-seraph-loginreason", "www-authenticate"}},
            "body": r.text[:1000].replace("\n", " "),
        })
    except requests.RequestException as exc:
        diagnostics.append({"check": "Jira API /myself", "url": myself_url, "network_error": str(exc)})

    # 4) Authenticated project endpoint, only if /myself succeeded.
    myself_status = next((d.get("status") for d in diagnostics if d.get("check") == "Jira API /myself"), None)
    if myself_status == 200:
        project_url = f"{base_url}/rest/api/3/project/{PROJECT_KEY}"
        try:
            r = session.get(project_url, timeout=30, allow_redirects=False)
            diagnostics.append({
                "check": "Jira API /project/TMOB",
                "url": r.url,
                "status": r.status_code,
                "headers": {k: v for k, v in r.headers.items() if k.lower() in {"server", "location", "x-arequestid", "x-transaction-id", "x-seraph-loginreason", "www-authenticate"}},
                "body": r.text[:1000].replace("\n", " "),
            })
        except requests.RequestException as exc:
            diagnostics.append({"check": "Jira API /project/TMOB", "url": project_url, "network_error": str(exc)})

    myself = next((d for d in diagnostics if d.get("check") == "Jira API /myself" and d.get("status") == 200), None)
    project = next((d for d in diagnostics if d.get("check") == "Jira API /project/TMOB" and d.get("status") == 200), None)
    return diagnostics, myself, project


def exists_issue(base_url, session, id_origen):
    # Use the field name in JQL (rather than the customfield_* key). Jira Cloud
    # currently exposes both GET/POST forms of the enhanced /search/jql API,
    # while some sites can transiently reject one form with HTTP 404.
    jql = (
        f'project = {PROJECT_KEY} AND issuetype = "{ISSUE_TYPE}" '
        f'AND "ID incidencia origen" = "{esc_jql(id_origen)}"'
    )
    payload = {"jql": jql, "maxResults": 2, "fields": ["summary", FIELD_ID_ORIGEN]}
    attempts = [
        ("POST /search/jql", lambda: session.post(f"{base_url}/rest/api/3/search/jql", json=payload, timeout=30)),
        ("GET /search/jql", lambda: session.get(f"{base_url}/rest/api/3/search/jql", params=payload, timeout=30)),
        ("POST /search", lambda: session.post(f"{base_url}/rest/api/3/search", json=payload, timeout=30)),
    ]
    errors = []
    for label, call in attempts:
        try:
            r = call()
        except requests.RequestException as exc:
            errors.append(f"{label}: {exc}")
            continue
        if r.ok:
            issues = r.json().get("issues", [])
            return issues[0]["key"] if issues else None
        errors.append(f"{label}: HTTP {r.status_code} {r.text[:500]}")
    raise RuntimeError(f"Error buscando {id_origen}. Intentos API Jira: " + " | ".join(errors))


def iso_datetime(v):
    if pd.isna(v):
        return None
    dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    # Source timestamps are local Barcelona/Madrid time. Let ZoneInfo apply
    # the correct +01:00/+02:00 offset for the date instead of hard-coding +02:00.
    return dt.to_pydatetime().replace(tzinfo=ZoneInfo("Europe/Madrid")).isoformat()


def source_to_jira_value(field_name, source_name):
    """Return the Jira select-field value to send. Jira validates the value server-side.
    No createmeta call is used because TMOB is Team-managed and the create-metadata
    endpoint can return 404 even when the project and issue type are valid.
    """
    value = clean(source_name)
    if field_name == "empresa" and value:
        return EMPRESA_ALIASES.get(value, value)
    if field_name == "grupo" and value:
        if value not in GROUP_OPTIONS:
            raise ValueError(f"Grupo externo sin mapeo definido: {value}")
        return value
    return value


def validate_row_mappings(row):
    """Validate source-side mappings before creation without calling createmeta."""
    errors = []
    company = clean(row["Empresa"])
    group = clean(row["Grupo externo resolutor"])
    reason = clean(row["Motivo de Pendiente"])
    if not AMBITO_PROVISIONAL:
        errors.append("Ámbito provisional vacío en configuración.")
    if group and group not in GROUP_OPTIONS:
        errors.append(f"Grupo externo sin mapeo definido: {group}")
    # Do not silently translate source reasons. Jira will validate the exact value
    # during creation; in particular, "Incidencia original pendiente" is not
    # silently changed to another Jira option.
    return errors


def create_issue(base_url, session, row):
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

    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"id": ISSUE_TYPE_ID},
        "summary": summary,
        FIELD_ID_ORIGEN: id_origen,
        FIELD_REQ: req,
        FIELD_AMBITO: {"value": AMBITO_PROVISIONAL},
        FIELD_FECHA_INICIO: start,
    }

    if company:
        fields[FIELD_EMPRESA] = {"value": source_to_jira_value("empresa", company)}

    if priority:
        fields["priority"] = {"name": priority}

    if group:
        fields[FIELD_GRUPO] = {"value": source_to_jira_value("grupo", group)}

    if reason:
        # Exact source text only. No automatic synonym/translation.
        fields[FIELD_MOTIVO] = {"value": reason}

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
    st.caption("Introduce el código de acceso de TMOB.")
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

upload = st.file_uploader("Sube el informe de incidencias activas", type=["xlsx", "xls"])

if upload:
    try:
        df = load_excel(upload)
        st.success(f"Excel cargado: {len(df)} incidencias con ID origen válido.")
        st.caption("Se importan todos los registros con ID INC válido. Las filas vacías o de fecha del pie del export se ignoran. No hay límite fijo de 5 registros.")
        with st.expander("Estructura detectada y tratamiento de campos", expanded=False):
            mapping_rows = pd.DataFrame([
                ["ID de la incidencia*+", "ID incidencia origen", "IMPORTAR", "customfield_10402"],
                ["ID de petición de servicio", "REQ incidencia", "IMPORTAR", "customfield_10296"],
                ["Empresa*+", "Empresa", "IMPORTAR", "customfield_10369"],
                ["Prioridad*", "Prioridad", "IMPORTAR", "priority"],
                ["Fecha de envío", "Fecha y hora de inicio", "IMPORTAR", "customfield_10472"],
                ["Fecha de última modificación", "Fecha última modificación origen", "NO SE ESCRIBE", "Jira updated es de sistema"],
                ["Resumen*", "Resumen", "IMPORTAR", "summary"],
                ["Grupo asignado*+", "Grupo externo resolutor", "IMPORTAR", "customfield_10330"],
                ["Estado*", "Estado origen", "NO SE COPIA", "todas las altas nuevas entran en CREADA"],
                ["Status_Reason_Hidden", "Motivo de Pendiente", "IMPORTAR SI EXISTE", "customfield_10332"],
            ], columns=["Campo XLS", "Campo interno", "Tratamiento", "Jira"])
            st.dataframe(mapping_rows, use_container_width=True, hide_index=True)

        preview_cols = [c for c in ["ID incidencia origen", "REQ incidencia", "Resumen", "Empresa",
                                     "Prioridad", "Grupo externo resolutor", "Estado origen",
                                     "Motivo de Pendiente"] if c in df.columns]
        st.dataframe(df[preview_cols].head(20), use_container_width=True)

        max_rows = min(5, len(df))
        limit = st.number_input(
            "Número de incidencias a procesar en esta prueba",
            min_value=1, max_value=len(df), value=max_rows
        )
        test_df = df.head(int(limit)).copy()

        st.subheader("Comportamiento de importación")
        st.write("**Nueva incidencia:** CREADA · Ámbito = Pendiente de asignación · Persona asignada = vacía")
        st.write("**Si ID incidencia origen ya existe:** OMITIR")
        st.write("**Status_Reason_Hidden:** se copia exactamente a Motivo de Pendiente; no se hacen traducciones automáticas.")

        try:
            base, sess = jira_session("https://support-secom.atlassian.net")
        except Exception as e:
            st.error(str(e))
            st.stop()

        if st.button("🔎 Comprobar conexión y duplicados (sin crear)", type="secondary"):
            try:
                # First prove that the Streamlit server can authenticate to Jira and
                # access TMOB. Only then perform duplicate searches. This separates
                # authentication/site availability errors from JQL/search errors.
                diagnostics, myself_ok, project_ok = jira_preflight(base, sess)
                with st.expander("Diagnóstico técnico de conexión", expanded=True):
                    st.write(f"URL base utilizada: `{base}`")
                    for d in diagnostics:
                        label = d.get("check", "Comprobación")
                        status = d.get("status")
                        if status is not None:
                            st.write(f"**{label}** → HTTP {status}")
                        elif d.get("network_error"):
                            st.write(f"**{label}** → error de red: {d['network_error']}")
                        if d.get("url"):
                            st.caption(d["url"])
                        if d.get("ips"):
                            st.write("IPs resueltas:", d["ips"])
                        if d.get("headers"):
                            st.json(d["headers"])
                        if d.get("body"):
                            st.code(d["body"], language="text")
                if not myself_ok:
                    st.session_state["preflight_ok"] = False
                    raise RuntimeError("La autenticación contra Jira no ha devuelto HTTP 200 en /rest/api/3/myself. No se ejecutará ninguna búsqueda ni creación.")
                if not project_ok:
                    st.session_state["preflight_ok"] = False
                    raise RuntimeError("La autenticación funciona, pero no se ha podido acceder al proyecto TMOB. No se ejecutará ninguna búsqueda ni creación.")
                myself_json = sess.get(f"{base}/rest/api/3/myself", timeout=30).json()
                project_json = sess.get(f"{base}/rest/api/3/project/{PROJECT_KEY}", timeout=30).json()
                st.success(
                    f"Conexión Jira OK · usuario API: {myself_json.get('displayName','')} · "
                    f"proyecto: {project_json.get('key', PROJECT_KEY)} — {project_json.get('name','')}"
                )
                rows = []
                preflight_errors = []
                for _, row in test_df.iterrows():
                    id0 = clean(row["ID incidencia origen"])
                    mapping_errors = validate_row_mappings(row)
                    existing = exists_issue(base, sess, id0)
                    result = "OMITIR" if existing else ("ERROR MAPEADO" if mapping_errors else "CREAR")
                    rows.append({
                        "ID incidencia origen": id0,
                        "Resultado": result,
                        "Jira existente": existing or "",
                        "Validación": "; ".join(mapping_errors),
                    })
                    if mapping_errors and not existing:
                        preflight_errors.extend([f"{id0}: {e}" for e in mapping_errors])
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
                st.session_state["preflight_ok"] = not preflight_errors
                if preflight_errors:
                    st.error("La precomprobación detectó problemas de catálogo. No se permite crear todavía.")
                    for e in preflight_errors:
                        st.write(f"- {e}")
                else:
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
                                key = create_issue(base, sess, row)
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
    st.info("Carga el informe de incidencias activas (.xlsx o .xls) para empezar.")
