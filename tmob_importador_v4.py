# TMOB - Importador seguro Excel -> Jira v4
#
# Objetivos v4:
# - Leer la hoja Report del Excel definitivo con header=0.
# - Resolver las 32 empresas mediante catálogo Jira + alias explícitos.
# - Resolver grupos mediante equivalencias explícitas.
# - Mantener los motivos de pendiente literalmente, salvo la equivalencia
#   aprobada: "Incidencia original pendiente" -> "Incidencia principal pendiente".
# - Detectar y BLOQUEAR la carga si queda cualquier valor sin mapear.
# - Comprobar duplicados por ID incidencia origen antes de crear.
# - Crear siempre en CREADA; Ámbito = Pendiente de asignación; assignee vacío.
# - Interpretar fechas del Excel como Europe/Madrid (incluye +01/+02 según fecha).
#
# Uso:
#   python tmob_importador_v4.py "INC activas (13).xlsx" --limit 5
#   python tmob_importador_v4.py "INC activas (13).xlsx" --limit 5 --real
#   python tmob_importador_v4.py "INC activas (13).xlsx" --validate-only
#
# Variables:
#   JIRA_BASE_URL=https://suport-secom.atlassian.net
#   JIRA_EMAIL=...
#   JIRA_API_TOKEN=...

import argparse
import os
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

PROJECT_KEY = "TMOB"
ISSUE_TYPE = "Incidencia"
ISSUE_TYPE_ID = "10223"

FIELD_ID_ORIGEN = "customfield_10402"
FIELD_REQ = "customfield_10296"
FIELD_GRUPO = "customfield_10330"
FIELD_MOTIVO_PENDIENTE = "customfield_10332"
FIELD_AMBITO = "customfield_10295"
FIELD_EMPRESA = "customfield_10369"
FIELD_FECHA_INICIO = "customfield_10472"
AMBITO_PROVISIONAL = "Pendiente de asignación"

# Las 32 empresas detectadas en la extracción histórica deben quedar
# resueltas contra el catálogo real de Jira. Para variantes conocidas,
# se define aquí el nombre Jira. Si no hay alias, se intenta coincidencia
# normalizada contra el catálogo obtenido de Jira.
EMPRESA_ALIASES = {
    # Variantes conocidas / normalizaciones documentadas
    "Autocares Prat": "Autocares PRAT",
    "Autocares PRAT": "Autocares PRAT",
    "MOHN (Grup Baixbús)": "Mohn",
    "MOHN (Grup Baixbus)": "Mohn",
    "Mohn (Grup Baixbus)": "Mohn",
    "Mohn (Grup Baixbús)": "Mohn",
}

# Equivalencias explícitas de grupos de origen -> Grupo externo resolutor Jira.
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

STATUS_MAP = {
    "Asignado": "ASIGNADA",
    "En curso": "EN PROGRESO",
    "Pendiente": "PENDIENTE",
}

# Equivalencia única que no debe hacerse de forma automática fuera de este mapa.
MOTIVO_MAP = {
    "Incidencia original pendiente": "Incidencia principal pendiente",
}

PRIORIDADES = {"Crítica", "Alta", "Media", "Baja"}


def normalizar(valor):
    if valor is None or pd.isna(valor):
        return ""
    s = str(valor).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())
    return s.casefold()


def valor_texto(valor):
    if valor is None or pd.isna(valor):
        return None
    s = str(valor).strip()
    return s if s else None


def cargar(path):
    # IMPORTANTE: el Excel definitivo tiene los encabezados en la primera fila.
    df = pd.read_excel(path, sheet_name="Report", header=0)
    df.columns = [str(c).strip() for c in df.columns]

    required = [
        "ID de la incidencia*+", "ID de petición de servicio", "Empresa*+",
        "Prioridad*", "Fecha de envío", "Fecha de última modificación",
        "Resumen*", "Grupo asignado*+", "Estado*", "Status_Reason_Hidden",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Faltan columnas en el Excel: " + ", ".join(missing))

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

    df["ID incidencia origen"] = df["ID incidencia origen"].astype("string").str.strip()
    df = df[df["ID incidencia origen"].fillna("").str.match(r"^INC\d+$", na=False)].copy()
    return df


def obtener_catalogo_empresa(base_url, session):
    """Obtiene allowedValues del campo Empresa desde createmeta."""
    urls = [
        f"{base_url}/rest/api/3/issue/createmeta/{PROJECT_KEY}/issuetypes/{ISSUE_TYPE_ID}",
        f"{base_url}/rest/api/3/issue/createmeta?projectKeys={PROJECT_KEY}&issuetypeNames={ISSUE_TYPE}&expand=projects.issuetypes.fields",
    ]
    last_error = None
    for url in urls:
        r = session.get(url, timeout=30)
        if not r.ok:
            last_error = f"HTTP {r.status_code}: {r.text[:500]}"
            continue
        data = r.json()
        fields = data.get("fields")
        if fields is None:
            projects = data.get("projects", [])
            if projects:
                its = projects[0].get("issuetypes", [])
                if its:
                    fields = its[0].get("fields", {})
        if not fields:
            continue
        f = fields.get(FIELD_EMPRESA, {})
        values = f.get("allowedValues", [])
        catalog = {}
        for item in values:
            name = item.get("value") or item.get("name")
            if name:
                catalog[str(name).strip()] = item
        if catalog:
            return catalog
    raise RuntimeError("No se pudo obtener el catálogo de Empresa de Jira. " + str(last_error or "sin allowedValues"))


def resolver_empresa(origen, catalogo):
    src = valor_texto(origen)
    if not src:
        return None, "VACIO"

    # 1. Alias explícito.
    alias = EMPRESA_ALIASES.get(src)
    if alias:
        if alias in catalogo:
            return alias, "ALIAS"
        # Si el alias aprobado no existe en el catálogo, no se sustituye por otro.
        return None, f"ALIAS_NO_EXISTE_EN_JIRA:{alias}"

    # 2. Coincidencia exacta.
    if src in catalogo:
        return src, "EXACTO"

    # 3. Coincidencia normalizada: acentos, mayúsculas y espacios.
    n = normalizar(src)
    matches = [name for name in catalogo if normalizar(name) == n]
    if len(matches) == 1:
        return matches[0], "NORMALIZADO"
    if len(matches) > 1:
        return None, "AMBIGUO:" + " | ".join(matches)

    return None, "NO_ENCONTRADO"


def resolver_grupo(origen):
    src = valor_texto(origen)
    if not src:
        return None, "VACIO"
    if src in GROUP_MAP:
        return GROUP_MAP[src], "EXPLICITO"
    return None, "NO_ENCONTRADO"


def resolver_motivo(origen):
    src = valor_texto(origen)
    if not src:
        return None, "VACIO"
    if src in MOTIVO_MAP:
        return MOTIVO_MAP[src], "EQUIVALENCIA_APROBADA"
    return src, "LITERAL"


def preparar(df, catalogo_empresa):
    out = df.copy()
    empresas = out["Empresa origen"].apply(lambda x: resolver_empresa(x, catalogo_empresa))
    out["Empresa"] = empresas.apply(lambda x: x[0])
    out["Empresa estado mapeo"] = empresas.apply(lambda x: x[1])

    grupos = out["Grupo externo resolutor origen"].apply(resolver_grupo)
    out["Grupo externo resolutor"] = grupos.apply(lambda x: x[0])
    out["Grupo estado mapeo"] = grupos.apply(lambda x: x[1])

    motivos = out["Motivo de Pendiente origen"].apply(resolver_motivo)
    out["Motivo de Pendiente"] = motivos.apply(lambda x: x[0])
    out["Motivo estado mapeo"] = motivos.apply(lambda x: x[1])

    out["Estado Jira propuesto"] = out["Estado origen"].astype("string").str.strip().map(STATUS_MAP)
    out["Fecha y hora de inicio"] = pd.to_datetime(out["Fecha y hora de inicio"], errors="coerce", dayfirst=True)

    out["Error de mapeo"] = ""
    for idx, row in out.iterrows():
        errors = []
        if not row["Empresa"]:
            errors.append(f"Empresa={row['Empresa origen']} ({row['Empresa estado mapeo']})")
        if not row["Grupo externo resolutor"]:
            errors.append(f"Grupo={row['Grupo externo resolutor origen']} ({row['Grupo estado mapeo']})")
        if row["Estado Jira propuesto"] is None or pd.isna(row["Estado Jira propuesto"]):
            errors.append(f"Estado={row['Estado origen']}")
        if errors:
            out.at[idx, "Error de mapeo"] = " | ".join(errors)
    return out


def validar(df):
    problemas = []
    for _, r in df.iterrows():
        ident = r["ID incidencia origen"]
        if r["Error de mapeo"]:
            problemas.append((ident, "MAPEO", r["Error de mapeo"]))
        if not valor_texto(r["Resumen"]):
            problemas.append((ident, "DATOS", "Resumen vacío"))
        if valor_texto(r["Prioridad"]) not in PRIORIDADES:
            problemas.append((ident, "DATOS", f"Prioridad no válida: {r['Prioridad']}"))
        if pd.isna(r["Fecha y hora de inicio"]):
            problemas.append((ident, "DATOS", "Fecha de envío no válida"))
        motivo = valor_texto(r["Motivo de Pendiente origen"])
        estado = valor_texto(r["Estado origen"])
        if estado == "Pendiente" and not motivo:
            problemas.append((ident, "DATOS", "Estado Pendiente sin Status_Reason_Hidden"))
    return problemas


def jira_session():
    base_url = os.getenv("JIRA_BASE_URL", "").rstrip("/")
    email = os.getenv("JIRA_EMAIL", "")
    token = os.getenv("JIRA_API_TOKEN", "")
    if not base_url or not email or not token:
        raise RuntimeError("Faltan JIRA_BASE_URL, JIRA_EMAIL o JIRA_API_TOKEN")
    s = requests.Session()
    s.auth = HTTPBasicAuth(email, token)
    s.headers.update({"Accept": "application/json", "Content-Type": "application/json"})
    return base_url, s


def esc_jql(value):
    return str(value).replace("\\", "\\\\").replace('"', '\\"')


def buscar_por_id_origen(base_url, session, id_origen):
    jql = f'project = {PROJECT_KEY} AND issuetype = "{ISSUE_TYPE}" AND "{FIELD_ID_ORIGEN}" = "{esc_jql(id_origen)}"'
    r = session.post(
        f"{base_url}/rest/api/3/search/jql",
        json={"jql": jql, "maxResults": 2, "fields": ["summary", FIELD_ID_ORIGEN]},
        timeout=30,
    )
    if not r.ok:
        raise RuntimeError(f"Error buscando {id_origen}: HTTP {r.status_code} {r.text[:1000]}")
    issues = r.json().get("issues", [])
    return {"exists": bool(issues), "key": issues[0].get("key") if issues else None}


def iso_datetime(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, pd.Timestamp):
        dt = valor.to_pydatetime()
    elif isinstance(valor, datetime):
        dt = valor
    else:
        dt = pd.to_datetime(valor, errors="coerce", dayfirst=True).to_pydatetime()
    # Excel se interpreta como hora local de Madrid. ZoneInfo aplica +01/+02 según fecha.
    local = dt.replace(tzinfo=ZoneInfo("Europe/Madrid"))
    return local.isoformat(timespec="seconds")


def construir_fields(row):
    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"name": ISSUE_TYPE},
        "summary": valor_texto(row["Resumen"]),
        FIELD_ID_ORIGEN: valor_texto(row["ID incidencia origen"]),
        FIELD_AMBITO: {"value": AMBITO_PROVISIONAL},
    }
    req = valor_texto(row["REQ incidencia"])
    if req:
        fields[FIELD_REQ] = req
    empresa = valor_texto(row["Empresa"])
    if empresa:
        fields[FIELD_EMPRESA] = {"value": empresa}
    prioridad = valor_texto(row["Prioridad"])
    if prioridad:
        fields["priority"] = {"name": prioridad}
    fecha = iso_datetime(row["Fecha y hora de inicio"])
    if fecha:
        fields[FIELD_FECHA_INICIO] = fecha
    grupo = valor_texto(row["Grupo externo resolutor"])
    if grupo:
        fields[FIELD_GRUPO] = grupo
    # El motivo solo se importa si el origen estaba Pendiente.
    if valor_texto(row["Estado origen"]) == "Pendiente":
        motivo = valor_texto(row["Motivo de Pendiente"])
        if motivo:
            fields[FIELD_MOTIVO_PENDIENTE] = motivo
    # NO se importa Estado, Updated ni assignee.
    return fields


def crear_issue(base_url, session, row):
    r = session.post(f"{base_url}/rest/api/3/issue", json={"fields": construir_fields(row)}, timeout=30)
    if not r.ok:
        raise RuntimeError(f"Error creando {row['ID incidencia origen']}: HTTP {r.status_code} {r.text[:1500]}")
    return r.json().get("key")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("--limit", type=int, default=5)
    ap.add_argument("--real", action="store_true")
    ap.add_argument("--validate-only", action="store_true")
    ap.add_argument("--resultado", default="TMOB_resultado_importacion_v4.xlsx")
    args = ap.parse_args()

    base_url, session = jira_session()
    df = cargar(args.excel)
    print(f"Excel: {len(df)} incidencias válidas")

    catalogo = obtener_catalogo_empresa(base_url, session)
    print(f"Catálogo Empresa Jira: {len(catalogo)} opciones")

    prepared = preparar(df, catalogo)
    problemas_globales = validar(prepared)

    # Resumen de mapeo por valor, útil para auditar las 32 empresas.
    print("\nEMPRESAS:")
    resumen_emp = prepared[["Empresa origen", "Empresa", "Empresa estado mapeo"]].drop_duplicates().sort_values("Empresa origen")
    print(resumen_emp.to_string(index=False))

    print("\nGRUPOS:")
    resumen_grp = prepared[["Grupo externo resolutor origen", "Grupo externo resolutor", "Grupo estado mapeo"]].drop_duplicates().sort_values("Grupo externo resolutor origen")
    print(resumen_grp.to_string(index=False))

    print("\nMOTIVOS:")
    resumen_mot = prepared[["Motivo de Pendiente origen", "Motivo de Pendiente", "Motivo estado mapeo"]].drop_duplicates().sort_values("Motivo de Pendiente origen")
    print(resumen_mot.to_string(index=False))

    if problemas_globales:
        print(f"\nBLOQUEADO: {len(problemas_globales)} problemas detectados. No se crea nada.")
        for p in problemas_globales[:100]:
            print(f"ERROR | {p[0]} | {p[1]} | {p[2]}")
        prepared.to_excel(args.resultado, index=False)
        return 2

    print("\nVALIDACIÓN: 0 errores de mapeo/datos.")

    seleccion = prepared.head(args.limit).copy()
    print(f"Registros seleccionados para esta ejecución: {len(seleccion)}")

    if args.validate_only or not args.real:
        print("DRY-RUN: no se escribirá nada en Jira.")
        for _, row in seleccion.iterrows():
            dup = buscar_por_id_origen(base_url, session, row["ID incidencia origen"])
            estado = "OMITIR" if dup["exists"] else "CREAR"
            print(f"{estado} | {row['ID incidencia origen']} | Empresa={row['Empresa']} | Grupo={row['Grupo externo resolutor']}")
        return 0

    resultados = []
    for _, row in seleccion.iterrows():
        ident = row["ID incidencia origen"]
        try:
            dup = buscar_por_id_origen(base_url, session, ident)
            if dup["exists"]:
                print(f"OMITIR | {ident} | ya existe como {dup['key']}")
                resultados.append({"ID incidencia origen": ident, "Resultado": "OMITIDA", "Jira Key": dup["key"], "Error": ""})
                continue
            key = crear_issue(base_url, session, row)
            print(f"CREAR | {ident} | {key}")
            resultados.append({"ID incidencia origen": ident, "Resultado": "CREADA", "Jira Key": key, "Error": ""})
        except Exception as exc:
            print(f"ERROR | {ident} | {exc}")
            resultados.append({"ID incidencia origen": ident, "Resultado": "ERROR", "Jira Key": "", "Error": str(exc)})

    pd.DataFrame(resultados).to_excel(args.resultado, index=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
