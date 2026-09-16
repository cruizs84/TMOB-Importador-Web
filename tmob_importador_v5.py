# TMOB - Importador Excel -> Jira V5
#
# V5: versión robusta para evitar bloqueos silenciosos.
#
# Cambios principales respecto a V4:
# - No usa createmeta como primera opción para obtener Empresa.
# - Obtiene los contextos/opciones del campo customfield_10369 mediante
#   la API de opciones de campos de Jira y solo usa createmeta como fallback.
# - Timeouts cortos y visibles en cada petición.
# - Preflight de /myself y del proyecto TMOB antes de consultar catálogos.
# - Comprueba explícitamente que el issue type Incidencia existe.
# - Los duplicados se buscan en UNA sola consulta JQL para la selección.
# - Nunca se queda esperando indefinidamente.
# - --validate-only nunca escribe incidencias.
# - --real solo crea después de superar todas las validaciones.
#
# Uso Windows:
#   python tmob_importador_v5.py "INC activas (13).xlsx" --limit 5 --validate-only
#   python tmob_importador_v5.py "INC activas (13).xlsx" --limit 5 --real
#
# Variables de entorno:
#   JIRA_BASE_URL=https://suport-secom.atlassian.net
#   JIRA_EMAIL=...
#   JIRA_API_TOKEN=...
#
# IMPORTANTE:
# El Excel definitivo se lee con header=0 y hoja Report.

import argparse
import os
import sys
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

PROJECT_KEY = "TMOB"
PROJECT_ID = "10166"
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

PRIORIDADES = {"Crítica", "Alta", "Media", "Baja"}

# Equivalencias explícitas conocidas.
EMPRESA_ALIASES = {
    "Autocares Prat": "Autocares PRAT",
    "Autocares PRAT": "Autocares PRAT",
    "MOHN (Grup Baixbús)": "Mohn",
    "MOHN (Grup Baixbus)": "Mohn",
    "Mohn (Grup Baixbus)": "Mohn",
    "Mohn (Grup Baixbús)": "Mohn",
}

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


def log(msg=""):
    print(msg, flush=True)


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
    log(f"[1/8] Leyendo Excel: {path}")
    df = pd.read_excel(path, sheet_name="Report", header=0)
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

    df["ID incidencia origen"] = (
        df["ID incidencia origen"].astype("string").str.strip()
    )
    df = df[
        df["ID incidencia origen"].fillna("").str.match(r"^INC\d+$", na=False)
    ].copy()

    log(f"      Excel cargado: {len(df)} incidencias con ID origen válido.")
    return df


def crear_sesion():
    base_url = os.getenv("JIRA_BASE_URL", "").rstrip("/")
    email = os.getenv("JIRA_EMAIL", "")
    token = os.getenv("JIRA_API_TOKEN", "")

    if not base_url or not email or not token:
        raise RuntimeError(
            "Faltan variables JIRA_BASE_URL, JIRA_EMAIL o JIRA_API_TOKEN."
        )

    # Protección adicional contra el error histórico de 'support' vs 'suport'.
    if "support-secom.atlassian.net" in base_url.lower():
        raise RuntimeError(
            "JIRA_BASE_URL es incorrecta. Debe ser "
            "https://suport-secom.atlassian.net (una sola 'p' en suport)."
        )

    session = requests.Session()
    session.auth = HTTPBasicAuth(email, token)
    session.headers.update({
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "TMOB-Importador-V5/1.0",
    })
    return base_url, session


def request_get(session, url, params=None, timeout=(5, 12)):
    try:
        return session.get(url, params=params, timeout=timeout)
    except requests.exceptions.Timeout as exc:
        raise RuntimeError(f"TIMEOUT al consultar Jira: {url}") from exc
    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"ERROR de conexión con Jira: {exc}") from exc


def request_post(session, url, json=None, timeout=(5, 15)):
    try:
        return session.post(url, json=json, timeout=timeout)
    except requests.exceptions.Timeout as exc:
        raise RuntimeError(f"TIMEOUT al consultar Jira: {url}") from exc
    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"ERROR de conexión con Jira: {exc}") from exc


def preflight(base_url, session):
    log("[2/8] Comprobando conexión con Jira...")
    r = request_get(session, f"{base_url}/rest/api/3/myself")
    if not r.ok:
        raise RuntimeError(
            f"/myself -> HTTP {r.status_code}: {r.text[:500]}"
        )
    me = r.json()
    log(f"      Conectado como: {me.get('displayName') or me.get('emailAddress') or 'usuario Jira'}")

    log("[3/8] Comprobando proyecto TMOB...")
    r = request_get(session, f"{base_url}/rest/api/3/project/{PROJECT_KEY}")
    if not r.ok:
        raise RuntimeError(
            f"/project/{PROJECT_KEY} -> HTTP {r.status_code}: {r.text[:500]}"
        )
    project = r.json()
    if str(project.get("id")) != PROJECT_ID:
        raise RuntimeError(
            f"TMOB encontrado, pero el ID no coincide: {project.get('id')} != {PROJECT_ID}"
        )
    log(f"      Proyecto OK: {project.get('name')} ({project.get('id')})")

    log("[4/8] Comprobando tipo de incidencia...")
    # Endpoint de issue type: no depende de createmeta.
    r = request_get(session, f"{base_url}/rest/api/3/issuetype/{ISSUE_TYPE_ID}")
    if not r.ok:
        raise RuntimeError(
            f"/issuetype/{ISSUE_TYPE_ID} -> HTTP {r.status_code}: {r.text[:500]}"
        )
    issue_type = r.json()
    if issue_type.get("name") != ISSUE_TYPE:
        raise RuntimeError(
            f"El tipo {ISSUE_TYPE_ID} no es '{ISSUE_TYPE}': {issue_type.get('name')}"
        )
    log(f"      Tipo OK: {issue_type.get('name')} ({issue_type.get('id')})")


def obtener_catalogo_empresa(base_url, session):
    """
    Obtiene las opciones de customfield_10369 sin depender inicialmente
    de createmeta.

    La API de Jira expone los contextos de un campo y las opciones de cada
    contexto. Se recorren las páginas necesarias, con timeouts cortos.
    """
    log("[5/8] Obteniendo catálogo real de Empresa desde Jira...")

    contexts_url = f"{base_url}/rest/api/3/field/{FIELD_EMPRESA}/context"
    r = request_get(
        session,
        contexts_url,
        params={"startAt": 0, "maxResults": 100},
    )

    if r.ok:
        data = r.json()
        contexts = data.get("values", [])
        if contexts:
            log(f"      Contextos encontrados: {len(contexts)}")
            catalog = {}

            for ctx in contexts:
                context_id = ctx.get("id")
                if not context_id:
                    continue

                options_url = (
                    f"{base_url}/rest/api/3/field/{FIELD_EMPRESA}"
                    f"/context/{context_id}/option"
                )
                start_at = 0

                while True:
                    rr = request_get(
                        session,
                        options_url,
                        params={
                            "startAt": start_at,
                            "maxResults": 100,
                            "onlyOptions": "true",
                        },
                    )
                    if not rr.ok:
                        log(
                            f"      Aviso: contexto {context_id} -> "
                            f"HTTP {rr.status_code}. Se continúa."
                        )
                        break

                    page = rr.json()
                    values = page.get("values", [])
                    for item in values:
                        name = item.get("value") or item.get("name")
                        if name and not item.get("disabled", False):
                            catalog[str(name).strip()] = item

                    if page.get("isLast", True) or not values:
                        break
                    start_at += len(values)

            if catalog:
                log(f"      Opciones Empresa encontradas: {len(catalog)}")
                return catalog

        log(
            f"      Contextos/opciones no utilizables "
            f"(HTTP {r.status_code if r is not None else 'n/a'})."
        )
    else:
        log(
            f"      Endpoint de contextos -> HTTP {r.status_code}. "
            "Se prueba fallback rápido."
        )

    # Fallback: createmeta, pero con timeout corto y solo una petición.
    log("      Fallback: consultando createmeta (máximo 12 s)...")
    url = (
        f"{base_url}/rest/api/3/issue/createmeta/"
        f"{PROJECT_KEY}/issuetypes/{ISSUE_TYPE_ID}"
    )
    rr = request_get(session, url, timeout=(5, 10))
    if rr.ok:
        data = rr.json()
        fields = data.get("fields", {})
        field = fields.get(FIELD_EMPRESA, {})
        values = field.get("allowedValues", [])
        catalog = {}
        for item in values:
            name = item.get("value") or item.get("name")
            if name:
                catalog[str(name).strip()] = item
        if catalog:
            log(f"      Opciones Empresa encontradas por fallback: {len(catalog)}")
            return catalog

    detail = (
        f"Contextos HTTP {r.status_code if r is not None else 'sin respuesta'}; "
        f"createmeta HTTP {rr.status_code}"
    )
    raise RuntimeError(
        "No se pudo obtener el catálogo de Empresa. " + detail +
        ". No se realiza ninguna carga."
    )


def resolver_empresa(origen, catalogo):
    src = valor_texto(origen)
    if not src:
        return None, "VACIO"

    alias = EMPRESA_ALIASES.get(src)
    if alias:
        if alias in catalogo:
            return alias, "ALIAS"
        return None, f"ALIAS_NO_EXISTE_EN_JIRA:{alias}"

    if src in catalogo:
        return src, "EXACTO"

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

    empresas = out["Empresa origen"].apply(
        lambda x: resolver_empresa(x, catalogo_empresa)
    )
    out["Empresa"] = empresas.apply(lambda x: x[0])
    out["Empresa estado mapeo"] = empresas.apply(lambda x: x[1])

    grupos = out["Grupo externo resolutor origen"].apply(resolver_grupo)
    out["Grupo externo resolutor"] = grupos.apply(lambda x: x[0])
    out["Grupo estado mapeo"] = grupos.apply(lambda x: x[1])

    motivos = out["Motivo de Pendiente origen"].apply(resolver_motivo)
    out["Motivo de Pendiente"] = motivos.apply(lambda x: x[0])
    out["Motivo estado mapeo"] = motivos.apply(lambda x: x[1])

    out["Estado Jira propuesto"] = (
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

    out["Error de mapeo"] = ""

    for idx, row in out.iterrows():
        errors = []

        if not row["Empresa"]:
            errors.append(
                f"Empresa={row['Empresa origen']} "
                f"({row['Empresa estado mapeo']})"
            )

        if not row["Grupo externo resolutor"]:
            errors.append(
                f"Grupo={row['Grupo externo resolutor origen']} "
                f"({row['Grupo estado mapeo']})"
            )

        if pd.isna(row["Estado Jira propuesto"]):
            errors.append(f"Estado={row['Estado origen']}")

        if errors:
            out.at[idx, "Error de mapeo"] = " | ".join(errors)

    return out


def validar(df):
    problemas = []

    for _, row in df.iterrows():
        ident = row["ID incidencia origen"]

        if row["Error de mapeo"]:
            problemas.append((ident, "MAPEO", row["Error de mapeo"]))

        if not valor_texto(row["Resumen"]):
            problemas.append((ident, "DATOS", "Resumen vacío"))

        if valor_texto(row["Prioridad"]) not in PRIORIDADES:
            problemas.append(
                (ident, "DATOS", f"Prioridad no válida: {row['Prioridad']}")
            )

        if pd.isna(row["Fecha y hora de inicio"]):
            problemas.append(
                (ident, "DATOS", "Fecha de envío no válida")
            )

        motivo = valor_texto(row["Motivo de Pendiente origen"])
        estado = valor_texto(row["Estado origen"])

        if estado == "Pendiente" and not motivo:
            problemas.append(
                (ident, "DATOS", "Estado Pendiente sin Status_Reason_Hidden")
            )

    return problemas


def esc_jql(value):
    return str(value).replace("\\", "\\\\").replace('"', '\\"')


def buscar_duplicados(base_url, session, ids):
    """Una única consulta para todos los IDs seleccionados."""
    if not ids:
        return {}

    valores = ", ".join(f'"{esc_jql(x)}"' for x in ids)
    jql = (
        f'project = {PROJECT_KEY} '
        f'AND "{FIELD_ID_ORIGEN}" in ({valores})'
    )

    log(f"[7/8] Comprobando duplicados de {len(ids)} registros en una sola consulta...")
    r = request_post(
        session,
        f"{base_url}/rest/api/3/search/jql",
        json={
            "jql": jql,
            "maxResults": len(ids),
            "fields": ["summary", FIELD_ID_ORIGEN],
        },
        timeout=(5, 15),
    )

    if not r.ok:
        raise RuntimeError(
            f"Búsqueda de duplicados -> HTTP {r.status_code}: {r.text[:1000]}"
        )

    issues = r.json().get("issues", [])
    result = {}

    for issue in issues:
        fields = issue.get("fields", {})
        origin = fields.get(FIELD_ID_ORIGEN)
        if origin:
            result[str(origin)] = issue.get("key")

    return result


def iso_datetime(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, pd.Timestamp):
        dt = valor.to_pydatetime()
    elif isinstance(valor, datetime):
        dt = valor
    else:
        dt = pd.to_datetime(
            valor, errors="coerce", dayfirst=True
        ).to_pydatetime()

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo("Europe/Madrid"))
    else:
        dt = dt.astimezone(ZoneInfo("Europe/Madrid"))

    return dt.isoformat(timespec="seconds")


def construir_fields(row):
    fields = {
        "project": {"key": PROJECT_KEY},
        "issuetype": {"id": ISSUE_TYPE_ID},
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

    if valor_texto(row["Estado origen"]) == "Pendiente":
        motivo = valor_texto(row["Motivo de Pendiente"])
        if motivo:
            fields[FIELD_MOTIVO_PENDIENTE] = motivo

    return fields


def crear_issue(base_url, session, row):
    ident = row["ID incidencia origen"]
    log(f"      Creando {ident}...")
    r = request_post(
        session,
        f"{base_url}/rest/api/3/issue",
        json={"fields": construir_fields(row)},
        timeout=(5, 20),
    )

    if not r.ok:
        raise RuntimeError(
            f"Creación {ident} -> HTTP {r.status_code}: {r.text[:1500]}"
        )

    return r.json().get("key")


def main():
    parser = argparse.ArgumentParser(
        description="Importador seguro TMOB Excel -> Jira V5"
    )
    parser.add_argument("excel")
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Número máximo de registros a procesar. Por defecto 5.",
    )
    parser.add_argument(
        "--real",
        action="store_true",
        help="Crea realmente las incidencias. Sin esta opción es DRY-RUN.",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Solo valida; no crea ni busca duplicados si la validación falla.",
    )
    parser.add_argument(
        "--resultado",
        default="TMOB_resultado_importacion_v5.xlsx",
    )

    args = parser.parse_args()

    if args.limit < 1:
        raise ValueError("--limit debe ser mayor que 0")

    if not os.path.isfile(args.excel):
        raise FileNotFoundError(f"No existe el Excel: {args.excel}")

    log("=" * 72)
    log("TMOB IMPORTADOR V5")
    log("=" * 72)
    log(f"Excel: {args.excel}")
    log(f"Límite de esta ejecución: {args.limit}")
    log(f"Modo: {'REAL' if args.real else 'DRY-RUN'}")
    log("")

    base_url, session = crear_sesion()
    log(f"Jira: {base_url}")

    df = cargar(args.excel)
    preflight(base_url, session)
    catalogo = obtener_catalogo_empresa(base_url, session)

    log("[6/8] Validando mapeos y datos...")
    prepared = preparar(df, catalogo)
    problemas = validar(prepared)

    # Mostrar las 32 empresas o todas las detectadas.
    log("")
    log("=== EMPRESAS DETECTADAS / MAPEADAS ===")
    resumen_emp = (
        prepared[
            ["Empresa origen", "Empresa", "Empresa estado mapeo"]
        ]
        .drop_duplicates()
        .sort_values("Empresa origen")
    )
    log(resumen_emp.to_string(index=False))

    log("")
    log("=== GRUPOS DETECTADOS / MAPEADOS ===")
    resumen_grp = (
        prepared[
            [
                "Grupo externo resolutor origen",
                "Grupo externo resolutor",
                "Grupo estado mapeo",
            ]
        ]
        .drop_duplicates()
        .sort_values("Grupo externo resolutor origen")
    )
    log(resumen_grp.to_string(index=False))

    log("")
    log("=== MOTIVOS DETECTADOS / MAPEADOS ===")
    resumen_mot = (
        prepared[
            [
                "Motivo de Pendiente origen",
                "Motivo de Pendiente",
                "Motivo estado mapeo",
            ]
        ]
        .drop_duplicates()
        .sort_values("Motivo de Pendiente origen")
    )
    log(resumen_mot.to_string(index=False))

    if problemas:
        log("")
        log(f"BLOQUEADO: {len(problemas)} problemas. No se crea nada.")
        for ident, tipo, detalle in problemas[:100]:
            log(f"ERROR | {ident} | {tipo} | {detalle}")

        prepared.to_excel(args.resultado, index=False)
        log(f"Resultado de validación guardado: {args.resultado}")
        return 2

    log("")
    log("VALIDACIÓN COMPLETA: 0 errores.")
    seleccion = prepared.head(args.limit).copy()
    ids = seleccion["ID incidencia origen"].tolist()

    # --validate-only es una validación pura: no consulta duplicados ni crea.
    if args.validate_only:
        log("")
        log("VALIDATE-ONLY: no se consulta duplicidad y no se escribe en Jira.")
        log(f"Registros validados: {len(seleccion)}")
        prepared.to_excel(args.resultado, index=False)
        log(f"Resultado guardado: {args.resultado}")
        return 0

    duplicados = buscar_duplicados(base_url, session, ids)

    log("")
    log("=== RESULTADO DE DUPLICADOS ===")
    for _, row in seleccion.iterrows():
        ident = row["ID incidencia origen"]
        if ident in duplicados:
            log(f"OMITIR | {ident} | ya existe como {duplicados[ident]}")
        else:
            log(
                f"CREAR | {ident} | "
                f"Empresa={row['Empresa']} | "
                f"Grupo={row['Grupo externo resolutor']}"
            )

    if not args.real:
        log("")
        log("DRY-RUN: no se ha creado ninguna incidencia.")
        prepared.to_excel(args.resultado, index=False)
        log(f"Resultado guardado: {args.resultado}")
        return 0

    log("")
    log("[8/8] INICIANDO CREACIÓN REAL...")
    resultados = []

    for _, row in seleccion.iterrows():
        ident = row["ID incidencia origen"]

        if ident in duplicados:
            resultados.append({
                "ID incidencia origen": ident,
                "Resultado": "OMITIDA",
                "Jira Key": duplicados[ident],
                "Error": "",
            })
            continue

        try:
            key = crear_issue(base_url, session, row)
            log(f"      OK -> {ident} -> {key}")
            resultados.append({
                "ID incidencia origen": ident,
                "Resultado": "CREADA",
                "Jira Key": key or "",
                "Error": "",
            })
        except Exception as exc:
            log(f"      ERROR -> {ident} -> {exc}")
            resultados.append({
                "ID incidencia origen": ident,
                "Resultado": "ERROR",
                "Jira Key": "",
                "Error": str(exc),
            })

    result_df = pd.DataFrame(resultados)
    result_df.to_excel(args.resultado, index=False)

    log("")
    log("=" * 72)
    log("PROCESO FINALIZADO")
    log("=" * 72)
    log(f"Resultado: {args.resultado}")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log("\nProceso cancelado por el usuario.")
        raise SystemExit(130)
    except Exception as exc:
        log("")
        log("ERROR FATAL:")
        log(str(exc))
        log("")
        log("No se ha realizado ninguna creación en Jira.")
        raise SystemExit(1)
