# TMOB – Importador Web

## Estructura

- `app_tmob_streamlit.py` → pantalla principal
- `pages/01_Incidencias.py` → importador de Incidencias
- `pages/02_Peticiones_WO.py` → importador de Peticiones/WO
- `app_tmob_streamlit_estable_corregida.py` → lógica estable de Incidencias
- `app_tmob_streamlit_peticiones.py` → lógica de Peticiones/WO

## Ejecución local

```bash
streamlit run app_tmob_streamlit.py
```

## Secrets

Mantener los mismos secrets que ya utiliza el importador:

```toml
TMOB_ACCESS_CODE = "..."
JIRA_URL = "https://suport-secom.atlassian.net"
JIRA_EMAIL = "..."
JIRA_API_TOKEN = "..."
```

No introducir credenciales en los archivos Python.

## Uso

1. Entrar en la aplicación.
2. Introducir el código de acceso.
3. Elegir `Incidencias` o `Peticiones / WO`.
4. Trabajar con el módulo seleccionado.
5. Volver al menú mediante la navegación de Streamlit.

Los módulos permanecen separados para que los cambios de uno no alteren la lógica del otro.
