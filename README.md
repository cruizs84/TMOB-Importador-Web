# TMOB Jira Importador Web

Aplicación Streamlit para importar `INC Activas.xlsx` a Jira Cloud TMOB.

## Comportamiento

- Comprueba `ID incidencia origen` antes de crear.
- Si existe: `OMITIDA`.
- Si no existe: crea `Incidencia`.
- Toda incidencia nueva se crea en `CREADA`.
- `Ámbito` se establece en `Pendiente de asignación`.
- `Persona asignada` queda vacía.
- `Status_Reason_Hidden` se copia a `Motivo de Pendiente`.
- No se asignan Ámbito operativo ni técnico automáticamente.
- Prueba por defecto: 5 incidencias.

## Requisito en Jira

Debe existir la opción de Ámbito:

`Pendiente de asignación`

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Despliegue web

Se puede desplegar en Streamlit Community Cloud desde un repositorio GitHub.
