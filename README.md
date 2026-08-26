# Portal de Roles CENASE

Aplicación Streamlit para consolidar mensualmente:
- Rol de Gerentes
- Rol de Administración
- Rol de Operativos

## Funciones
- Carga independiente de los tres roles.
- Normalización automática de las diferencias de estructura.
- KPIs: empleados, ingresos, egresos y neto.
- Filtros por rol, nombre/cédula, cargo, puesto/cliente, neto y días.
- Cuadre aritmético.
- Descarga de un único Excel con hojas:
  - Resumen
  - Consolidado
  - Gerentes
  - Administrativos
  - Operativos

## Publicar en Streamlit Community Cloud
1. Crea un repositorio nuevo en GitHub.
2. Sube `app.py` y `requirements.txt`.
3. En Streamlit Community Cloud, crea una app desde ese repositorio.
4. Selecciona `app.py` como archivo principal.
5. Deploy.

No es necesario subir los roles al repositorio. Los archivos mensuales se cargan desde la pantalla de la app.
