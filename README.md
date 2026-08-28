# Portal Roles CENASE v14 — Guardado permanente

Esta versión conserva toda la lógica de v13 y agrega **historial mensual permanente**.

## Qué guarda

Al pulsar **Guardar / actualizar mes**, la aplicación conserva en Supabase los datos procesados de:

- Roles de Gerentes, Administrativos y Operativos.
- Consolidado IESS.
- Reporte de planillas IESS pagadas.

Al abrir un mes guardado, la app reconstruye automáticamente conciliaciones, beneficios, BI, los tres asientos contables y los PDF/Excel sin volver a cargar los archivos.

## Configuración única de Supabase

1. Crea un proyecto gratuito en Supabase.
2. Abre **SQL Editor** y ejecuta el archivo `supabase_setup.sql` incluido en este ZIP.
3. En Supabase copia:
   - Project URL.
   - `service_role` key del proyecto.
4. En Streamlit Cloud entra a tu app > **Settings > Secrets** y agrega:

```toml
[supabase]
url = "TU_PROJECT_URL"
service_role_key = "TU_SERVICE_ROLE_KEY"
```

No subas esas credenciales a GitHub. Streamlit Secrets las mantiene fuera del repositorio.

## Uso

- Para un mes nuevo: carga los 5 archivos, revisa resultados y pulsa **Guardar / actualizar mes**.
- Para consultar un mes anterior: selecciónalo en **Historial mensual** y pulsa **Abrir**.
- Guardar de nuevo el mismo período actualiza ese cierre, no crea duplicados.
- La eliminación requiere marcar una confirmación explícita.

## Archivos del repositorio

- `app.py`
- `requirements.txt`
- `supabase_setup.sql`
- `README.md`

## v15 - PDFs legibles
- Rol Unificado rediseñado como ROL DE PAGOS imprimible, con letra mayor, totales y columna de firma.
- Resumen del Rol ahora replica el control mensual: ingresos, egresos, beneficios pagados y beneficios acumulados.
- Todos los demás PDFs usan fuente mayor y dividen tablas anchas en bloques legibles en vez de reducir excesivamente la letra.

## v17 - Rol PDF tipo tabla dinamica
- El botón **Descargar Rol en PDF** genera ahora el rol mensual en formato resumen, igual al control Excel validado.
- Bloque 1: ingresos por Administrativos, Gerentes, Operativos y Total General.
- Bloque 2: egresos y Neto a Recibir.
- Bloque 3: XIII, XIV y Fondo de Reserva pagados en el Rol.
- Bloque 4: XIII, XIV y Fondo de Reserva acumulados.
- Los cuatro bloques se presentan en una sola pagina A3 horizontal para conservar la letra legible y todas las columnas visibles.
