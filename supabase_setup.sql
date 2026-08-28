-- Ejecutar una sola vez en Supabase > SQL Editor.
create table if not exists public.cenase_monthly_closings (
    period text primary key,
    saved_at timestamptz not null default now(),
    payload text not null
);

-- La APP usa la service_role_key guardada exclusivamente en Streamlit Secrets.
-- No publiques esa clave dentro de GitHub ni dentro de app.py.
alter table public.cenase_monthly_closings enable row level security;
