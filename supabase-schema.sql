-- Weber Tracker — esquema de Supabase
-- Ejecutar completo en: Supabase → SQL Editor → New query → Run

-- ── Tablas ──────────────────────────────────────────────────────────

create table if not exists task_overrides (
  task_id text primary key,
  estado text,
  link text,
  updated_at timestamptz not null default now()
);

create table if not exists notes (
  id text primary key,
  text text not null,
  author text,
  created_at timestamptz not null default now()
);

create table if not exists custom_tasks (
  id text primary key,
  tarea text not null,
  area text,
  tema text,
  responsable text,
  fase integer,
  inicio text,
  cierre text,
  estado text,
  obs text,
  link text,
  created_at timestamptz not null default now()
);

create table if not exists meetings (
  id text primary key,
  fecha text,
  responsable text,
  duracion text,
  resumen text,
  created_at timestamptz not null default now()
);

-- ── Row Level Security ──────────────────────────────────────────────
-- El tracker no tiene login (igual que la versión anterior en
-- localStorage): cualquiera con el link del sitio puede leer y escribir.
-- La anon key queda embebida en el HTML público, así que estas políticas
-- abiertas mantienen el mismo nivel de acceso que ya tenía la app.
-- Si en el futuro querés pedir login, avisame y sumamos autenticación.

alter table task_overrides enable row level security;
alter table notes enable row level security;
alter table custom_tasks enable row level security;
alter table meetings enable row level security;

create policy "public select task_overrides" on task_overrides for select using (true);
create policy "public insert task_overrides" on task_overrides for insert with check (true);
create policy "public update task_overrides" on task_overrides for update using (true);

create policy "public select notes" on notes for select using (true);
create policy "public insert notes" on notes for insert with check (true);
create policy "public update notes" on notes for update using (true);
create policy "public delete notes" on notes for delete using (true);

create policy "public select custom_tasks" on custom_tasks for select using (true);
create policy "public insert custom_tasks" on custom_tasks for insert with check (true);
create policy "public update custom_tasks" on custom_tasks for update using (true);
create policy "public delete custom_tasks" on custom_tasks for delete using (true);

create policy "public select meetings" on meetings for select using (true);
create policy "public insert meetings" on meetings for insert with check (true);
create policy "public update meetings" on meetings for update using (true);
create policy "public delete meetings" on meetings for delete using (true);

-- ── Realtime ────────────────────────────────────────────────────────
-- Necesario para que dos usuarios vean los cambios del otro sin recargar.

alter publication supabase_realtime add table task_overrides;
alter publication supabase_realtime add table notes;
alter publication supabase_realtime add table custom_tasks;
alter publication supabase_realtime add table meetings;
