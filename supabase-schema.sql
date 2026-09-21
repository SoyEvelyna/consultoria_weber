-- Weber Tracker — esquema MULTI-CLIENTE de Supabase, con login real
-- Un único proyecto de Supabase, varios clientes, aislamiento real por RLS.
--
-- CÓMO FUNCIONA EL AISLAMIENTO:
-- Cada usuario que va a usar el tracker inicia sesión con su propia cuenta
-- de Supabase Auth (creada por vos, la administradora, desde Authentication
-- → Users). La tabla `profiles` liga esa cuenta a un client_id (o a
-- is_admin = true si puede ver todos los clientes, como tu propia cuenta).
-- Las políticas RLS de abajo solo dejan leer/escribir filas cuyo client_id
-- coincide con el de la cuenta logueada — o cualquier fila, si es admin.
--
-- Vos, como administradora del proyecto, además ves y gestionás TODO desde
-- el Table Editor / SQL Editor de Supabase: el dueño del proyecto siempre
-- tiene acceso completo ahí, sin pasar por estas políticas.
--
-- Este script es seguro de volver a correr aunque ya hayas ejecutado una
-- versión anterior (usa IF NOT EXISTS / OR REPLACE / DROP...CREATE en las
-- políticas).
--
-- Ejecutar en: Supabase → SQL Editor → New query → Run

-- ── Registro de clientes ────────────────────────────────────────────

create table if not exists clients (
  id text primary key,          -- slug corto, ej: 'weber'
  name text not null,           -- nombre visible, ej: 'Weber'
  created_at timestamptz not null default now()
);

-- ── Perfiles: liga cada cuenta de Supabase Auth a un cliente ────────

create table if not exists profiles (
  user_id uuid primary key references auth.users(id) on delete cascade,
  client_id text references clients(id),
  is_admin boolean not null default false
);

-- ── Función helper para las políticas ───────────────────────────────
-- security definer: puede leer `profiles` aunque esa tabla también tenga
-- RLS activado (evita recursión). Devuelve true si el usuario autenticado
-- es admin, o si su client_id coincide con el que se está pidiendo.

create or replace function has_client_access(target_client_id text)
returns boolean
language sql
security definer
stable
as $$
  select exists (
    select 1 from profiles
    where user_id = auth.uid()
      and (is_admin = true or client_id = target_client_id)
  );
$$;

-- ── Tablas de datos (todas con client_id) ───────────────────────────
-- PK compuesta (client_id, id): los ids como "t1", "c172..." se generan
-- de forma independiente en cada tracker, así que sin el client_id en la
-- clave dos clientes distintos podrían pisarse datos entre sí.

create table if not exists task_overrides (
  client_id text not null references clients(id),
  task_id text not null,
  estado text,
  link text,
  inicio text,
  cierre text,
  hidden boolean not null default false,
  updated_at timestamptz not null default now(),
  primary key (client_id, task_id)
);
-- por si ya existía la tabla de una corrida anterior sin estas columnas:
alter table task_overrides add column if not exists inicio text;
alter table task_overrides add column if not exists cierre text;
alter table task_overrides add column if not exists hidden boolean not null default false;

create table if not exists notes (
  client_id text not null references clients(id),
  id text not null,
  text text not null,
  author text,
  created_at timestamptz not null default now(),
  primary key (client_id, id)
);

create table if not exists custom_tasks (
  client_id text not null references clients(id),
  id text not null,
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
  created_at timestamptz not null default now(),
  primary key (client_id, id)
);

create table if not exists meetings (
  client_id text not null references clients(id),
  id text not null,
  fecha text,
  responsable text,
  duracion text,
  resumen text,
  created_at timestamptz not null default now(),
  primary key (client_id, id)
);

create index if not exists task_overrides_client_idx on task_overrides(client_id);
create index if not exists notes_client_idx on notes(client_id);
create index if not exists custom_tasks_client_idx on custom_tasks(client_id);
create index if not exists meetings_client_idx on meetings(client_id);

-- ── Row Level Security ──────────────────────────────────────────────

alter table clients enable row level security;
alter table profiles enable row level security;
alter table task_overrides enable row level security;
alter table notes enable row level security;
alter table custom_tasks enable row level security;
alter table meetings enable row level security;

drop policy if exists "profiles: self select" on profiles;
create policy "profiles: self select" on profiles
  for select using (user_id = auth.uid());

drop policy if exists "clients: select own or admin" on clients;
create policy "clients: select own or admin" on clients
  for select using (has_client_access(id));

drop policy if exists "task_overrides: select" on task_overrides;
create policy "task_overrides: select" on task_overrides
  for select using (has_client_access(client_id));
drop policy if exists "task_overrides: insert" on task_overrides;
create policy "task_overrides: insert" on task_overrides
  for insert with check (has_client_access(client_id));
drop policy if exists "task_overrides: update" on task_overrides;
create policy "task_overrides: update" on task_overrides
  for update using (has_client_access(client_id));

drop policy if exists "notes: select" on notes;
create policy "notes: select" on notes
  for select using (has_client_access(client_id));
drop policy if exists "notes: insert" on notes;
create policy "notes: insert" on notes
  for insert with check (has_client_access(client_id));
drop policy if exists "notes: update" on notes;
create policy "notes: update" on notes
  for update using (has_client_access(client_id));
drop policy if exists "notes: delete" on notes;
create policy "notes: delete" on notes
  for delete using (has_client_access(client_id));

drop policy if exists "custom_tasks: select" on custom_tasks;
create policy "custom_tasks: select" on custom_tasks
  for select using (has_client_access(client_id));
drop policy if exists "custom_tasks: insert" on custom_tasks;
create policy "custom_tasks: insert" on custom_tasks
  for insert with check (has_client_access(client_id));
drop policy if exists "custom_tasks: update" on custom_tasks;
create policy "custom_tasks: update" on custom_tasks
  for update using (has_client_access(client_id));
drop policy if exists "custom_tasks: delete" on custom_tasks;
create policy "custom_tasks: delete" on custom_tasks
  for delete using (has_client_access(client_id));

drop policy if exists "meetings: select" on meetings;
create policy "meetings: select" on meetings
  for select using (has_client_access(client_id));
drop policy if exists "meetings: insert" on meetings;
create policy "meetings: insert" on meetings
  for insert with check (has_client_access(client_id));
drop policy if exists "meetings: update" on meetings;
create policy "meetings: update" on meetings
  for update using (has_client_access(client_id));
drop policy if exists "meetings: delete" on meetings;
create policy "meetings: delete" on meetings
  for delete using (has_client_access(client_id));

-- ── Realtime ────────────────────────────────────────────────────────
-- (seguro de re-ejecutar: si la tabla ya está en la publicación, lo ignora)

do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname='supabase_realtime' and schemaname='public' and tablename='task_overrides'
  ) then
    alter publication supabase_realtime add table task_overrides;
  end if;
  if not exists (
    select 1 from pg_publication_tables
    where pubname='supabase_realtime' and schemaname='public' and tablename='notes'
  ) then
    alter publication supabase_realtime add table notes;
  end if;
  if not exists (
    select 1 from pg_publication_tables
    where pubname='supabase_realtime' and schemaname='public' and tablename='custom_tasks'
  ) then
    alter publication supabase_realtime add table custom_tasks;
  end if;
  if not exists (
    select 1 from pg_publication_tables
    where pubname='supabase_realtime' and schemaname='public' and tablename='meetings'
  ) then
    alter publication supabase_realtime add table meetings;
  end if;
end $$;

-- ── Primer cliente ──────────────────────────────────────────────────

insert into clients (id, name) values ('weber', 'Weber')
  on conflict (id) do nothing;
