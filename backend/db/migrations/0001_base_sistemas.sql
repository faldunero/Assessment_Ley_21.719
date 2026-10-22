-- Fase 1: organizaciones, pertenencia y sistemas/proveedores.
-- Todas las tablas con datos de cliente tienen Row Level Security activo.
-- Compatible con PostgreSQL (Supabase). No contiene datos reales.

create extension if not exists pgcrypto;

create table if not exists empresa (
  id uuid primary key default gen_random_uuid(),
  nombre text not null check (length(nombre) between 2 and 200),
  rut text not null unique,
  estado text not null default 'pendiente'
    check (estado in ('pendiente','autorizada','revocada')),
  creada_en timestamptz not null default now()
);

create table if not exists miembro_empresa (
  empresa_id uuid not null references empresa(id) on delete cascade,
  usuario_id uuid not null,               -- id de Supabase Auth (auth.users)
  rol text not null check (rol in ('administrador','evaluador','entrevistador','lectura')),
  primary key (empresa_id, usuario_id)
);

create table if not exists sistema_proveedor (
  id uuid primary key default gen_random_uuid(),
  empresa_id uuid not null references empresa(id) on delete cascade,
  nombre text not null check (length(nombre) between 1 and 150),
  categoria text not null check (categoria in (
    'SaaS (nube multiempresa)','Licencia instalada','Desarrollo a medida',
    'Infraestructura en nube (IaaS o PaaS)','Herramienta de productividad',
    'Pagos','Videovigilancia o biometría','Otro')),
  rol text not null check (rol in ('Responsable','Encargado','Responsable y encargado')),
  tipo_usuario text not null check (tipo_usuario in ('Nominativo','Genérico','Mixto','Sin identificar')),
  datos_clientes boolean not null default false,
  datos_sensibles boolean not null default false,
  ubicacion text not null check (ubicacion in ('Chile','Fuera de Chile','Chile y fuera de Chile','Por definir')),
  contrato_encargo text not null default 'No' check (contrato_encargo in ('Sí','No','En revisión')),
  mfa boolean not null default false,
  cifrado boolean not null default false,
  cuenta_individual boolean not null default false,
  datos_reales_pruebas boolean not null default false,
  responsable text,
  cuenta text,
  eliminado_en timestamptz,               -- borrado lógico
  actualizado_en timestamptz not null default now(),
  constraint encargado_requiere_contrato check (
    rol = 'Responsable' or contrato_encargo in ('Sí','No','En revisión')),
  constraint sin_identificar_requiere_cuenta check (
    tipo_usuario <> 'Sin identificar' or coalesce(responsable,'') <> '' or coalesce(cuenta,'') <> '')
);

create table if not exists auditoria (
  id bigint generated always as identity primary key,
  empresa_id uuid not null references empresa(id) on delete cascade,
  usuario_id uuid not null,
  accion text not null,
  entidad text not null,
  entidad_id uuid,
  antes jsonb,
  despues jsonb,
  ocurrido_en timestamptz not null default now()
);

-- ---------- Seguridad por fila ----------
-- Un usuario solo ve y modifica datos de las empresas donde es miembro.
create or replace function es_miembro(eid uuid) returns boolean
  language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from miembro_empresa m
    where m.empresa_id = eid and m.usuario_id = auth.uid()
  );
$$;

create or replace function es_admin_o_evaluador(eid uuid) returns boolean
  language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from miembro_empresa m
    where m.empresa_id = eid and m.usuario_id = auth.uid()
      and m.rol in ('administrador','evaluador')
  );
$$;

alter table empresa enable row level security;
alter table miembro_empresa enable row level security;
alter table sistema_proveedor enable row level security;
alter table auditoria enable row level security;

create policy empresa_lectura on empresa for select using (es_miembro(id));
create policy miembro_lectura on miembro_empresa for select using (es_miembro(empresa_id));
create policy sistema_lectura on sistema_proveedor for select using (es_miembro(empresa_id));
create policy sistema_escritura on sistema_proveedor for insert with check (es_admin_o_evaluador(empresa_id));
create policy sistema_actualiza on sistema_proveedor for update
  using (es_admin_o_evaluador(empresa_id)) with check (es_admin_o_evaluador(empresa_id));
create policy auditoria_lectura on auditoria for select using (es_miembro(empresa_id));
-- Sin políticas de escritura sobre auditoria: solo el servidor, con rol de servicio.
