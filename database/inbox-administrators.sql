-- Applied via add_protected_inbox_administrators. Grant membership separately
-- to a verified existing agent ID; no administrator is inferred from a name.
create table public.inbox_administrators (
  agent_id uuid primary key references public.agents(id) on delete cascade,
  created_at timestamptz not null default now()
);
alter table public.inbox_administrators enable row level security;
revoke all on public.inbox_administrators from public, anon, authenticated;
grant select on public.inbox_administrators to service_role;
