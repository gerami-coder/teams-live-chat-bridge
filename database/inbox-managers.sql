-- Applied via add_inbox_managers_and_website_assignment_management
create table public.inbox_managers(agent_id uuid primary key references public.agents(id) on delete cascade, created_at timestamptz not null default now());
alter table public.inbox_managers enable row level security;
revoke all on public.inbox_managers from public, anon, authenticated;
grant select, insert, delete on public.inbox_managers to service_role;
alter table public.sites add column enforce_agent_assignments boolean not null default false;
CREATE OR REPLACE FUNCTION public.set_website_agents(p_site_id uuid, p_agent_ids uuid[], p_expected_ids uuid[])
 RETURNS void
 LANGUAGE plpgsql
 SET search_path TO ''
AS $function$ declare current_ids uuid[]; begin perform 1 from public.sites where id=p_site_id and is_active for update; if not found then raise exception 'site_not_found'; end if; select coalesce(array_agg(agent_id order by agent_id),'{}'::uuid[]) into current_ids from public.site_agents where site_id=p_site_id and department='general' and is_active; if current_ids is distinct from array(select distinct x from unnest(p_expected_ids) x order by x) then raise exception 'assignments_changed' using errcode='40001'; end if; if exists(select 1 from unnest(p_agent_ids) x where not exists(select 1 from public.agents a where a.id=x and a.is_active and btrim(a.display_name)<>'' and lower(btrim(a.display_name)) not in ('support agent','agent','support','unknown'))) then raise exception 'invalid_agent'; end if; delete from public.site_agents where site_id=p_site_id and department='general'; insert into public.site_agents(site_id,agent_id,department,priority,is_active) select p_site_id,x,'general',row_number() over(order by x),true from (select distinct unnest(p_agent_ids) x) ids; update public.sites set enforce_agent_assignments=true where id=p_site_id; end; $function$

revoke all on function public.set_website_agents(uuid,uuid[],uuid[]) from public,anon,authenticated;
grant execute on function public.set_website_agents(uuid,uuid[],uuid[]) to service_role;
