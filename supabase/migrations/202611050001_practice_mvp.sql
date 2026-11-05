-- Isolated MVP backend. Apply only to a new test Supabase project.
create table public.iq_platform_admins (
 user_id uuid primary key references auth.users(id) on delete cascade
);
create table public.iq_practices (
 id uuid primary key default gen_random_uuid(),
 name text not null check (length(trim(name)) between 2 and 120),
 tenant_name text not null unique check (tenant_name ~ '^[a-z0-9_-]{2,80}$'),
 active boolean not null default true,
 created_at timestamptz not null default now()
);
create table public.iq_memberships (
 practice_id uuid not null references public.iq_practices(id),
 user_id uuid not null references auth.users(id) on delete cascade,
 role text not null check (role in ('practice_admin','viewer')),
 status text not null default 'pending' check (status in ('pending','approved','suspended')),
 created_at timestamptz not null default now(),
 primary key(practice_id,user_id)
);
create index iq_memberships_user on public.iq_memberships(user_id);
create table public.iq_invites (
 id uuid primary key default gen_random_uuid(), practice_id uuid not null references public.iq_practices(id),
 email text not null, role text not null check(role in ('practice_admin','viewer')),
 token uuid not null unique default gen_random_uuid(),
 expires_at timestamptz not null default now()+interval '7 days',
 used_by uuid references auth.users(id), created_by uuid not null references auth.users(id)
);
create table public.iq_rh_imports (
 id uuid primary key default gen_random_uuid(),practice_id uuid not null references public.iq_practices(id),
 imported_by uuid not null references auth.users(id), imported_at timestamptz not null default now(),
 source_name text not null check(length(source_name)<=200),
 added integer not null check(added>=0), skipped integer not null check(skipped>=0)
);
create index iq_rh_imports_practice on public.iq_rh_imports(practice_id,imported_at desc);
create table public.iq_rh_state (
 practice_id uuid primary key references public.iq_practices(id),
 latest_request timestamp, coverage_start date, coverage_end date,
 record_count integer not null default 0, updated_at timestamptz,
 check(record_count>=0)
);
-- Only fields needed for the MVP's counts/cutoff are retained; no CSV or patient details.
create table public.iq_rh_daily (
 practice_id uuid not null references public.iq_practices(id),request_date date not null,
 other_enquiries integer not null check(other_enquiries>=0), primary key(practice_id,request_date)
);
create table public.iq_audit (
 id bigint generated always as identity primary key, practice_id uuid references public.iq_practices(id),
 actor uuid references auth.users(id), action text not null, occurred_at timestamptz not null default now()
);

create function public.iq_is_platform_admin() returns boolean
language sql stable security definer set search_path=''
as $$select exists(select 1 from public.iq_platform_admins where user_id=auth.uid())$$;

create function public.iq_can_access(p_practice uuid,p_manage boolean default false) returns boolean
language sql stable security definer set search_path=''
as $$select auth.uid() is not null and exists(
 select 1 from public.iq_practices p where p.id=p_practice and p.active and (
 public.iq_is_platform_admin() or exists(select 1 from public.iq_memberships m
 where m.practice_id=p.id and m.user_id=auth.uid() and m.status='approved'
 and (not p_manage or m.role='practice_admin'))))$$;

create function public.iq_my_context() returns jsonb
language sql stable security definer set search_path=''
as $$select jsonb_build_object('platform_admin',public.iq_is_platform_admin(),
 'practices',coalesce((select jsonb_agg(jsonb_build_object('id',p.id,'name',p.name,
 'tenant_name',p.tenant_name,'active',p.active,'role',case when public.iq_is_platform_admin() then 'platform_admin' else m.role end,
 'status',case when public.iq_is_platform_admin() then 'approved' else m.status end) order by p.name)
 from public.iq_practices p left join public.iq_memberships m on m.practice_id=p.id and m.user_id=auth.uid()
 where auth.uid() is not null and (public.iq_is_platform_admin() or m.user_id is not null)),'[]'::jsonb))$$;

create function public.iq_create_practice(p_name text,p_tenant_name text) returns uuid
language plpgsql security definer set search_path='' as $$
declare v_id uuid;
begin
 if not public.iq_is_platform_admin() then raise exception 'Platform administrator required' using errcode='42501'; end if;
 insert into public.iq_practices(name,tenant_name) values(trim(p_name),lower(trim(p_tenant_name))) returning id into v_id;
 insert into public.iq_audit(practice_id,actor,action) values(v_id,auth.uid(),'practice_created');
 return v_id;
end $$;

create function public.iq_create_invite(p_practice uuid,p_email text,p_role text default 'viewer') returns jsonb
language plpgsql security definer set search_path='' as $$
declare v_inv public.iq_invites;
begin
 if not public.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501'; end if;
 if p_role='practice_admin' and not public.iq_is_platform_admin() then raise exception 'Only the platform administrator can appoint practice administrators' using errcode='42501'; end if;
 if p_email is null or p_email !~ '^[^@[:space:]]+@[^@[:space:]]+\.[^@[:space:]]+$' then raise exception 'Enter a valid email address'; end if;
 insert into public.iq_invites(practice_id,email,role,created_by) values(p_practice,lower(trim(p_email)),p_role,auth.uid()) returning * into v_inv;
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'invitation_created');
 return jsonb_build_object('token',v_inv.token,'expires_at',v_inv.expires_at);
end $$;

create function public.iq_redeem_invite(p_token uuid) returns jsonb
language plpgsql security definer set search_path='' as $$
declare v_inv public.iq_invites;v_email text;v_verified timestamptz;
begin
 if auth.uid() is null then raise exception 'Sign in first' using errcode='42501'; end if;
 select lower(email),email_confirmed_at into v_email,v_verified from auth.users where id=auth.uid();
 if v_verified is null then raise exception 'Verify your email address first' using errcode='42501'; end if;
 select * into v_inv from public.iq_invites where token=p_token for update;
 if not found or v_inv.expires_at<now() or v_inv.email is distinct from v_email
 or (v_inv.used_by is not null and v_inv.used_by<>auth.uid()) then raise exception 'Invitation is invalid or expired' using errcode='42501'; end if;
 if not exists(select 1 from public.iq_practices where id=v_inv.practice_id and active) then raise exception 'Practice is inactive' using errcode='42501'; end if;
 -- Repeated acceptance must never reinstate a suspended account or promote a role.
 insert into public.iq_memberships(practice_id,user_id,role,status) values(v_inv.practice_id,auth.uid(),v_inv.role,'pending') on conflict do nothing;
 update public.iq_invites set used_by=auth.uid() where id=v_inv.id;
 insert into public.iq_audit(practice_id,actor,action) values(v_inv.practice_id,auth.uid(),'invitation_accepted');
 return public.iq_my_context();
end $$;

create function public.iq_list_members(p_practice uuid) returns jsonb
language plpgsql stable security definer set search_path='' as $$
begin
 if not public.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501'; end if;
 return coalesce((select jsonb_agg(jsonb_build_object('user_id',m.user_id,'email',u.email,'role',m.role,'status',m.status) order by u.email)
 from public.iq_memberships m join auth.users u on u.id=m.user_id where m.practice_id=p_practice),'[]'::jsonb);
end $$;

create function public.iq_set_member_status(p_practice uuid,p_user uuid,p_status text) returns void
language plpgsql security definer set search_path='' as $$
begin
 if not public.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501'; end if;
 if p_user=auth.uid() then raise exception 'You cannot change your own approval status'; end if;
 if not public.iq_is_platform_admin() and exists(select 1 from public.iq_memberships where practice_id=p_practice and user_id=p_user and role='practice_admin') then raise exception 'Platform administrator required for administrator changes' using errcode='42501'; end if;
 if p_status not in ('approved','suspended') then raise exception 'Invalid approval status'; end if;
 update public.iq_memberships set status=p_status where practice_id=p_practice and user_id=p_user;
 if not found then raise exception 'Member not found'; end if;
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'membership_'||p_status);
end $$;

-- One transaction, with a per-practice lock: parallel imports cannot double count.
-- Client CSVs are reduced to minute-level counts; the server owns the cutoff.
create function public.iq_import_rh(p_practice uuid,p_tenant text,p_source text,p_rows jsonb) returns jsonb
language plpgsql security definer set search_path='' as $$
declare v_cutoff timestamp;v_added integer;v_skipped integer;v_first date;v_last date;v_latest timestamp;v_import uuid;
begin
 if not public.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501'; end if;
 if not exists(select 1 from public.iq_practices where id=p_practice and tenant_name=p_tenant) then raise exception 'CSV tenant does not match the selected practice'; end if;
 if p_rows is null or jsonb_typeof(p_rows)<>'array' or jsonb_array_length(p_rows)=0 or jsonb_array_length(p_rows)>100000 then raise exception 'Invalid import batch'; end if;
 if exists(select 1 from jsonb_array_elements(p_rows) r where
  (r->>'at') is null or (r->>'at') !~ '^20[0-9]{2}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:00$'
  or (r->>'total') is null or (r->>'total') !~ '^[0-9]{1,6}$' or (r->>'total')::int<1
  or (r->>'other') is null or (r->>'other') !~ '^[0-9]{1,6}$' or (r->>'other')::int>(r->>'total')::int) then raise exception 'Invalid request timestamp or counts'; end if;
 -- Reject duplicate minute groups rather than accidentally double counting a malformed batch.
 if exists(select 1 from jsonb_array_elements(p_rows) r group by r->>'at' having count(*)>1) then raise exception 'Duplicate minute groups in import'; end if;
 insert into public.iq_rh_state(practice_id) values(p_practice) on conflict do nothing;
 select latest_request into v_cutoff from public.iq_rh_state where practice_id=p_practice for update;
 select coalesce(sum((r->>'total')::int) filter(where v_cutoff is null or (r->>'at')::timestamp>v_cutoff),0),
 coalesce(sum((r->>'total')::int) filter(where v_cutoff is not null and (r->>'at')::timestamp<=v_cutoff),0),
 min((r->>'at')::date) filter(where v_cutoff is null or (r->>'at')::timestamp>v_cutoff),
 max((r->>'at')::date) filter(where v_cutoff is null or (r->>'at')::timestamp>v_cutoff),
 max((r->>'at')::timestamp) filter(where v_cutoff is null or (r->>'at')::timestamp>v_cutoff)
 into v_added,v_skipped,v_first,v_last,v_latest from jsonb_array_elements(p_rows) r;
 if v_added=0 then return jsonb_build_object('added',0,'skipped',v_skipped,'latest_request',v_cutoff); end if;
 insert into public.iq_rh_daily(practice_id,request_date,other_enquiries)
 select p_practice,(r->>'at')::date,sum((r->>'other')::int)::int from jsonb_array_elements(p_rows) r
 where v_cutoff is null or (r->>'at')::timestamp>v_cutoff group by (r->>'at')::date
 on conflict(practice_id,request_date) do update set other_enquiries=public.iq_rh_daily.other_enquiries+excluded.other_enquiries;
 update public.iq_rh_state set latest_request=v_latest,coverage_start=least(coverage_start,v_first),coverage_end=greatest(coverage_end,v_last),record_count=record_count+v_added,updated_at=now() where practice_id=p_practice;
 insert into public.iq_rh_imports(practice_id,imported_by,source_name,added,skipped) values(p_practice,auth.uid(),p_source,v_added,v_skipped) returning id into v_import;
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'rapid_health_imported');
 return jsonb_build_object('added',v_added,'skipped',v_skipped,'latest_request',v_latest,'import_id',v_import);
end $$;

alter table public.iq_platform_admins enable row level security;
alter table public.iq_practices enable row level security;
alter table public.iq_memberships enable row level security;
alter table public.iq_invites enable row level security;
alter table public.iq_rh_imports enable row level security;
alter table public.iq_rh_state enable row level security;
alter table public.iq_rh_daily enable row level security;
alter table public.iq_audit enable row level security;
create policy iq_admin_self on public.iq_platform_admins for select to authenticated using(user_id=auth.uid());
create policy iq_practice_read on public.iq_practices for select to authenticated using(public.iq_can_access(id));
create policy iq_membership_read on public.iq_memberships for select to authenticated using(user_id=auth.uid() or public.iq_can_access(practice_id,true));
create policy iq_imports_read on public.iq_rh_imports for select to authenticated using(public.iq_can_access(practice_id));
create policy iq_state_read on public.iq_rh_state for select to authenticated using(public.iq_can_access(practice_id));
create policy iq_daily_read on public.iq_rh_daily for select to authenticated using(public.iq_can_access(practice_id));
create policy iq_audit_read on public.iq_audit for select to authenticated using(public.iq_can_access(practice_id,true));
-- No client can directly insert, update, or delete rows, including their own membership.
revoke all on public.iq_platform_admins,public.iq_practices,public.iq_memberships,public.iq_invites,public.iq_rh_imports,public.iq_rh_state,public.iq_rh_daily,public.iq_audit from anon,authenticated;
grant select on public.iq_platform_admins,public.iq_practices,public.iq_memberships,public.iq_rh_imports,public.iq_rh_state,public.iq_rh_daily,public.iq_audit to authenticated;
revoke execute on function public.iq_is_platform_admin(),public.iq_can_access(uuid,boolean),public.iq_my_context(),public.iq_create_practice(text,text),public.iq_create_invite(uuid,text,text),public.iq_redeem_invite(uuid),public.iq_list_members(uuid),public.iq_set_member_status(uuid,uuid,text),public.iq_import_rh(uuid,text,text,jsonb) from public,anon;
grant execute on function public.iq_is_platform_admin(),public.iq_can_access(uuid,boolean),public.iq_my_context(),public.iq_create_practice(text,text),public.iq_create_invite(uuid,text,text),public.iq_redeem_invite(uuid),public.iq_list_members(uuid),public.iq_set_member_status(uuid,uuid,text),public.iq_import_rh(uuid,text,text,jsonb) to authenticated;
