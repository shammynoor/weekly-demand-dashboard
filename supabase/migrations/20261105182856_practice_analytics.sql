-- Practice-owned compressed aggregate snapshots. Original CSVs are not retained.
create table public.iq_analytics_sources (
 practice_id uuid not null references public.iq_practices(id) on delete cascade,
 source_key text not null check(length(source_key)=64),
 kind text not null check(kind in ('slot','signposting')),
 source_name text not null check(length(source_name) between 1 and 200),
 clinical_at timestamptz not null,
 imported_at timestamptz not null default now(),
 imported_by uuid not null references auth.users(id),
 row_count integer not null check(row_count>=0),
 payload text not null check(length(payload) between 1 and 12000000),
 revision bigint not null default 1,
 primary key(practice_id,source_key)
);
create index iq_analytics_sources_actor on public.iq_analytics_sources(imported_by);
create table public.iq_analytics_settings (
 practice_id uuid primary key references public.iq_practices(id) on delete cascade,
 settings jsonb not null default '{}'::jsonb check(jsonb_typeof(settings)='object'),
 revision bigint not null default 1,
 updated_at timestamptz not null default now(),
 updated_by uuid not null references auth.users(id)
);
create index iq_analytics_settings_actor on public.iq_analytics_settings(updated_by);
alter table public.iq_analytics_sources enable row level security;
alter table public.iq_analytics_settings enable row level security;
create policy iq_sources_read on public.iq_analytics_sources for select to authenticated using(iq_private.iq_can_access(practice_id));
create policy iq_settings_read on public.iq_analytics_settings for select to authenticated using(iq_private.iq_can_access(practice_id));
revoke all on public.iq_analytics_sources,public.iq_analytics_settings from anon,authenticated;
grant select on public.iq_analytics_sources,public.iq_analytics_settings to authenticated;

create function iq_private.iq_save_analytics_source(p_practice uuid,p_key text,p_kind text,p_name text,p_clinical timestamptz,p_count integer,p_payload text,p_expected bigint) returns bigint
language plpgsql security definer set search_path='' as $$
declare current_revision bigint; next_revision bigint;
begin
 if not iq_private.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501';end if;
 -- Per-practice lock makes first creation and refresh atomic, including parallel uploads.
 perform 1 from public.iq_practices where id=p_practice for update;
 select revision into current_revision from public.iq_analytics_sources where practice_id=p_practice and source_key=p_key;
 if coalesce(current_revision,0) is distinct from p_expected then raise exception 'Report changed. Refresh before uploading again.' using errcode='40001';end if;
 if p_payload !~ '^[A-Za-z0-9+/]+={0,2}$' then raise exception 'Invalid compressed report';end if;
 if p_clinical is null or p_clinical>'2100-01-01' or p_clinical<'2000-01-01' then raise exception 'Invalid report date';end if;
 next_revision:=coalesce(current_revision,0)+1;
 insert into public.iq_analytics_sources(practice_id,source_key,kind,source_name,clinical_at,row_count,payload,imported_by,revision)
 values(p_practice,p_key,p_kind,trim(p_name),p_clinical,p_count,p_payload,auth.uid(),next_revision)
 on conflict(practice_id,source_key) do update set kind=excluded.kind,source_name=excluded.source_name,clinical_at=excluded.clinical_at,row_count=excluded.row_count,payload=excluded.payload,imported_by=excluded.imported_by,revision=excluded.revision,imported_at=now();
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'analytics_report_saved');
 return next_revision;
end $$;
create function iq_private.iq_delete_analytics_source(p_practice uuid,p_key text,p_expected bigint) returns void
language plpgsql security definer set search_path='' as $$
begin
 if not iq_private.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501';end if;
 delete from public.iq_analytics_sources where practice_id=p_practice and source_key=p_key and revision=p_expected;
 if not found then raise exception 'Report changed. Refresh before removing it.' using errcode='40001';end if;
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'analytics_report_removed');
end $$;
create function iq_private.iq_save_analytics_settings(p_practice uuid,p_settings jsonb,p_expected bigint) returns bigint
language plpgsql security definer set search_path='' as $$
declare current_revision bigint; next_revision bigint;
begin
 if not iq_private.iq_can_access(p_practice,true) then raise exception 'Practice administrator required' using errcode='42501';end if;
 if jsonb_typeof(p_settings) is distinct from 'object' or octet_length(p_settings::text)>2000000 then raise exception 'Invalid practice settings';end if;
 perform 1 from public.iq_practices where id=p_practice for update;
 select revision into current_revision from public.iq_analytics_settings where practice_id=p_practice;
 if coalesce(current_revision,0) is distinct from p_expected then raise exception 'Settings changed in another session. Reload before editing.' using errcode='40001';end if;
 next_revision:=coalesce(current_revision,0)+1;
 insert into public.iq_analytics_settings(practice_id,settings,updated_by,revision) values(p_practice,p_settings,auth.uid(),next_revision)
 on conflict(practice_id) do update set settings=excluded.settings,updated_by=excluded.updated_by,revision=excluded.revision,updated_at=now();
 return next_revision;
end $$;
revoke all on function iq_private.iq_save_analytics_source(uuid,text,text,text,timestamptz,integer,text,bigint),iq_private.iq_delete_analytics_source(uuid,text,bigint),iq_private.iq_save_analytics_settings(uuid,jsonb,bigint) from public,anon;
grant execute on function iq_private.iq_save_analytics_source(uuid,text,text,text,timestamptz,integer,text,bigint),iq_private.iq_delete_analytics_source(uuid,text,bigint),iq_private.iq_save_analytics_settings(uuid,jsonb,bigint) to authenticated;
create function public.iq_save_analytics_source(p_practice uuid,p_key text,p_kind text,p_name text,p_clinical timestamptz,p_count integer,p_payload text,p_expected bigint) returns bigint language sql security invoker set search_path='' as $$select iq_private.iq_save_analytics_source($1,$2,$3,$4,$5,$6,$7,$8)$$;
create function public.iq_delete_analytics_source(p_practice uuid,p_key text,p_expected bigint) returns void language sql security invoker set search_path='' as $$select iq_private.iq_delete_analytics_source($1,$2,$3)$$;
create function public.iq_save_analytics_settings(p_practice uuid,p_settings jsonb,p_expected bigint) returns bigint language sql security invoker set search_path='' as $$select iq_private.iq_save_analytics_settings($1,$2,$3)$$;
revoke all on function public.iq_save_analytics_source(uuid,text,text,text,timestamptz,integer,text,bigint),public.iq_delete_analytics_source(uuid,text,bigint),public.iq_save_analytics_settings(uuid,jsonb,bigint) from public,anon;
grant execute on function public.iq_save_analytics_source(uuid,text,text,text,timestamptz,integer,text,bigint),public.iq_delete_analytics_source(uuid,text,bigint),public.iq_save_analytics_settings(uuid,jsonb,bigint) to authenticated;
