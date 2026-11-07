-- Test the actual policies and RPCs with disposable fixtures.
begin;
select set_config('iq.test.platform',gen_random_uuid()::text,true),
       set_config('iq.test.admin',gen_random_uuid()::text,true),
       set_config('iq.test.viewer',gen_random_uuid()::text,true);
insert into auth.users(id,email,email_confirmed_at) values
 (current_setting('iq.test.platform')::uuid,'platform-isolation@example.test',now()),
 (current_setting('iq.test.admin')::uuid,'admin-isolation@example.test',now()),
 (current_setting('iq.test.viewer')::uuid,'viewer-isolation@example.test',now());
insert into public.iq_platform_admins values(current_setting('iq.test.platform')::uuid);

set local role authenticated;
select set_config('request.jwt.claim.sub',current_setting('iq.test.platform'),true);
select set_config('iq.test.practice',public.iq_create_practice('Role isolation fixture','role_isolation_fixture')::text,true);
reset role;
insert into public.iq_memberships(practice_id,user_id,role,status) values
 (current_setting('iq.test.practice')::uuid,current_setting('iq.test.platform')::uuid,'practice_admin','approved'),
 (current_setting('iq.test.practice')::uuid,current_setting('iq.test.admin')::uuid,'practice_admin','pending'),
 (current_setting('iq.test.practice')::uuid,current_setting('iq.test.viewer')::uuid,'viewer','approved');
insert into public.iq_rh_state(practice_id,record_count) values(current_setting('iq.test.practice')::uuid,1);
insert into public.iq_rh_daily values(current_setting('iq.test.practice')::uuid,'2026-11-07',1);
insert into public.iq_rh_imports(practice_id,imported_by,source_name,added,skipped)
 values(current_setting('iq.test.practice')::uuid,current_setting('iq.test.admin')::uuid,'fixture.csv',1,0);
insert into public.iq_analytics_settings(practice_id,settings,updated_by)
 values(current_setting('iq.test.practice')::uuid,'{}',current_setting('iq.test.admin')::uuid);
insert into public.iq_analytics_sources(practice_id,source_key,kind,source_name,clinical_at,row_count,payload,imported_by)
 values(current_setting('iq.test.practice')::uuid,repeat('a',64),'slot','fixture.csv',now(),1,'H4sIAAAA',current_setting('iq.test.admin')::uuid);

set local role authenticated;
do $$
declare p uuid := current_setting('iq.test.practice')::uuid;
begin
 if not exists(select 1 from public.iq_practices where id=p)
    or not (public.iq_my_context()->>'platform_admin')::boolean then
   raise exception 'Platform directory unavailable';
 end if;
 if exists(select 1 from public.iq_memberships)
    or exists(select 1 from public.iq_rh_state)
    or exists(select 1 from public.iq_rh_daily)
    or exists(select 1 from public.iq_rh_imports)
    or exists(select 1 from public.iq_analytics_settings)
    or exists(select 1 from public.iq_analytics_sources)
    or exists(select 1 from public.iq_audit where action<>'invitation_created') then
   raise exception 'Platform account could read practice data';
 end if;
 begin
   perform public.iq_save_analytics_settings(p,'{}',1);
   raise exception 'Platform account could write practice settings';
 exception when insufficient_privilege then null;
 end;
 begin
   perform public.iq_import_rh(p,'role_isolation_fixture','fixture.csv','[]');
   raise exception 'Platform account could import practice data';
 exception when insufficient_privilege then null;
 end;
 begin
   perform public.iq_set_member_status(p,current_setting('iq.test.viewer')::uuid,'suspended');
   raise exception 'Platform account could manage a viewer';
 exception when insufficient_privilege then null;
 end;
 if exists(select 1 from jsonb_array_elements(public.iq_list_members(p)) member where member->>'role'<>'practice_admin') then
   raise exception 'Platform account could list practice viewers';
 end if;
 perform public.iq_set_member_status(p,current_setting('iq.test.admin')::uuid,'approved');
 perform public.iq_create_invite(p,'new-admin@example.test','practice_admin');
 if not exists(select 1 from public.iq_audit where action='invitation_created') then
   raise exception 'Invitation rate limiting cannot see the caller’s event';
 end if;
end $$;

select set_config('request.jwt.claim.sub',current_setting('iq.test.admin'),true);
do $$
declare p uuid := current_setting('iq.test.practice')::uuid;
begin
 if not exists(select 1 from public.iq_analytics_sources where practice_id=p)
    or not exists(select 1 from public.iq_rh_daily where practice_id=p) then
   raise exception 'Approved practice administrator lost access';
 end if;
 perform public.iq_set_member_status(p,current_setting('iq.test.viewer')::uuid,'suspended');
 perform public.iq_create_invite(p,'new-viewer@example.test','viewer');
end $$;
rollback;
select 'Platform isolation and practice administrator onboarding passed' as result;
