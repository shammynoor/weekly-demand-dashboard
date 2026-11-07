-- Hosted analytics verification: only synthetic fixtures, rolled back completely.
begin;
select set_config('iq.test.admin',gen_random_uuid()::text,true),set_config('iq.test.owner',gen_random_uuid()::text,true),set_config('iq.test.viewer',gen_random_uuid()::text,true),set_config('iq.test.other',gen_random_uuid()::text,true);
insert into auth.users(id,email,email_confirmed_at) values
 (current_setting('iq.test.admin')::uuid,'hosted_analytics.sql-admin@example.test',now()),
(current_setting('iq.test.owner')::uuid,'analytics-owner@example.test',now()),(current_setting('iq.test.viewer')::uuid,'analytics-viewer@example.test',now()),(current_setting('iq.test.other')::uuid,'analytics-other@example.test',now());
insert into public.iq_platform_admins values(current_setting('iq.test.owner')::uuid);
set local role authenticated;
select set_config('request.jwt.claim.sub',current_setting('iq.test.owner'),true);
select set_config('iq.test.a',public.iq_create_practice('Analytics test A','analytics_test_a')::text,true),set_config('iq.test.b',public.iq_create_practice('Analytics test B','analytics_test_b')::text,true);
reset role;
insert into public.iq_memberships(practice_id,user_id,role,status) values
 (current_setting('iq.test.a')::uuid,current_setting('iq.test.admin')::uuid,'practice_admin','approved'),(current_setting('iq.test.a')::uuid,current_setting('iq.test.viewer')::uuid,'viewer','approved'),(current_setting('iq.test.b')::uuid,current_setting('iq.test.other')::uuid,'viewer','approved');
set local role authenticated;
select set_config('request.jwt.claim.sub',current_setting('iq.test.admin'),true);
select public.iq_save_analytics_source(current_setting('iq.test.a')::uuid,repeat('a',64),'slot','synthetic.csv','2026-11-01T10:00:00Z',12,'H4sIAAAA',0);
select public.iq_save_analytics_source(current_setting('iq.test.a')::uuid,repeat('a',64),'slot','synthetic.csv','2026-11-01T10:00:00Z',12,'H4sIAAAA',1);
select public.iq_save_analytics_settings(current_setting('iq.test.a')::uuid,'{"version":9}',0);
do $$begin
 if (select count(*) from public.iq_analytics_sources where practice_id=current_setting('iq.test.a')::uuid)<>1 then raise exception 'Duplicate source';end if;
 begin
  perform public.iq_save_analytics_settings(current_setting('iq.test.a')::uuid,'{}',0);raise exception 'Stale settings write accepted';
 exception when serialization_failure then null;end;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.viewer'),true);
do $$begin
 if (select count(*) from public.iq_analytics_sources)<>1 or (select count(*) from public.iq_analytics_settings)<>1 then raise exception 'Viewer cannot read own practice';end if;
 begin perform public.iq_save_analytics_settings(current_setting('iq.test.a')::uuid,'{}',1);raise exception 'Viewer could save';exception when insufficient_privilege then null;end;
 begin delete from public.iq_analytics_sources;raise exception 'Direct writes permitted';exception when insufficient_privilege then null;end;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.other'),true);
do $$begin
 if exists(select 1 from public.iq_analytics_sources) or exists(select 1 from public.iq_analytics_settings) then raise exception 'Cross-practice data leak';end if;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.admin'),true);
select public.iq_set_member_status(current_setting('iq.test.a')::uuid,current_setting('iq.test.viewer')::uuid,'suspended');
select set_config('request.jwt.claim.sub',current_setting('iq.test.viewer'),true);
do $$begin if exists(select 1 from public.iq_analytics_sources) or exists(select 1 from public.iq_analytics_settings) then raise exception 'Suspension did not revoke reads';end if;end $$;
rollback;
select 'Hosted analytics isolation, write permissions, replacement, settings conflicts and suspension passed; all fixtures rolled back.' as result;
