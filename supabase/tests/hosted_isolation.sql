-- Transactional verification on the dedicated test project. Everything rolls back.
begin;
select set_config('iq.test.owner',gen_random_uuid()::text,true),
 set_config('iq.test.viewer',gen_random_uuid()::text,true),
 set_config('iq.test.other',gen_random_uuid()::text,true);
insert into auth.users(id,email,email_confirmed_at) values
 (current_setting('iq.test.owner')::uuid,'iq-owner@example.test',now()),
 (current_setting('iq.test.viewer')::uuid,'iq-viewer@example.test',now()),
 (current_setting('iq.test.other')::uuid,'iq-other@example.test',now());
insert into public.iq_platform_admins(user_id) values(current_setting('iq.test.owner')::uuid);
set local role authenticated;
select set_config('request.jwt.claim.sub',current_setting('iq.test.owner'),true);
select set_config('iq.test.a',public.iq_create_practice('Isolated test A','isolation_a')::text,true),
 set_config('iq.test.b',public.iq_create_practice('Isolated test B','isolation_b')::text,true);
reset role;
insert into public.iq_memberships(practice_id,user_id,role,status) values
 (current_setting('iq.test.a')::uuid,current_setting('iq.test.viewer')::uuid,'viewer','pending'),
 (current_setting('iq.test.b')::uuid,current_setting('iq.test.other')::uuid,'viewer','approved');
set local role authenticated;
select set_config('request.jwt.claim.sub',current_setting('iq.test.owner'),true);
do $$declare result jsonb;begin
 result:=public.iq_import_rh(current_setting('iq.test.a')::uuid,'isolation_a','synthetic.csv','[{"at":"2026-10-01T09:00:00","total":2,"other":1}]');
 if (result->>'added')::int<>2 then raise exception 'Initial import failed';end if;
 result:=public.iq_import_rh(current_setting('iq.test.a')::uuid,'isolation_a','repeated.csv','[{"at":"2026-10-01T09:00:00","total":2,"other":1}]');
 if (result->>'added')::int<>0 then raise exception 'Repeated import counted twice';end if;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.viewer'),true);
do $$begin
 if exists(select 1 from public.iq_rh_daily) then raise exception 'Pending user could read data';end if;
 begin
  perform public.iq_import_rh(current_setting('iq.test.a')::uuid,'isolation_a','bad.csv','[{"at":"2026-10-02T09:00:00","total":1,"other":1}]');
  raise exception 'Viewer could import';
 exception when insufficient_privilege then null;end;
 begin
  update public.iq_memberships set status='approved' where user_id=auth.uid();
  raise exception 'Viewer could approve themselves';
 exception when insufficient_privilege then null;end;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.owner'),true);
select public.iq_set_member_status(current_setting('iq.test.a')::uuid,current_setting('iq.test.viewer')::uuid,'approved');
select set_config('request.jwt.claim.sub',current_setting('iq.test.viewer'),true);
do $$begin
 if (select count(*) from public.iq_rh_daily)<>1 then raise exception 'Approved member cannot read own data';end if;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.other'),true);
do $$begin
 if exists(select 1 from public.iq_rh_daily) then raise exception 'Cross-practice data leak';end if;
end $$;
select set_config('request.jwt.claim.sub',current_setting('iq.test.owner'),true);
select public.iq_set_member_status(current_setting('iq.test.a')::uuid,current_setting('iq.test.viewer')::uuid,'suspended');
select set_config('request.jwt.claim.sub',current_setting('iq.test.viewer'),true);
do $$begin
 if exists(select 1 from public.iq_rh_daily) then raise exception 'Suspended session still has data access';end if;
 begin
  perform public.iq_claim_platform_admin();raise exception 'Unauthorised administrator claim';
 exception when insufficient_privilege then null;end;
end $$;
rollback;
select 'Hosted isolation, approval, suspension, import replay and self-promotion tests passed; all fixtures rolled back.' as result;
