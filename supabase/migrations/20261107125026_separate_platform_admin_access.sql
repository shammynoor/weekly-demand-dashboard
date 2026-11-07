-- The platform directory is administrative metadata. Practice access always
-- requires a separate account with an approved membership.
create or replace function iq_private.iq_can_access(p_practice uuid, p_manage boolean default false)
returns boolean language sql stable security definer set search_path='' as $$
 select auth.uid() is not null
   and not iq_private.iq_is_platform_admin()
   and exists (
     select 1
     from public.iq_practices p
     join public.iq_memberships m on m.practice_id=p.id
     where p.id=p_practice and p.active
       and m.user_id=auth.uid() and m.status='approved'
       and (not p_manage or m.role='practice_admin')
   )
$$;

create or replace function iq_private.iq_my_context()
returns jsonb language sql stable security definer set search_path='' as $$
 select jsonb_build_object(
   'platform_admin', iq_private.iq_is_platform_admin(),
   'practices', coalesce((
     select jsonb_agg(jsonb_build_object(
       'id',p.id,'name',p.name,'tenant_name',p.tenant_name,'active',p.active,
       'role',case when iq_private.iq_is_platform_admin() then 'platform_admin' else m.role end,
       'status',case when iq_private.iq_is_platform_admin() then null else m.status end
     ) order by p.name)
     from public.iq_practices p
     left join public.iq_memberships m on m.practice_id=p.id and m.user_id=auth.uid()
     where auth.uid() is not null and (iq_private.iq_is_platform_admin() or m.user_id is not null)
   ), '[]'::jsonb)
 )
$$;

alter policy iq_practice_read on public.iq_practices
using (iq_private.iq_can_access(id) or (select iq_private.iq_is_platform_admin()));

-- Restrictive so existing self-membership policies cannot give a platform
-- account access, even if that account has a historical membership.
create policy iq_membership_exclude_platform on public.iq_memberships
as restrictive for all to authenticated
using (not (select iq_private.iq_is_platform_admin()))
with check (not (select iq_private.iq_is_platform_admin()));

-- The invitation sender counts the caller's own invitation events.
alter policy iq_audit_read on public.iq_audit using (
 iq_private.iq_can_access(practice_id,true)
 or ((select iq_private.iq_is_platform_admin())
     and actor=(select auth.uid()) and action='invitation_created')
);

create or replace function iq_private.iq_create_invite(p_practice uuid,p_email text,p_role text default 'viewer')
returns jsonb language plpgsql security definer set search_path='' as $$
declare v_inv public.iq_invites;
begin
 if iq_private.iq_is_platform_admin() then
   if p_role is distinct from 'practice_admin' or not exists (
     select 1 from public.iq_practices where id=p_practice and active
   ) then
     raise exception 'Platform administrators can only invite administrators to active practices' using errcode='42501';
   end if;
 elsif not iq_private.iq_can_access(p_practice,true) or p_role is distinct from 'viewer' then
   raise exception 'Practice administrator required to invite viewers' using errcode='42501';
 end if;
 if p_email is null or p_email !~ '^[^@[:space:]]+@[^@[:space:]]+\.[^@[:space:]]+$' then
   raise exception 'Enter a valid email address';
 end if;
 insert into public.iq_invites(practice_id,email,role,created_by)
 values(p_practice,lower(trim(p_email)),p_role,auth.uid()) returning * into v_inv;
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'invitation_created');
 return jsonb_build_object('token',v_inv.token,'expires_at',v_inv.expires_at);
end $$;

create or replace function iq_private.iq_list_members(p_practice uuid)
returns jsonb language plpgsql stable security definer set search_path='' as $$
declare platform boolean := iq_private.iq_is_platform_admin();
begin
 if not platform and not iq_private.iq_can_access(p_practice,true) then
   raise exception 'Practice administrator required' using errcode='42501';
 end if;
 return coalesce((
   select jsonb_agg(jsonb_build_object('user_id',m.user_id,'email',u.email,'role',m.role,'status',m.status) order by u.email)
   from public.iq_memberships m join auth.users u on u.id=m.user_id
   where m.practice_id=p_practice and (not platform or m.role='practice_admin')
 ),'[]'::jsonb);
end $$;

create or replace function iq_private.iq_set_member_status(p_practice uuid,p_user uuid,p_status text)
returns void language plpgsql security definer set search_path='' as $$
declare platform boolean := iq_private.iq_is_platform_admin();
begin
 if platform then
   if not exists (select 1 from public.iq_practices where id=p_practice and active) then
     raise exception 'Active practice required' using errcode='42501';
   end if;
 elsif not iq_private.iq_can_access(p_practice,true) then
   raise exception 'Practice administrator required' using errcode='42501';
 end if;
 if p_user=auth.uid() then raise exception 'You cannot change your own approval status'; end if;
 if p_status is null or p_status not in ('approved','suspended') then raise exception 'Invalid approval status'; end if;
 update public.iq_memberships set status=p_status
 where practice_id=p_practice and user_id=p_user
   and role=case when platform then 'practice_admin' else 'viewer' end;
 if not found then raise exception 'You cannot manage this membership' using errcode='42501'; end if;
 insert into public.iq_audit(practice_id,actor,action) values(p_practice,auth.uid(),'membership_'||p_status);
end $$;
