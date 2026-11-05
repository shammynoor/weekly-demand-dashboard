// Real PostgreSQL execution via PGlite, or a dedicated Supabase test database.
// PGlite stays in task-owned tooling rather than adding runtime dependencies.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
const modulePath=process.env.PGLITE_MODULE;
if(!modulePath)throw Error('Set PGLITE_MODULE to the installed @electric-sql/pglite module.');
const {PGlite}=await import(pathToFileURL(modulePath));const db=new PGlite();
const sql=await fs.readFile(new URL('../supabase/migrations/202611050001_practice_mvp.sql',import.meta.url),'utf8');
await db.exec(`create role anon;create role authenticated;create schema auth;
create table auth.users(id uuid primary key,email text,email_confirmed_at timestamptz);
create function auth.uid() returns uuid language sql stable as $$select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid$$;
grant usage on schema auth to anon,authenticated;grant execute on function auth.uid() to anon,authenticated;`);
await db.exec(sql);
const ids={owner:'00000000-0000-0000-0000-000000000001',adminA:'00000000-0000-0000-0000-000000000002',viewerA:'00000000-0000-0000-0000-000000000003',viewerB:'00000000-0000-0000-0000-000000000004',pending:'00000000-0000-0000-0000-000000000005',outsider:'00000000-0000-0000-0000-000000000006'};
for(const [name,id] of Object.entries(ids))await db.query('insert into auth.users values($1,$2,now())',[id,name.toLowerCase()+'@example.test']);
await db.query('insert into public.iq_platform_admins values($1)',[ids.owner]);
async function as(name,role='authenticated'){await db.exec('reset role');await db.query("select set_config('request.jwt.claim.sub',$1,false)",[ids[name]||'']);await db.exec('set role '+role);}
async function rpc(name,args=[]){const result=await db.query(`select public.iq_${name}(${args.map((_,i)=>'$'+(i+1)).join(',')}) as result`,args);return result.rows[0].result;}
async function reject(action,pattern){await assert.rejects(action,pattern);}
try{
 await as('owner');const a=await rpc('create_practice',['Darwin test','darwin']);const b=await rpc('create_practice',['Riverside test','riverside']);
 await db.exec('reset role');
 for(const [practice,who,role,status] of [[a,'adminA','practice_admin','approved'],[a,'viewerA','viewer','approved'],[b,'viewerB','viewer','approved'],[a,'pending','viewer','pending']])await db.query('insert into iq_memberships values($1,$2,$3,$4,now())',[practice,ids[who],role,status]);
 const rows=JSON.stringify([{at:'2025-07-01T08:00:00',total:2,other:1},{at:'2025-07-02T09:00:00',total:3,other:2}]);
 await as('adminA');assert.equal((await rpc('import_rh',[a,'darwin','initial.csv',rows])).added,5);
 assert.equal((await rpc('import_rh',[a,'darwin','repeat.csv',rows])).added,0);
 assert.equal((await db.query('select record_count from iq_rh_state')).rows[0].record_count,5);
 const mixed=JSON.stringify([{at:'2025-07-02T09:00:00',total:7,other:7},{at:'2025-07-03T09:00:00',total:2,other:2}]);
 assert.deepEqual(Object.keys(await rpc('import_rh',[a,'darwin','mixed.csv',mixed])).sort(),['added','import_id','latest_request','skipped']);
 assert.equal((await db.query('select record_count from iq_rh_state')).rows[0].record_count,7);
 await reject(rpc('import_rh',[b,'riverside','wrong.csv',rows]),/administrator required/i);
 await reject(rpc('import_rh',[a,'riverside','wrong.csv',rows]),/tenant does not match/i);
 await reject(rpc('create_practice',['Unauthorised','bad']),/Platform administrator/);
 await reject(rpc('create_invite',[a,'new@example.test','practice_admin']),/Only the platform/);
 await reject(db.query("update iq_memberships set role='practice_admin' where user_id=$1",[ids.viewerA]),/permission denied/i);
 await reject(db.query('insert into iq_rh_daily values($1,$2,$3)',[a,'2025-07-04',99]),/permission denied/i);
 await reject(rpc('import_rh',[a,'darwin','bad.csv',JSON.stringify([{at:'2025-02-30T09:00:00',total:1,other:1}])]),/date|range/i);
 assert.equal((await db.query('select record_count from iq_rh_state')).rows[0].record_count,7,'failed import must roll back');
 await reject(rpc('import_rh',[a,'darwin','duplicate.csv',JSON.stringify([{at:'2025-07-04T09:00:00',total:1,other:1},{at:'2025-07-04T09:00:00',total:1,other:1}])]),/Duplicate minute/);
 const invite=await rpc('create_invite',[a,'outsider@example.test','viewer']);
 await as('viewerB');assert.equal((await db.query('select * from iq_rh_daily')).rows.length,0,'other practice data must be invisible');
 await reject(rpc('import_rh',[b,'riverside','viewer.csv',rows]),/administrator/);
 await reject(rpc('list_members',[a]),/administrator/);await reject(rpc('redeem_invite',[invite.token]),/invalid or expired/);
 await as('outsider');await rpc('redeem_invite',[invite.token]);assert.equal((await rpc('my_context')).practices[0].status,'pending');
 assert.equal((await db.query('select * from iq_rh_daily')).rows.length,0,'pending member cannot read data');
 await as('adminA');await rpc('set_member_status',[a,ids.outsider,'approved']);
 await as('outsider');assert.equal((await db.query('select * from iq_rh_daily')).rows.length,3);
 await reject(rpc('set_member_status',[a,ids.pending,'approved']),/administrator/);
 await as('adminA');await rpc('set_member_status',[a,ids.outsider,'suspended']);
 await as('outsider');await rpc('redeem_invite',[invite.token]);assert.equal((await rpc('my_context')).practices[0].status,'suspended','reusing an invite cannot restore access');
 assert.equal((await db.query('select * from iq_rh_daily')).rows.length,0,'suspension applies to an existing session');
 await as('viewerA');assert.equal((await db.query('select sum(other_enquiries) as n from iq_rh_daily')).rows[0].n,5);
 await reject(db.query("update iq_memberships set status='approved' where user_id=$1",[ids.viewerA]),/permission denied/);
 await as('pending');assert.equal((await db.query('select * from iq_rh_state')).rows.length,0);
 await as(null,'anon');await reject(rpc('my_context'),/permission denied/);await reject(db.query('select * from iq_rh_daily'),/permission denied/);
 console.log('PostgreSQL isolation tests passed: two practices, roles, approvals, revocation, invite replay, direct-table attacks, cutoff imports, and rollback.');
}finally{await db.close();}
