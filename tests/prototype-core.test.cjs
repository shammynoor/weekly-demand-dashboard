const test=require('node:test'),assert=require('node:assert/strict'),rh=require('../direct-rh.js');
test('prototype CSV preparation preserves same-minute requests and counts only Admin Other enquiries',async()=>{
 const {parseUpload}=await import('../prototype/core.mjs');
 const text='tenant_name,request_date,request_time,request_type,admin_type,age_between_created_and_processed\ndarwin,01-Jul-25,08:00,Admin,OTHER_ENQUIRY,0 hour(s)\ndarwin,01-Jul-25,08:00,Medical,OTHER_ENQUIRY,0 hour(s)\ndarwin,01-Jul-25,08:01,Admin,FIT_NOTE,0 hour(s)';
 const papa={parse:s=>({data:s.split('\n').map(x=>x.split(',')),errors:[]})};
 assert.deepEqual(parseUpload(text,papa,rh),{tenant:'darwin',count:3,rows:[{at:'2025-07-01T08:00:00',total:2,other:1},{at:'2025-07-01T08:01:00',total:1,other:0}]});
 assert.throws(()=>parseUpload(text.replace('darwin,01-Jul-25,08:01','other,01-Jul-25,08:01'),papa,rh),/several tenants/);
});
test('prototype averages include zero days and partial-month unavailable weekdays',async()=>{
 const {monthlyAverages}=await import('../prototype/core.mjs');
 const july=monthlyAverages({coverage_start:'2025-07-01',coverage_end:'2025-07-31'},[{request_date:'2025-07-02',other_enquiries:16}])[0];
 assert.equal(july.averages[2],3.2);assert.deepEqual(july.days,[4,5,5,5,4,4,4]);
 assert.deepEqual(monthlyAverages({coverage_start:'2025-08-01',coverage_end:'2025-08-01'},[])[0].averages,[null,null,null,null,0,null,null]);
});
test('browser configuration rejects server secrets and unencrypted backend connections',async()=>{
 const {validateConfig,canManage}=await import('../prototype/core.mjs');
 assert.equal(validateConfig({}),false);assert.throws(()=>validateConfig({supabaseUrl:'https://test.supabase.co',publishableKey:'sb_secret_never'}));
 assert.throws(()=>validateConfig({supabaseUrl:'http://test.supabase.co',publishableKey:'sb_publishable_test'}));
 assert.equal(canManage({active:true,status:'pending',role:'practice_admin'}),false);assert.equal(canManage({active:true,status:'approved',role:'viewer'}),false);
});
test('demo walkthrough isolates practices and supports approval and suspension without persistent storage',async()=>{
 const {DemoBackend}=await import('../prototype/demo.mjs');const demo=new DemoBackend();
 await demo.setStatus('darwin-demo','pending','approved');demo.identity='pending';assert.equal((await demo.context()).practices[0].status,'approved');
 assert.equal((await demo.report('darwin-demo')).state.record_count,1400);await assert.rejects(demo.report('riverside-demo'),/no access/);
 demo.identity='admin';await demo.setStatus('darwin-demo','pending','suspended');demo.identity='pending';await assert.rejects(demo.report('darwin-demo'),/no access/);
 demo.identity='viewer';await assert.rejects(demo.importRH('darwin-demo',{tenant:'darwin',rows:[]},'test.csv'),/no access/);
});
