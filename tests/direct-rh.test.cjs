const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const rh=require('../direct-rh.js');
const fields=['tenant_name','request_date','request_time','request_type','admin_type','age_between_created_and_processed'];
const row=(date,time,type='OTHER_ENQUIRY')=>['darwin',date,time,'Admin',type,'2 hour(s) 0 minute(s)'];
const table=rows=>({fields,rows});
test('overlapping full and incremental exports append only strictly newer minutes, in order',()=>{
 const old=table([row('30-Jun-25','23:59'),row('01-Jul-25','08:01')]);
 const incoming=table([row('02-Jul-25','07:00'),row('01-Jul-25','08:01','FIT_NOTE'),row('30-Jun-25','23:59'),row('01-Jul-25','08:02'),row('01-Jul-25','08:02','FIT_NOTE')]);
 const plan=rh.plan(old,incoming);
 assert.equal(plan.skipped,2);assert.equal(plan.rows.length,3);
 assert.deepEqual(plan.rows.map(r=>r.slice(1,3)),[['01-Jul-25','08:02'],['01-Jul-25','08:02'],['02-Jul-25','07:00']]);
 assert.equal(rh.plan(table([...old.rows,...plan.rows]),incoming).rows.length,0);
 assert.equal(old.rows.length,2);
});
test('cutoff is maximum timestamp, independent of original order and browser timezone',()=>{
 const old=table([row('01-Jan-26','00:00'),row('31-Dec-25','23:59')]);
 assert.equal(rh.plan(old,table([row('01-Jan-26','00:01')])).rows.length,1);
 assert.equal(rh.timestamp('29-Oct-26','01:30'),Date.UTC(2026,9,29,1,30));
});
test('schema reorder is mapped correctly; invalid dates, schema changes, and other tenants are rejected',()=>{
 const old=table([row('01-Jul-25','08:00')]),newRow=row('01-Jul-25','08:01');
 assert.deepEqual(rh.plan(old,{fields:[...fields].reverse(),rows:[[...newRow].reverse()]}).rows,[newRow]);
 for(const date of ['31-Jun-25','29-Feb-25','01-Abc-25'])assert.throws(()=>rh.plan(old,table([row(date,'08:01')])));
 assert.throws(()=>rh.plan(old,table([row('01-Jul-25','24:00')])));
 assert.throws(()=>rh.plan(old,{fields:[...fields,'extra'],rows:[[...newRow,'x']]}));
 const other=[...newRow];other[0]='different';assert.throws(()=>rh.plan(old,table([other])));
});
// A small parser stand-in isolates upload decisions; production uses the existing Papa Parse library.
const papa=process.env.DIRECT_RH_PAPA?require(process.env.DIRECT_RH_PAPA):{parse(text,options={}){const data=text.replace(/^\uFEFF/,'').trimEnd().split(/\r?\n/).filter(Boolean).map(s=>s.split(','));return {data:options.preview?data.slice(0,options.preview):data,errors:[]}},unparse(rows,{newline}){return rows.map(r=>r.join(',')).join(newline)}};
const csv=rows=>fields.join(',')+'\r\n'+rows.map(r=>r.join(',')).join('\r\n');
test('merge preserves original content and appends one headerless block; no-op returns original bytes',()=>{
 const original='\uFEFF'+csv([row('01-Jul-25','08:00')]);
 const incoming=csv([row('01-Jul-25','08:00'),row('01-Jul-25','08:01')]);
 const result=rh.merge(original,incoming,papa);
 assert.equal(result.text,original+'\r\n'+row('01-Jul-25','08:01').join(',')+'\r\n');
 assert.equal(rh.merge(result.text,incoming,papa).text,result.text);
 assert.equal(rh.recognises(incoming,papa),true);assert.equal(rh.recognises('Appointment Date,Slot Type\nx,y',papa),false);
 assert.throws(()=>rh.merge(original,incoming+'\r\nbad,row',papa));
});
function fixture({cancel=false,status=200,etag='version-1',original=csv([row('01-Jul-25','08:00')]),incoming=csv([row('01-Jul-25','08:00'),row('01-Jul-25','08:01')]),owner=true,missing=false}={}){
 const writes=[],messages=[],confirmations=[];let refreshes=0;
 const ctx={DirectRHCSV:rh,Papa:papa,Blob,csvDropBusy:false,isOwner:owner,activeAccount:{},selectedFolder:{id:'folder'},GRAPH:'https://graph.test',csvContributionCache:new Map(),lastFingerprint:'old',console:{error(){}},
  csvDropStatus:(m,error)=>messages.push({m,error}),showRefreshFeedback:m=>messages.push({m}),
  children:async()=>missing?[]:[{id:'canonical',name:'Direct RH Data.csv',file:{}},{id:'slots',name:'Slot Report.csv',file:{}}],
  graph:async()=>({eTag:etag,parentReference:{id:'folder'},'@microsoft.graph.downloadUrl':'https://download.test'}),
  awaitCSVConfirmation:async(...args)=>{confirmations.push(args);return !cancel},token:async()=>'fake',
  fetch:async(url,options)=>{if(options?.method==='PUT'){writes.push({url,...options});return {ok:status===200,status,json:async()=>({})}}return {ok:true,text:async()=>original}},
  loadFolder:async()=>{refreshes++},$:()=>null};
 const source=fs.readFileSync(require.resolve('../app.js'),'utf8');vm.createContext(ctx);
 vm.runInContext(source.slice(source.indexOf('async function replaceExistingCSV('),source.indexOf('function initCSVDragDrop(')),ctx);
 return {ctx,writes,messages,confirmations,refreshes:()=>refreshes,file:{name:'20261104_120000.csv',size:incoming.length,text:async()=>incoming}};
}
test('timestamp filename updates only canonical item with concurrency guard; no settings or extra file writes',async()=>{
 const f=fixture();await f.ctx.replaceExistingCSV(f.file);
 assert.equal(f.writes.length,1);assert.equal(f.writes[0].url,'https://graph.test/me/drive/items/canonical/content');
 assert.equal(f.writes[0].headers['If-Match'],'version-1');
 const text=await f.writes[0].body.text();assert.equal(text.split('08:00').length-1,1);assert.equal(text.split('08:01').length-1,1);
 assert.match(f.confirmations[0][2],/Append 1 new records/);assert.equal(f.refreshes(),1);assert.equal(f.ctx.csvDropBusy,false);
});
test('no new records, cancelled uploads, missing source, missing version, invalid CSV, and viewers do not write',async()=>{
 for(const options of [{incoming:csv([row('01-Jul-25','08:00')])},{cancel:true},{missing:true},{etag:null},{incoming:csv([row('31-Jun-25','08:01')])},{owner:false}]){
  const f=fixture(options);await f.ctx.replaceExistingCSV(f.file);assert.equal(f.writes.length,0,JSON.stringify(options));assert.equal(f.ctx.csvDropBusy,false);
 }
});
test('concurrent modification and failed upload remain visible and do not refresh or retry a write',async()=>{
 for(const status of [412,500]){const f=fixture({status});await f.ctx.replaceExistingCSV(f.file);assert.equal(f.writes.length,1);assert.equal(f.refreshes(),0);assert.ok(f.messages.at(-1).error);if(status===412)assert.match(f.messages.at(-1).m,/changed during this upload/)}
});
