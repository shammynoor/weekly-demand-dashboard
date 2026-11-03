const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const app=fs.readFileSync(__dirname+'/../app.js','utf8');
function fixture(extra={}){
 const ctx={categoryNames:[],clinicianCategoryNames:[],purposeCategoryNames:[],capacityCategoryNames:[],clinicianGroupNames:[],groupUrgencyNames:[],appointmentCategoryNames:[],...extra};
 vm.createContext(ctx);
 vm.runInContext(app.slice(app.indexOf('function validCategoryNames('),app.indexOf('async function loadCloudCategories('))+app.slice(app.indexOf('function parseSettingsJSON('),app.indexOf('// Shared chart disclosure')),ctx);
 return ctx;
}
const settings=()=>({version:9,categories:['Routine'],assignments:{slot:'Routine'},analyticsRoles:{appointmentTypes:{Routine:{kind:'routine'}},clinicianGroups:{Doctors:{doctor:true}}}});
test('current and legacy JSON, including a UTF-8 BOM, load without authentication',()=>{
 const ctx=fixture();assert.equal(ctx.parseSettingsJSON('\uFEFF'+JSON.stringify(settings())).version,9);
 assert.equal(ctx.parseSettingsJSON('{"categories":["Routine"],"assignments":{}}').categories[0],'Routine');
});
test('malformed JSON and invalid optional fields are rejected before applying any settings',async()=>{
 const cases=[null,[],{...settings(),version:10},{...settings(),practiceSettings:{anchors:[]}},{...settings(),purposeAssignments:[]},{...settings(),peopleLocum:null},{...settings(),togglePriority:{urgency:null}},{...settings(),folder:{id:7}},{...settings(),excludedSlotTypes:[null]}];
 const ctx=fixture();assert.throws(()=>ctx.parseSettingsJSON('<html>Sign in</html>'),/not valid JSON/);
 for(const data of cases){
  ctx.activeAccount={};ctx.cloudChangeSerial=0;ctx.cloudReady=false;ctx.categoryDirty=false;ctx.cloudBusy=false;
  ctx.cloudFile=async()=>data;ctx.cloudNotice=()=>{};ctx.analyticsRoles={sentinel:true};
  vm.runInContext(app.slice(app.indexOf('async function loadCloudCategories('),app.indexOf('let categoryDirty=')),ctx);
  // null means a missing file to the loader; reject it at the JSON boundary instead.
  if(data===null){assert.throws(()=>ctx.parseSettingsJSON('null'),/Invalid settings/);continue;}
  await assert.rejects(ctx.loadCloudCategories({strict:true}),/Invalid settings/);
  assert.equal(ctx.analyticsRoles.sentinel,true);assert.equal(ctx.cloudReady,false);
 }
});
test('shared-file failures never fall back to and republish private settings',async()=>{
 let requests=0;
 const ctx=fixture({isOwner:true,readSharedSettings:async()=>{throw Error('HTTP 403')},token:async()=>'',fetch:async()=>{requests++;return {ok:true,text:async()=>JSON.stringify(settings())}},GRAPH:'graph',CLOUD_SETTINGS_NAME:'settings.json'});
 vm.runInContext(app.slice(app.indexOf('async function cloudFile('),app.indexOf('function validCategoryNames(')),ctx);
 await assert.rejects(ctx.cloudFile(),/403/);assert.equal(requests,0);
 ctx.readSharedSettings=async()=>null;
 assert.equal((await ctx.cloudFile()).version,9);assert.equal(requests,1);
 ctx.isOwner=false;assert.equal(await ctx.cloudFile(),null);assert.equal(requests,1);
});
