const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const rh=require('../direct-rh.js'),app=fs.readFileSync(__dirname+'/../app.js','utf8'),html=fs.readFileSync(__dirname+'/../index.html','utf8');
const fields=['tenant_name','request_date','request_time','request_type','admin_type','age_between_created_and_processed'];
const papa=process.env.DIRECT_RH_PAPA?require(process.env.DIRECT_RH_PAPA):{parse:text=>({data:text.trim().split(/\r?\n/).map(line=>line.split(',')),errors:[]})};
const row=(date,type='Admin',admin='OTHER_ENQUIRY')=>['darwin',date,'09:00',type,admin,'0 hour(s) 1 minute(s)'];
const csv=rows=>[fields,...rows].map(row=>row.join(',')).join('\n');
test('weekly counts use request dates, include Sunday in the preceding Monday, fill zero weeks, and exclude other request types',()=>{
 const summary=rh.weeklyOtherEnquiries(csv([row('29-Dec-25'),row('04-Jan-26'),row('05-Jan-26','U16'),row('19-Jan-26'),row('19-Jan-26','Admin','FIT_NOTE')]),papa);
 assert.deepEqual(summary,{records:5,total:3,weeks:[{week:'2025-12-29',count:2},{week:'2026-01-05',count:0},{week:'2026-01-12',count:0},{week:'2026-01-19',count:1}]});
 assert.equal(rh.weeklyOtherEnquiries(csv([row('19-Jan-26','Medical')]),papa).total,0);
 assert.throws(()=>rh.weeklyOtherEnquiries(csv([row('31-Feb-26')]),papa));
});
function ui(){
 class Node{
  constructor(){this.children=[];this.textContent='';this.hidden=false;this.classes=new Set();this.classList={contains:x=>this.classes.has(x),toggle:(x,on)=>on?this.classes.add(x):this.classes.delete(x),add:x=>this.classes.add(x),remove:x=>this.classes.delete(x)}}
  append(...nodes){for(const node of nodes)this.children.push(...(node.fragment?node.children:[node]))}replaceChildren(){this.children=[]}
 }
 const nodes={};const $=id=>nodes[id]??=(new Node());
 const ctx={$,rapidHealthSummary:null,document:{createElement:()=>new Node(),createDocumentFragment:()=>Object.assign(new Node(),{fragment:true})},isSignpostingFile:name=>name==='Signposting Auto Report.csv'};
 vm.createContext(ctx);vm.runInContext(app.slice(app.indexOf('function isDirectRHFile('),app.indexOf('// Parsed per-file contributions')),ctx);
 return {ctx,nodes,$};
}
test('home dates use each named file modification time in London and missing reports are explicit',()=>{
 const f=ui();f.ctx.updateHomeReportDates([{name:'Direct RH Data.csv',lastModifiedDateTime:'2026-07-01T08:15:00Z'},{name:'Slot Report.csv',lastModifiedDateTime:'2026-06-01T09:00:00Z'},{name:'20261104_120000.csv',lastModifiedDateTime:'2026-11-04T12:00:00Z'}]);
 assert.match(f.nodes.homeRapidHealthModified.textContent,/1 July 2026.*09:15/);
 assert.match(f.nodes.homeSlotModified.textContent,/1 June 2026.*10:00/);
 assert.equal(f.nodes.homeSignpostingModified.textContent,'No report found');
 f.ctx.updateHomeReportDates([]);assert.equal(f.nodes.homeRapidHealthModified.textContent,'No report found');
});
test('table handles missing files, formatted counts and replacement without accumulating old rows',()=>{
 const f=ui();f.ctx.renderRapidHealthTable();assert.equal(f.nodes.rapidHealthTable.hidden,true);assert.match(f.nodes.rapidHealthStatus.textContent,/No Direct RH/);
 f.ctx.rapidHealthSummary={weeks:[{week:'2026-01-05',count:1234},{week:'2026-01-12',count:0}]};f.ctx.renderRapidHealthTable();
 assert.equal(f.nodes.rapidHealthTable.hidden,false);assert.equal(f.nodes.rapidHealthBody.children[0].children[0].textContent,'5 Jan 2026');assert.equal(f.nodes.rapidHealthBody.children[0].children[1].textContent,'1,234');
 f.ctx.rapidHealthSummary={weeks:[{week:'2026-01-19',count:7}]};f.ctx.renderRapidHealthTable();assert.equal(f.nodes.rapidHealthBody.children.length,1);
});
test('Rapid Health is an Activity submenu page, available to viewers and hidden when another page is selected',()=>{
 const f=ui();let rendered=0;
 Object.assign(f.ctx,{isOwner:false,updateDashboardDateContext(){},refreshPrimaryApptTypeBoxes(){},closeDayView2Calendar(){},closeDashboardMenus(){},hideDayView2Tooltip(){},updateHomePracticeName(){},renderRapidHealthTable(){rendered++}});
 const start=app.indexOf('function showPage('),end=app.indexOf('}document.getElementById',start)+1;vm.runInContext(app.slice(start,end),f.ctx);
 f.ctx.showPage('rapidHealth');assert.equal(rendered,1);assert.equal(f.$('rapidHealthPage').classList.contains('hidden'),false);assert.equal(f.$('navRapidHealth').classList.contains('active'),true);assert.equal(f.$('signpostingPage').classList.contains('hidden'),true);
 f.ctx.showPage('home');assert.equal(f.$('rapidHealthPage').classList.contains('hidden'),true);
 assert.match(html,/<button type="button" id="navSignposting">Signposting<\/button><button type="button" id="navRapidHealth">Rapid Health<\/button>/);
 assert.match(html,/<th scope="col">Week beginning<\/th><th scope="col">Other enquiries<\/th>/);
});
test('canonical RH file gets a separate contribution; timestamp exports cannot be counted twice',()=>{
 const ctx={isDirectRHFile:name=>name.toLowerCase()==='direct rh data.csv',DirectRHCSV:rh,Papa:papa,isSignpostingFile:()=>false};vm.createContext(ctx);
 vm.runInContext(app.slice(app.indexOf('function parseCSVContribution('),app.indexOf('async function loadFolder(')),ctx);
 const text=csv([row('19-Jan-26')]);assert.equal(ctx.parseCSVContribution(text,'Direct RH Data.csv').summary.total,1);assert.equal(ctx.parseCSVContribution(text,'20260120_090000.csv').kind,'other');
});
test('folder refresh replaces the RH contribution, reuses its cache without double counting, and clears deleted data',async()=>{
 const f=ui(),noop=()=>{};let downloads=0,version=1,text=csv([row('19-Jan-26')]);
 let items=[{id:'rh',name:'Direct RH Data.csv',file:{},lastModifiedDateTime:'2026-01-20T12:00:00Z',size:version}];
 const ctx=f.ctx;
 Object.assign(ctx,{selectedFolder:{id:'folder'},busy:false,isOwner:true,lastFingerprint:'',csvContributionCache:new Map(),DirectRHCSV:rh,Papa:papa,
  setButtons:noop,showRefreshFeedback:noop,setSync:noop,children:async()=>items,restoreCSVCache:async()=>'scope',sourceItemUrl:id=>id,
  graph:async()=>({'@microsoft.graph.downloadUrl':'https://test.invalid'}),fetch:async()=>{downloads++;return {ok:true,text:async()=>text}},writeCSVCache:async()=>{},
  GPCorrectness:{coverageFromEntries:()=>[]},updateHeaderSlotReportTime:noop,renderSlotInclusion:noop,rebuildPeopleTypeLookup:noop,renderPeople:noop,renderSlotTable:noop,
  saveCategories:sync=>assert.equal(sync,false),rebuildCategoryFilter:noop,renderEditor:noop,renderClinicianEditor:noop,renderClinicianGroupEditor:noop,renderConsultEditor:noop,
  renderActiveAnalyticsPage:async()=>ctx.renderRapidHealthTable(),scheduleChartResize:noop,assignments:{},console:{error:noop}});
 vm.runInContext(app.slice(app.indexOf('function parseCSVContribution('),app.indexOf("$('signIn').onclick=")),ctx);
 await ctx.loadFolder(true);assert.equal(ctx.rapidHealthSummary.total,1);assert.equal(downloads,1);
 await ctx.loadFolder(true);assert.equal(ctx.rapidHealthSummary.total,1);assert.equal(downloads,1);
 text=csv([row('19-Jan-26'),row('20-Jan-26')]);items[0].size=++version;await ctx.loadFolder(true);assert.equal(ctx.rapidHealthSummary.total,2);assert.equal(downloads,2);
 items=[{id:'other',name:'unrelated.csv',file:{},size:1}];await ctx.loadFolder(true);assert.equal(ctx.rapidHealthSummary,null);assert.equal(f.nodes.rapidHealthTable.hidden,true);
 assert.equal(ctx.records.length,0);assert.equal(ctx.signpostingRecords.length,0);
});
