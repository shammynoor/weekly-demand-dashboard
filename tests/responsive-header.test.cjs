const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync(__dirname+'/../app.js','utf8');
const start=source.indexOf('function updateHeaderIdentity('),end=source.indexOf('function applyAccessMode(',start);
function fixture(){
 const nodes={},events={},media={matches:false,addEventListener:(_,fn)=>{media.changed=fn}};
 class Node{
  constructor(id){this.id=id;this.attrs={};this.handlers={};this.classes=new Set();this.textContent='';this.classList={contains:c=>this.classes.has(c),add:c=>this.classes.add(c),remove:c=>this.classes.delete(c),toggle:(c,on)=>{on?this.classes.add(c):this.classes.delete(c)}}}
  setAttribute(k,v){this.attrs[k]=v}getAttribute(k){return this.attrs[k]}removeAttribute(k){delete this.attrs[k]}
  addEventListener(k,v){this.handlers[k]=v}focus(){this.focused=true}contains(el){return el===this||el.parent===this}
  querySelectorAll(){return [{textContent:'Home',closest:()=>null}]}
 }
 for(const id of ['accountLabel','headerSlotRefreshed','headerMenuToggle','headerCurrentPage','headerPersonalSettings','headerAccountMenu','headerReportAge','headerReportAgeRow','headerEmisReminder'])nodes[id]=new Node(id);
 const nav=new Node('nav');let closeCount=0,page;
 const ctx={$:id=>nodes[id],activeAccount:{name:'Microsoft Name',username:'user@example.com'},readPersonalSettings:()=>({}),window:{matchMedia:()=>media},document:{querySelector:()=>nav,addEventListener:(name,fn)=>{events[name]=fn}},MutationObserver:class{observe(){}},closeDashboardMenus(){closeCount++},showPage:p=>{page=p}};
 vm.createContext(ctx);vm.runInContext(source.slice(start,end),ctx);ctx.setupResponsiveHeader();
 return {ctx,nodes,nav,events,media,Node,closeCount:()=>closeCount,page:()=>page};
}
test('wide navigation stays expanded by CSS while narrow windows can toggle and select pages',()=>{
 const f=fixture(),toggle=f.nodes.headerMenuToggle;
 toggle.handlers.click();assert.equal(toggle.getAttribute('aria-expanded'),'false');assert.equal(f.nav.classList.contains('header-menu-open'),false);
 f.media.matches=true;f.media.changed();toggle.handlers.click();assert.equal(toggle.getAttribute('aria-expanded'),'true');assert.equal(f.nav.classList.contains('header-menu-open'),true);
 const button={id:'navHome',closest:()=>null};f.nav.handlers.click({target:{closest:()=>button}});assert.equal(toggle.getAttribute('aria-expanded'),'false');
 toggle.handlers.click();f.media.matches=false;f.media.changed();assert.equal(f.nav.classList.contains('header-menu-open'),false);
});
test('Escape and outside clicks dismiss narrow navigation',()=>{
 const f=fixture();f.media.matches=true;f.nodes.headerMenuToggle.handlers.click();f.events.keydown({key:'Escape'});assert.equal(f.nodes.headerMenuToggle.focused,true);assert.equal(f.nav.classList.contains('header-menu-open'),false);
 f.nodes.headerMenuToggle.handlers.click();f.events.pointerdown({target:{}});assert.equal(f.nav.classList.contains('header-menu-open'),false);
});
test('account actions remain clickable; clicking outside closes the account menu',()=>{
 const f=fixture(),account=f.nodes.headerAccountMenu;account.setAttribute('open','');
 f.events.pointerdown({target:{parent:account}});assert.equal(account.getAttribute('open'),'');
 f.nodes.headerPersonalSettings.handlers.click();assert.equal(f.page(),'personalSettings');assert.equal(account.getAttribute('open'),undefined);
 account.setAttribute('open','');f.events.pointerdown({target:{}});assert.equal(account.getAttribute('open'),undefined);
});
test('header name uses the personal name, then Microsoft name, then email',()=>{
 const f=fixture();f.ctx.readPersonalSettings=()=>({name:'Personal Name'});f.ctx.updateHeaderIdentity();assert.equal(f.nodes.accountLabel.textContent,'Personal Name');assert.equal(f.nodes.accountLabel.title,'user@example.com');
 f.ctx.readPersonalSettings=()=>({name:' '});f.ctx.updateHeaderIdentity();assert.equal(f.nodes.accountLabel.textContent,'Microsoft Name');f.ctx.activeAccount.name='';f.ctx.updateHeaderIdentity();assert.equal(f.nodes.accountLabel.textContent,'user@example.com');
});
test('Clinical download shows the imported file date and time in London, including midnight and DST',()=>{
 const f=fixture();f.ctx.updateHeaderSlotReportTime(Date.parse('2026-07-01T23:30:00Z'));assert.match(f.nodes.headerSlotRefreshed.textContent,/Last clinical system download: 2 Jul 2026,? 00:30/);
 f.ctx.updateHeaderSlotReportTime(Date.parse('2026-10-27T14:35:00Z'));assert.match(f.nodes.headerSlotRefreshed.textContent,/Last clinical system download: 27 Oct 2026,? 14:35/);
 f.ctx.updateHeaderSlotReportTime(null);assert.equal(f.nodes.headerSlotRefreshed.textContent,'Last clinical system download: unavailable');f.ctx.updateHeaderSlotReportTime(1e300);assert.equal(f.nodes.headerSlotRefreshed.textContent,'Last clinical system download: unavailable');
});

test('opening the responsive Menu does not dismiss a visible Forecast page',()=>{
 const forecast=fs.readFileSync(__dirname+'/../forecast.js','utf8'),a=forecast.indexOf("document.querySelectorAll('.dashboard-nav button').forEach("),b=forecast.indexOf(";$('dailyForecastWeekday')",a);
 let hides=0;const buttons=['headerMenuToggle','navHome'].map(id=>({id,handlers:{},matches:selector=>selector.split(',').includes('#'+id),addEventListener(name,fn){this.handlers[name]=fn}}));
 const ctx={document:{querySelectorAll:()=>buttons},$:()=>({classList:{add:()=>hides++}})};vm.createContext(ctx);vm.runInContext(forecast.slice(a,b),ctx);
 assert.equal(buttons[0].handlers.click,undefined);buttons[1].handlers.click();assert.equal(hides,2);
});


test('header startup initializes the element lookup before executing responsive setup',()=>{
 const lookup=source.match(/^const \$=.+$/m)[0];
 const call='setupResponsiveHeader();';
 const pieces=[{at:start,code:source.slice(start,end)},{at:source.indexOf(lookup),code:lookup},{at:source.indexOf('\n'+call),code:call}].sort((a,b)=>a.at-b.at);
 const ctx={document:{querySelector:()=>null,getElementById:()=>null}};vm.createContext(ctx);
 assert.doesNotThrow(()=>vm.runInContext(pieces.map(p=>p.code).join('\n'),ctx));
});

test('compact dropdowns are kept within the navigation bounds',()=>{
 const begin=source.indexOf('function positionDashboardSubmenu('),finish=source.indexOf('function setDashboardMenuOpen(',begin);
 let narrow=true,settings=false,position={left:200,right:450};const style={setProperty:(key,value)=>{style[key]=value}};
 const menu={classList:{contains:()=>settings},querySelector:()=>({style,getBoundingClientRect:()=>position}),closest:()=>({getBoundingClientRect:()=>({left:40,right:300})})};
 const ctx={window:{matchMedia:()=>({matches:narrow})}};vm.createContext(ctx);vm.runInContext(source.slice(begin,finish),ctx);
 ctx.positionDashboardSubmenu(menu);assert.equal(style['--nav-dropdown-shift'],'-154px');
 position={left:30,right:280};ctx.positionDashboardSubmenu(menu);assert.equal(style['--nav-dropdown-shift'],'14px');
 style['--nav-dropdown-shift']='untouched';narrow=false;ctx.positionDashboardSubmenu(menu);assert.equal(style['--nav-dropdown-shift'],'untouched');
 narrow=true;settings=true;ctx.positionDashboardSubmenu(menu);assert.equal(style['--nav-dropdown-shift'],'untouched');
});

test('refresh progress replaces the clinical download label in the same header area',()=>{
 const f=fixture();f.nodes.refreshFeedback=new f.Node('refreshFeedback');
 const a=source.indexOf('function showRefreshFeedback('),b=source.indexOf('function resizeVisibleCharts(',a);vm.runInContext(source.slice(a,b),f.ctx);
 f.ctx.updateHeaderSlotReportTime(Date.parse('2026-10-27T14:35Z'));
 f.ctx.showRefreshFeedback('Updating data…',true);assert.equal(f.nodes.headerSlotRefreshed.textContent,'Updating data…');assert.equal(f.nodes.headerSlotRefreshed.getAttribute('aria-busy'),'true');assert.equal(f.nodes.refreshFeedback.textContent,'');
 f.ctx.updateHeaderSlotReportTime(Date.parse('2026-10-28T09:15Z'));assert.equal(f.nodes.headerSlotRefreshed.textContent,'Updating data…');
 f.ctx.showRefreshFeedback('Data refreshed successfully');assert.match(f.nodes.headerSlotRefreshed.textContent,/Last clinical system download: 28 Oct 2026,? 09:15/);assert.equal(f.nodes.headerSlotRefreshed.getAttribute('aria-busy'),'false');
 assert.equal(f.nodes.refreshFeedback.classList.contains('header-feedback-quiet'),true);
 f.ctx.showRefreshFeedback('Refreshing data…',true);f.ctx.showRefreshFeedback('Refresh failed: offline');assert.match(f.nodes.headerSlotRefreshed.textContent,/28 Oct 2026/);assert.equal(f.nodes.refreshFeedback.textContent,'Refresh failed: offline');assert.equal(f.nodes.refreshFeedback.classList.contains('header-feedback-quiet'),false);
});


test('refresh age includes minutes below two hours and an EMIS reminder only above two hours',()=>{
 const f=fixture(),timestamp=Date.parse('2026-10-28T10:00Z');
 f.ctx.updateHeaderReportAge(timestamp,timestamp+119*60000);assert.equal(f.nodes.headerReportAge.textContent,'Refreshed 1 hour 59 minutes ago');assert.equal(f.nodes.headerEmisReminder.hidden,true);
 f.ctx.updateHeaderReportAge(timestamp,timestamp+120*60000);assert.equal(f.nodes.headerReportAge.textContent,'Refreshed 2 hours ago');assert.equal(f.nodes.headerEmisReminder.hidden,true);
 f.ctx.updateHeaderReportAge(timestamp,timestamp+121*60000);assert.equal(f.nodes.headerEmisReminder.hidden,false);assert.equal(f.nodes.headerReportAgeRow.hidden,false);
 f.ctx.updateHeaderReportAge(timestamp,timestamp+20*60000);assert.equal(f.nodes.headerReportAge.textContent,'Refreshed 0 hours 20 minutes ago');
 f.ctx.updateHeaderReportAge(null);assert.equal(f.nodes.headerReportAgeRow.hidden,true);assert.equal(f.nodes.headerEmisReminder.hidden,true);
});
function csvFixture({cancel=false,fail=false,refreshFail=false}={}){
 const f=fixture();f.nodes.refreshFeedback=new f.Node('refreshFeedback');f.nodes.csvDropStatus=new f.Node('csvDropStatus');
 const a=source.indexOf('function showRefreshFeedback('),b=source.indexOf('function resizeVisibleCharts(',a);vm.runInContext(source.slice(a,b),f.ctx);
 vm.runInContext(source.slice(source.indexOf('function csvDropStatus('),source.indexOf('function csvDropOverlay(')),f.ctx);
 let puts=0;
 Object.assign(f.ctx,{csvDropBusy:false,isOwner:true,selectedFolder:{id:'folder'},GRAPH:'https://example.test',csvContributionCache:new Map(),lastFingerprint:'old',console:{error(){}},children:async()=>[{id:'report',name:'Slot Report.csv',file:{}}],validateCSVForReplacement:()=>({columns:['a','b']}),graph:async()=>({parentReference:{id:'folder'},'@microsoft.graph.downloadUrl':'https://example.test/original'}),awaitCSVConfirmation:async()=>!cancel,token:async()=> 'test-token',fetch:async(url,options)=>{if(options?.method==='PUT'){puts++;return {ok:!fail,status:500,json:async()=>({error:{message:'Upload failed'}})}}return {ok:true,text:async()=> 'a,b\n1,2'}},loadFolder:async()=>{
  f.ctx.showRefreshFeedback('Refreshing data…',true);
  if(refreshFail)f.ctx.showRefreshFeedback('Refresh failed: offline');else{f.ctx.updateHeaderSlotReportTime(Date.parse('2026-10-28T12:15Z'));f.ctx.showRefreshFeedback('Data refreshed successfully')}
 }});
 vm.runInContext(source.slice(source.indexOf('async function replaceExistingCSV('),source.indexOf('function initCSVDragDrop(')),f.ctx);
 f.ctx.updateHeaderSlotReportTime(Date.parse('2026-10-27T14:35Z'));
 return {...f,puts:()=>puts,file:{name:'Slot Report.csv',size:10,text:async()=> 'a,b\n1,2'}};
}
test('CSV replacement clears upload progress after the dashboard refresh completes',async()=>{
 const f=csvFixture();await f.ctx.replaceExistingCSV(f.file);assert.equal(f.puts(),1);assert.equal(f.ctx.csvDropBusy,false);assert.equal(f.nodes.csvDropStatus.textContent,'');assert.equal(f.nodes.refreshFeedback.textContent,'');assert.match(f.nodes.headerSlotRefreshed.textContent,/Last clinical system download: 28 Oct 2026,? 12:15/);
});
test('cancelled CSV replacement clears progress without uploading',async()=>{
 const f=csvFixture({cancel:true});await f.ctx.replaceExistingCSV(f.file);assert.equal(f.puts(),0);assert.equal(f.nodes.refreshFeedback.textContent,'');assert.match(f.nodes.headerSlotRefreshed.textContent,/27 Oct 2026/);
});
test('CSV upload or refresh failures remain visible and restore the last download label',async()=>{
 for(const options of [{fail:true},{refreshFail:true}]){
  const f=csvFixture(options);await f.ctx.replaceExistingCSV(f.file);assert.ok(f.nodes.refreshFeedback.textContent.includes('failed'));assert.equal(f.nodes.refreshFeedback.classList.contains('header-feedback-quiet'),false);assert.match(f.nodes.headerSlotRefreshed.textContent,/Last clinical system download:/);assert.equal(f.ctx.csvDropBusy,false);
 }
});
