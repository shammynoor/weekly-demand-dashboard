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
 for(const id of ['accountLabel','headerSlotRefreshed','headerMenuToggle','headerCurrentPage','headerPersonalSettings','headerAccountMenu'])nodes[id]=new Node(id);
 const nav=new Node('nav');let closeCount=0,page;
 const ctx={$:id=>nodes[id],activeAccount:{name:'Microsoft Name',username:'user@example.com'},readPersonalSettings:()=>({}),window:{matchMedia:()=>media},document:{querySelector:()=>nav,addEventListener:(name,fn)=>{events[name]=fn}},MutationObserver:class{observe(){}},closeDashboardMenus(){closeCount++},showPage:p=>{page=p}};
 vm.createContext(ctx);vm.runInContext(source.slice(start,end),ctx);
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
test('Refreshed shows the imported file date and time in London, including midnight and DST',()=>{
 const f=fixture();f.ctx.updateHeaderSlotReportTime(Date.parse('2026-07-01T23:30:00Z'));assert.match(f.nodes.headerSlotRefreshed.textContent,/Refreshed: 2 Jul 2026,? 00:30/);
 f.ctx.updateHeaderSlotReportTime(Date.parse('2026-10-27T14:35:00Z'));assert.match(f.nodes.headerSlotRefreshed.textContent,/Refreshed: 27 Oct 2026,? 14:35/);
 f.ctx.updateHeaderSlotReportTime(null);assert.equal(f.nodes.headerSlotRefreshed.textContent,'Refreshed: unavailable');f.ctx.updateHeaderSlotReportTime(1e300);assert.equal(f.nodes.headerSlotRefreshed.textContent,'Refreshed: unavailable');
});

test('opening the responsive Menu does not dismiss a visible Forecast page',()=>{
 const forecast=fs.readFileSync(__dirname+'/../forecast.js','utf8'),a=forecast.indexOf("document.querySelectorAll('.dashboard-nav button').forEach("),b=forecast.indexOf(";$('dailyForecastWeekday')",a);
 let hides=0;const buttons=['headerMenuToggle','navHome'].map(id=>({id,handlers:{},matches:selector=>selector.split(',').includes('#'+id),addEventListener(name,fn){this.handlers[name]=fn}}));
 const ctx={document:{querySelectorAll:()=>buttons},$:()=>({classList:{add:()=>hides++}})};vm.createContext(ctx);vm.runInContext(forecast.slice(a,b),ctx);
 assert.equal(buttons[0].handlers.click,undefined);buttons[1].handlers.click();assert.equal(hides,2);
});
