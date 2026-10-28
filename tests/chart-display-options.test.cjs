const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync(__dirname+'/../app.js','utf8');
function fixture(){
 const nodes={},selects=[],pickers=[],events={},saved={one:{},two:{}};let account='one',focus,changes=0;
 class Element{
  constructor(tag='div'){this.tagName=tag.toUpperCase();this.children=[];this.dataset={};this.attrs={};this.handlers={};this.hidden=false;this.disabled=false;this.checked=false;this.textContent='';this.className=''}
  setAttribute(k,v){this.attrs[k]=v}getAttribute(k){return this.attrs[k]??null}
  addEventListener(type,fn){(this.handlers[type]??=[]).push(fn)}dispatchEvent(event){this.onchange?.(event);for(const fn of this.handlers[event.type]||[])fn(event)}
  append(...children){this.children.push(...children);for(const child of children)child.parent=this}
  after(picker){pickers.push(picker)}focus(){focus=this}contains(node){return this===node||this.children.some(c=>c.contains(node))}
  closest(){return {id:'chartPage'}}querySelectorAll(){return pickers}
  get selectedIndex(){return this.options.findIndex(o=>o.value===this.value)}
 }
 const button=new Element('button');button.setAttribute('aria-controls','options');nodes.options=new Element();nodes.options.hidden=true;
 for(const [id,checkedValue,values,checked] of [['periodMode','month',['week','month'],false],['displayMode','continuous',['continuous','overlay'],true]]){
  const model=new Element('input');model.checked=checked;model.onchange=()=>{changes++};nodes[id]=model;
  const select=new Element('select');select.id=id+'Select';select.dataset={chartProxy:id,checkedValue};select.options=values.map(value=>({value,textContent:value}));select.value=values[0];select.setAttribute('aria-label',id);selects.push(select);
 }
 const ctx={Map,Event:class{constructor(type,options){this.type=type;Object.assign(this,options)}},window:{},document:{createElement:tag=>new Element(tag),addEventListener:(type,fn)=>{(events[type]??=[]).push(fn)},querySelectorAll:selector=>selector.startsWith('.chart-display-options-button')?[button]:selector.startsWith('select[')?selects:pickers},$:id=>nodes[id],scheduleChartResize(){},chartPreferenceAccount:()=>account,readChartPreferences:()=>saved[account],writeChartPreference:(key,value)=>{saved[account][key]=value}};
 vm.createContext(ctx);const begin=source.indexOf('// Shared chart disclosure and rounded select menus.'),end=source.indexOf('async function refreshDashboardData(',begin);vm.runInContext(source.slice(begin,end),ctx);
 const key=()=>({type:'keydown',key:'ArrowDown',preventDefault(){},stopPropagation(){}});
 return {ctx,nodes,button,selects,pickers,saved,events,key,changes:()=>changes,focus:()=>focus,setAccount:a=>{account=a}};
}
test('Data and Timeline dropdowns retain the original checkbox models and restore saved changes',()=>{
 const f=fixture();assert.equal(f.pickers.length,2);assert.equal(f.selects[0].hidden,true);assert.equal(f.pickers[1].children[0].textContent,'continuous');
 const [trigger,list]=f.pickers[0].children;trigger.dispatchEvent({type:'click'});assert.equal(list.hidden,false);list.children[1].dispatchEvent({type:'click'});
 assert.equal(f.nodes.periodMode.checked,true);assert.equal(f.changes(),1);assert.equal(trigger.textContent,'month');assert.equal(list.hidden,true);
 f.nodes.periodMode.checked=false;f.nodes.periodMode.dispatchEvent({type:'change'});assert.equal(trigger.textContent,'week');
});
test('rounded dropdown menus support keyboard focus, Escape and outside dismissal',()=>{
 const f=fixture(),[button,list]=f.pickers[0].children;button.dispatchEvent(f.key());assert.equal(list.hidden,false);assert.equal(f.focus(),list.children[0]);
 list.children[0].dispatchEvent(f.key());assert.equal(f.focus(),list.children[1]);const escape={...f.key(),key:'Escape'};list.children[1].dispatchEvent(escape);assert.equal(list.hidden,true);assert.equal(f.focus(),button);
 button.dispatchEvent({type:'click'});for(const fn of f.events.pointerdown)fn({target:new Object()});assert.equal(list.hidden,true);
});
test('Display options opens in flow and saves expansion independently for each account',()=>{
 const f=fixture();assert.equal(f.nodes.options.hidden,true);f.button.dispatchEvent({type:'click'});assert.equal(f.nodes.options.hidden,false);
 const key='displayOptions:chartPage:options';assert.equal(f.saved.one[key],true);
 f.setAccount('two');f.ctx.restoreChartDisplayOptions();assert.equal(f.nodes.options.hidden,true);f.setAccount('one');f.ctx.restoreChartDisplayOptions();assert.equal(f.nodes.options.hidden,false);
 f.nodes.options.dispatchEvent({...f.key(),key:'Escape'});assert.equal(f.nodes.options.hidden,true);assert.equal(f.button.getAttribute('aria-expanded'),'false');assert.equal(f.saved.one[key],false);
});
test('opening one chart dropdown closes the other and closing Display options closes both',()=>{
 const f=fixture();f.button.dispatchEvent({type:'click'});f.pickers[0].children[0].dispatchEvent({type:'click'});f.pickers[1].children[0].dispatchEvent({type:'click'});assert.equal(f.pickers[0].children[1].hidden,true);assert.equal(f.pickers[1].children[1].hidden,false);
 f.button.dispatchEvent({type:'click'});assert.ok(f.pickers.every(p=>p.children[1].hidden));
});
