const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync(__dirname+'/../app.js','utf8');
const helpers=source.slice(source.indexOf('const chartShadowControls=['),source.indexOf('window.gpChartPreferences='));
function loadHelpers({inputs={},saved={one:{},two:{}},account='one'}={}){
 const nodes={...inputs},writes=[];
 const ctx={WeakSet,chartPreferenceAccount:()=>account,readChartPreferences:()=>saved[account]||{},writeChartPreference:(key,value)=>{(saved[account]??={})[key]=value;writes.push([key,value])},$:id=>nodes[id]||null};
 vm.createContext(ctx);vm.runInContext(helpers,ctx);return {ctx,nodes,saved,writes,setAccount:value=>{account=value}};
}
test('capacity chart shadow switch saves its account preference and styles the Today capsule',()=>{
 const input={checked:true,handlers:{},addEventListener(type,fn){this.handlers[type]=fn}};
 const f=loadHelpers({inputs:{urgentChartShadows:input}});
 const rect={attrs:{},setAttribute(name,value){this.attrs[name]=value}},group={textContent:'Today · 2 Nov',style:{},querySelector:()=>rect};
 const barLayer={style:{}},target={id:'availabilityChart',querySelector:()=>barLayer,querySelectorAll:()=>[group]};f.nodes.availabilityChart=target;
 f.ctx.restoreChartShadowControls();f.ctx.applyCapacityChartShadows(target,false);
 assert.match(barLayer.style.filter,/drop-shadow/);assert.match(group.style.filter,/drop-shadow/);assert.equal(rect.attrs.rx,'999');
 input.checked=false;input.handlers.change();
 assert.equal(f.saved.one['chartShadows:urgent'],false);assert.equal(barLayer.style.filter,'none');assert.equal(group.style.filter,'none');
});
test('historic chart shadow switch styles historical lines and excludes estimate connectors',()=>{
 const input={checked:true,handlers:{},addEventListener(type,fn){this.handlers[type]=fn}};
 const f=loadHelpers({inputs:{historicChartShadows:input}});
 const historicalPath={style:{}},estimatePath={style:{}};
 const groups=[{querySelectorAll:()=>[historicalPath]},{querySelectorAll:()=>[estimatePath]}];
 const target={data:[{type:'scatter',mode:'lines'},{type:'scatter',mode:'lines',meta:{estimateConnector:true}}],querySelectorAll:()=>groups};
 f.ctx.applyHistoricChartShadows(target);assert.match(historicalPath.style.filter,/drop-shadow/);assert.equal(estimatePath.style.filter,'none');
 input.checked=false;f.ctx.applyHistoricChartShadows(target);assert.equal(historicalPath.style.filter,'none');
});
test('shadow preferences restore separately for each account',()=>{
 const input={checked:true,handlers:{},addEventListener(type,fn){this.handlers[type]=fn}};
 const f=loadHelpers({inputs:{routineChartShadows:input},saved:{one:{'chartShadows:routine':false},two:{'chartShadows:routine':true}}});
 f.ctx.restoreChartShadowControls();assert.equal(input.checked,false);
 f.setAccount('two');f.ctx.restoreChartShadowControls();assert.equal(input.checked,true);
});

