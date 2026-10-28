const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const app=fs.readFileSync(__dirname+'/../app.js','utf8');
function fixture(){
 const key='expanded:chartPage:test',saved={[key]:true},root={hidden:false},summary={hidden:true},sign={textContent:'−'},log=[];
 let account='one';const attrs={'aria-expanded':'true','aria-controls':'test'};
 const button={getAttribute:n=>attrs[n],setAttribute:(n,v)=>{log.push('aria:'+v);attrs[n]=v},querySelector:()=>sign,closest:()=>({querySelector:()=>summary}),addEventListener:(_,fn)=>{button.handler=fn}};
 const ctx={Map,document:{querySelectorAll:()=>[button]},chartPreferenceRestoring:false,chartPreferenceAccount:()=>account,readChartPreferences:()=>saved,writeChartPreference:(k,v)=>{saved[k]=v;log.push('save:'+v)},chartExpansionKey:()=>key,$:()=>root,updateCollapsedFilterSummaries(){}};
 vm.createContext(ctx);
 const start=app.indexOf('const chartRestoredExpansions='),end=app.indexOf("document.addEventListener('DOMContentLoaded'",start);assert.ok(start>=0&&end>start);
 vm.runInContext(app.slice(start,end),ctx);
 const hs=app.indexOf("document.querySelectorAll('.collapse-filter-button').forEach(button=>{",app.indexOf('function updateCollapsedFilterSummaries(')),he=app.indexOf('function installAvailabilityCardTooltip(',hs);
 vm.runInContext(app.slice(hs,he),ctx);
 return {ctx,button,root,summary,sign,saved,key,log,setAccount:value=>{account=value}};
}
test('minus click saves before changing the DOM and remains collapsed after restoration',()=>{
 const f=fixture();f.ctx.restoreChartExpansions();f.log.length=0;f.button.handler();
 assert.equal(f.log[0],'save:false');assert.equal(f.saved[f.key],false);assert.equal(f.root.hidden,true);assert.equal(f.summary.hidden,false);assert.equal(f.sign.textContent,'+');
 for(let i=0;i<3;i++)f.ctx.restoreChartExpansions();
 assert.equal(f.button.getAttribute('aria-expanded'),'false');assert.equal(f.root.hidden,true);
 f.button.handler();assert.equal(f.root.hidden,false);assert.equal(f.saved[f.key],true);
});
test('saved expansion is applied once per account and changes reload on account switch',()=>{
 const f=fixture();f.ctx.restoreChartExpansions();const count=f.log.length;
 f.ctx.restoreChartExpansions();assert.equal(f.log.length,count);
 f.setAccount('two');f.saved[f.key]=false;f.ctx.restoreChartExpansions();
 assert.equal(f.root.hidden,true);assert.equal(f.button.getAttribute('aria-expanded'),'false');
});
