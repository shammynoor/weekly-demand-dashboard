const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync(__dirname+'/../app.js','utf8');
test('Day View restores account date before initializing and formatting its first render',()=>{
 const begin=source.indexOf('function syncDayView2Preferences()'),end=source.indexOf('function chooseDayView2Date(',begin);
 const renderStart=source.indexOf('function renderDayView2(){'),renderEnd=source.indexOf(' const date=input.value',renderStart);
 const nodes={dayView2Date:{value:''},dayView2DatePicker:{},dayView2Page:{classList:{contains:()=>false}},dayView2UrgencyToggles:{}};
 let account='one',saved={};const ctx={$:id=>nodes[id],chartPreferenceAccount:()=>account,readChartPreferences:()=>saved,dayView2DisabledUrgencies:new Set(),dayView2PreferencesAccount:null,dayView2WorkingDate:d=>d,dayView2LocalDate:()=> '2026-10-28'};
 vm.createContext(ctx);vm.runInContext(source.slice(begin,end)+source.slice(renderStart,renderEnd)+'return input.value;}',ctx);
 assert.equal(ctx.renderDayView2(),'2026-10-28');assert.match(nodes.dayView2DatePicker.textContent,/28 Oct 2026/);
 account='two';saved={'dayView2:selectedDate':'2026-10-27'};assert.equal(ctx.renderDayView2(),'2026-10-27');assert.match(nodes.dayView2DatePicker.textContent,/27 Oct 2026/);
});
