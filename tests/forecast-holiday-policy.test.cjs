const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),core=require('../data-correctness.js');
const source=fs.readFileSync(__dirname+'/../forecast.js','utf8');
function holidays(year){
 const may=new Date(Date.UTC(year,4,1)),aug=new Date(Date.UTC(year,7,31));
 may.setUTCDate(1+(8-may.getUTCDay())%7);aug.setUTCDate(31-(aug.getUTCDay()+6)%7);
 const iso=d=>d.toISOString().slice(0,10),observed=(m,d)=>{const date=new Date(Date.UTC(year,m,d)),dow=date.getUTCDay();if(dow===6)date.setUTCDate(d+2);if(dow===0)date.setUTCDate(d+1);return iso(date)};
 return new Set([iso(may),iso(aug),observed(0,1),observed(11,25),observed(11,26)]);
}
function fixture(){
 class Element{
  constructor(tag='div'){this.tag=tag;this.children=[];this.hidden=false;this.style={setProperty(){}};this.isConnected=true;this.clientWidth=900;this.classes=new Set();this.classList={contains:c=>this.classes.has(c),add:c=>this.classes.add(c)};this.dataset={};this.textContent='';this.value='';this.checked=false}
  append(...nodes){this.children.push(...nodes)}replaceChildren(...nodes){this.children=nodes}
 }
 const nodes=Object.fromEntries(['dailyForecastHistoryChart','dailyForecastHistorySummary','dailyForecastWeekday','forecastAnalyticsPanel','forecastPredictionsPanel','dailyForecastHistory','dailyForecastAdjustListSize','dailyForecastNeedsTable','dailyForecastExcludeBankHolidayWeeks'].map(id=>[id,new Element()]));
 nodes.dailyForecastHistoryChart.parentElement=new Element();nodes.dailyForecastWeekday.value='1';nodes.dailyForecastHistory.value='52';nodes.forecastAnalyticsPanel.classes.add('hidden');
 const holidayWeeks=new Set();for(let year=2023;year<=2027;year++)for(const d of holidays(year))holidayWeeks.add(core.monday(d));
 const records=[];let i=0;for(let date='2023-01-02';date<='2026-10-27';date=core.shift(date,1),i++){
  const dow=new Date(date+'T12:00Z').getUTCDay();if(dow<1||dow>5)continue;
  records.push({bookedDay:date,slot:'Urgent slot',count:Math.round((100+dow*8+10*Math.sin(i*.1))*(holidayWeeks.has(core.monday(date))?4:1))});
 }
 const ctx={$:id=>nodes[id],document:{createElement:tag=>new Element(tag)},window:{gpChartPreferences:{account:()=> 'test-account'}},forecastLocalDate:()=> '2026-10-28',shift:core.shift,monday:core.monday,records,excludedSlotTypes:new Set(),englandWalesBankHolidays:holidays,capacityUrgencyCategory:()=> 'Urgent',forecastUrgencyNames:()=>['Urgent'],forecastUrgencyEnabled:()=>true,urgencyPlanningMode:new Map([['Urgent','daily']]),updateForecastHorizonLabels(){},syncForecastPreferences(){},ensureAnalyticsRoles:()=>({}),renderDailyUrgencyFilters(){},renderUrgencyPlanningModes(){},forecastRenderCache:{predictions:null,analytics:null},applyForecastHorizon:()=>true,forecastSuggestionRows:[],forecastSuggestionBody:null,assignments:{'Urgent slot':'Urgent'},groupUrgencyAssignments:{},practiceSettings:{}};
 vm.createContext(ctx);
 const policy=source.slice(source.indexOf('function forecastHistoryPolicy('),source.indexOf('function renderHistoricalWeekday('));
 const begin=source.indexOf('function renderHistoricalWeekday('),end=source.indexOf('\nfunction selectForecastTab(',begin);assert.ok(end>begin);
 const renderer=source.slice(begin,end).replace(' const fittedModel=special?', ' return {days,start,weekday,special,excludeHolidayWeeks};\n const fittedModel=special?');
 vm.runInContext(policy+renderer,ctx);
 const numbers=()=>{const out=[];const visit=node=>{if(node.className==='forecast-needs-number')out.push(node.textContent);for(const child of node.children||[])visit(child)};visit(nodes.dailyForecastNeedsTable);return out};
 const analytics=day=>{nodes.dailyForecastWeekday.value=String(day);nodes.forecastAnalyticsPanel.classes.delete('hidden');nodes.forecastPredictionsPanel.classes.add('hidden');return ctx.renderHistoricalWeekday()};
 const suggestions=()=>{nodes.forecastAnalyticsPanel.classes.add('hidden');nodes.forecastPredictionsPanel.classes.delete('hidden');ctx.renderHistoricalWeekday();return numbers()};
 return {ctx,nodes,holidayWeeks,analytics,suggestions};
}
test('suggestion results remain identical after inspecting any weekday or holiday-adjacent analytics',()=>{
 const f=fixture(),baseline=f.suggestions();assert.equal(baseline.length,67);assert.ok(baseline.some(n=>Number(n)>0));
 for(const day of [6,1,7,2,5]){
  f.analytics(day);f.ctx.forecastRenderCache.predictions=null;assert.deepEqual(f.suggestions(),baseline,'Changed suggestions after analytics day '+day);
 }
});
test('analytics Monday to Friday always exclude bank holiday weeks',()=>{
 const f=fixture();for(const day of [1,2,3,4,5]){
  const result=f.analytics(day);assert.equal(result.excludeHolidayWeeks,true);assert.ok(result.days.length>12);
  assert.ok(result.days.every(d=>new Date(d+'T12:00Z').getUTCDay()===day&&!f.holidayWeeks.has(core.monday(d))));
 }
});
test('holiday-adjacent analytics include holiday weeks and use all available history',()=>{
 const f=fixture(),before=f.analytics(6);assert.equal(before.excludeHolidayWeeks,false);assert.equal(before.start,'2023-01-02');assert.ok(before.days.includes('2025-12-24'));
 const after=f.analytics(7);assert.equal(after.excludeHolidayWeeks,false);assert.ok(after.days.includes('2026-05-05'));assert.ok(after.days.includes('2023-05-02'));
 assert.ok(after.days.some(d=>f.holidayWeeks.has(core.monday(d))));
});
test('suggestion cache ignores analytics day selection and obsolete checkbox preferences',()=>{
 const f=fixture();f.suggestions();const signature=f.ctx.forecastRenderCache.predictions.signature;
 for(const day of [6,7,2]){f.nodes.dailyForecastWeekday.value=String(day);f.nodes.dailyForecastExcludeBankHolidayWeeks.checked=false;f.ctx.forecastRenderCache.predictions=null;f.suggestions();assert.equal(f.ctx.forecastRenderCache.predictions.signature,signature)}
 assert.ok(!fs.readFileSync(__dirname+'/../index.html','utf8').includes('id="dailyForecastExcludeBankHolidayWeeks"'));
 assert.ok(!source.includes('dailyForecastExcludeBankHolidayWeeks'));
});

test('holiday-adjacent suggestion cells retain special history while ordinary days exclude holiday weeks',()=>{
 const f=fixture();f.suggestions();
 const cell=day=>f.ctx.forecastSuggestionRows.flatMap(row=>row.children).find(td=>td.title?.startsWith(day+': predicted'));
 const value=td=>Number(td.children.find(child=>child.className==='forecast-needs-number').textContent);
 const ordinary=cell('2026-11-03'),after=cell('2026-12-29');assert.ok(ordinary&&after);assert.ok(value(ordinary)<200);assert.ok(value(after)>300);assert.ok(after.classes.has('forecast-needs-after-holiday'));
});
test('weekly appointment allocations also remain independent of analytics weekday selection',()=>{
 const f=fixture();f.ctx.urgencyPlanningMode.set('Urgent','weekly');const baseline=f.suggestions();assert.equal(baseline.length,14);
 for(const day of [6,7,1,2]){f.analytics(day);f.ctx.forecastRenderCache.predictions=null;assert.deepEqual(f.suggestions(),baseline)}
});
