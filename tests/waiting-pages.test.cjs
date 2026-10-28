const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),core=require('../data-correctness.js');
const app=fs.readFileSync(__dirname+'/../app.js','utf8');
function fixture(){
 class Element{
  constructor(){this.children=[];this.dataset={};this.style={setProperty(){}};this.handlers={};this.classList={contains:()=>false};this.parentElement={clientWidth:400,getBoundingClientRect:()=>({left:0})};this.clientWidth=400}
  append(...nodes){this.children.push(...nodes)}replaceChildren(...nodes){this.children=nodes}setAttribute(){}addEventListener(event,fn){this.handlers[event]=fn}getBoundingClientRect(){return {left:0}}
 }
 const nodes={};for(const id of ['waitingTimesChart','waitingUrgencyToggles','waitingTimesPage','waitingTrendChart','waitingTrendUrgencyToggles','waitingOverTimePage','waitingSuppressSmallBins','waitingTrendSuppressSmall','waitingTrendIgnoreRoutineSameDay','waitingRangeStart','waitingRangeEnd','waitingRangeFill','waitingRangeFromLabel','waitingRangeToLabel','waitingTrendRangeStart','waitingTrendRangeEnd','waitingTrendRangeFill','waitingTrendRangeFromLabel','waitingTrendRangeToLabel'])nodes[id]=new Element();
 nodes.waitingSuppressSmallBins.checked=false;nodes.waitingTrendSuppressSmall.checked=false;nodes.waitingTrendIgnoreRoutineSameDay.checked=false;
 const records=[{slot:'Phone',clinician:'GP',person:'Doctor',appointmentMonth:'2026-10',appointmentDate:'2026-10-15',leadMinutes:2880,day:2,status:'Booked',count:4},{slot:'Legacy',clinician:'GP',person:'Doctor',appointmentMonth:'2026-10',appointmentDate:'2026-10-15',leadMinutes:4320,day:3,status:'Booked',count:3}];
 let rendered;
 const ctx={GPCorrectness:core,document:{createElement:()=>new Element()},window:{Plotly:{}},$:id=>nodes[id],Plotly:{react:(id,traces,layout)=>{rendered={traces,layout};return {then(){}}}},
  groupUrgencyNames:['Urgent','Routine'],assignments:{Phone:'Routine',Legacy:'Urgent detail'},groupUrgencyAssignments:{'Urgent detail':'Urgent'},consultCategoryNames:['Phone'],consultAssignments:{Phone:'Phone',Legacy:'Phone'},purposeCategoryNames:[],purposeAssignments:{},capacityCategoryNames:['Core'],capacityAssignments:{Phone:'Core',Legacy:'Core'},clinicianGroupAssignments:{GP:'Medical'},clinicianGroupNames:['Medical'],effectiveClinicianType:()=> 'GP',suggestConsultMode:()=>'',
  waitingTimeRecords:records,waitingTrendRecords:records,waitingTrendAvailableMonths:['2026-10'],waitingDisabledUrgency:new Set(),waitingTrendDisabledUrgency:new Set(),waitingFilterDisabled:Object.fromEntries(['consult','purpose','capacity','clinician'].map(k=>[k,new Set()])),waitingTrendDisabled:Object.fromEntries(['consult','purpose','capacity','clinician'].map(k=>[k,new Set()])),excludedSlotTypes:new Set(),waitingAppointmentMonthStart:'',waitingAppointmentMonthEnd:'',waitingAppointmentMonthBounds:'',waitingTrendMonthStart:'',waitingTrendMonthEnd:'',waitingTrendMonthBounds:'',capacityCSVLatestModified:Date.parse('2026-10-28T12:00Z'),
  readChartPreferences:()=>({}),writeChartPreference(){},isDashboardToggleOption:n=>!!n,urgencyGroupColour:()=> '#1268a8',clinicianGroupColour:()=> '#1268a8',typeColour:()=> '#1268a8',consultModeColours:{},purposeColours:{},capacityColours:{},clinicianGroupColours:{},updateCollapsedFilterSummaries(){}};
 vm.createContext(ctx);
 assert.ok(app.indexOf('function monthRangeSelection(')<app.indexOf('function refresh(){'),'Waiting range helpers must be in shared scope');
 vm.runInContext(app.slice(app.indexOf('function monthRangeSelection('),app.indexOf('function refresh(){')),ctx);
 vm.runInContext(app.slice(app.indexOf('function chartDimensionLookup('),app.indexOf('function installWaitingCardTooltip(')),ctx);
 vm.runInContext(app.slice(app.indexOf('function capacityUrgencyCategory('),app.indexOf('function renderUrgentCategoryToggles(')),ctx);
 vm.runInContext(app.slice(app.indexOf('function renderWaitingTimesChart('),app.indexOf("\n$('waitingSuppressSmallBins')")),ctx);
 vm.runInContext(app.slice(app.indexOf('function renderWaitingTrendChart('),app.indexOf('// Day View uses')),ctx);
 return {ctx,nodes,rendered:()=>rendered};
}
for(const [fn,root] of [['renderWaitingTimesChart','waitingUrgencyToggles'],['renderWaitingTrendChart','waitingTrendUrgencyToggles']]){
 test(fn+' renders directly and shows both direct and legacy appointment type toggles',()=>{
  const f=fixture();f.ctx[fn]();assert.equal(f.nodes[root].children.length,2);assert.ok(f.rendered().traces.some(t=>t.y.some(v=>v>0)));
  const labels=f.nodes[root].children.map(row=>row.children[1].textContent);assert.deepEqual(labels,['Urgent','Routine']);
 });
 test(fn+' keeps primary filters available with conflicting secondary filters and empty data',()=>{
  const f=fixture();for(const disabled of [f.ctx.waitingFilterDisabled,f.ctx.waitingTrendDisabled]){disabled.consult.add('Phone');disabled.capacity.add('Core')}
  f.ctx[fn]();assert.equal(f.nodes[root].children.length,2);
  f.ctx.waitingTimeRecords=[];f.ctx.waitingTrendRecords=[];f.ctx[fn]();assert.equal(f.nodes[root].children.length,2);
 });
 test(fn+' appointment type toggles continue to filter and redraw the chart',()=>{
  const f=fixture();f.ctx[fn]();const row=f.nodes[root].children.find(row=>row.children[1].textContent==='Routine'),input=row.children[0].children[0];
  input.checked=false;input.handlers.change();assert.equal(f.nodes[root].children.length,2);
  if(fn==='renderWaitingTimesChart')assert.equal(f.rendered().traces[0].y.reduce((a,b)=>a+b,0),3);else assert.ok(f.rendered().traces.every(t=>t.name!=='Routine'));
 });
}
