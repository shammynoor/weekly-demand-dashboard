const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),core=require('../data-correctness.js');
const source=fs.readFileSync(__dirname+'/../app.js','utf8');
const draw=source.slice(source.indexOf('function drawCombinedDemandChart(){'),source.indexOf('\ndrawCombinedDemandChart();',source.indexOf('function drawCombinedDemandChart(){')));
function render({continuous=true,monthly=false,count=2,width=800,show=true,error=false,combined=false}={}){
 const categories=Array.from({length:count},(_,i)=>'Type '+i),dates=monthly?['2026-08','2026-09','2026-10']:['2026-10-12','2026-10-19','2026-10-26'];
 let result;
 const ctx={GPCorrectness:{...core,currentPeriodNowcast:()=>error?{error:'Insufficient history'}:{observed:40,rawObserved:40,estimate:100,low:80,high:120}},
 $:id=>id==='combinedDemandChart'?{clientWidth:width+157}:null,renderDemandClassificationToggles(){},setupDemandMonthSlider(){},hideEmptyHistoricBookingFilters(){},
 adjustDemandForListSize:false,purposeCategoryNames:[],purposeAssignments:{},disabledDemandPurposes:new Set(),purposeColours:{},capacityCategoryNames:[],capacityAssignments:{},disabledDemandCapacities:new Set(),capacityColours:{},
 groupUrgencyNames:categories,isDashboardToggleOption:()=>true,disabledDemandUrgencyGroups:new Set(),additiveLines:combined,selectedWeeks:dates,demandMonthStart:'2026-08',demandMonthEnd:'2026-10',
 continuousMode:continuous,monthlyMode:monthly,monthOfPeriod:d=>monthly?d+'-01':d,monthPosition:d=>9+(Number(d.slice(8,10))-1)/31,
 axis:{type:'linear',range:[0,12]},demandCategoryTotals:Object.fromEntries(categories.map(c=>[c,new Map(dates.map((d,i)=>[d,i===2?40:100]))])),
 thisMonth:'2026-10',thisWeek:'2026-10-26',todayLondon:'2026-10-28',showCurrentWeek:show,capacityCSVLatestModified:null,records:[],bookingTimeRecords:[],slotReportCoverage:[],excludeBankHolidays:false,
 demandSmoothingLevel:()=>0,periodIsCurrent:d=>d===dates.at(-1),urgencyGroupColour:()=>'#2678ad',yearBackgroundShapes:()=>[],
 Plotly:{react:(id,traces,layout)=>{result={traces,layout};return {then(){}}}}};
 vm.createContext(ctx);vm.runInContext(draw+'\ndrawCombinedDemandChart();',ctx);return result;
}
test('single current column stays on its booking date without an auxiliary axis',()=>{
 for(const continuous of [true,false])for(const monthly of [true,false]){
  const {traces,layout}=render({continuous,monthly,count:1});
  assert.equal(layout.xaxis2,undefined);assert.deepEqual(Array.from(layout.xaxis.domain),[0,1]);
  const line=traces.find(t=>t.meta?.estimateConnector),range=traces.find(t=>t.meta?.nowcast),circle=traces.find(t=>t.meta?.currentObserved),caps=traces.find(t=>t.meta?.rangeCaps);
  assert.equal(line.x.at(-1),range.x[0]);assert.equal(circle.x[0],range.x[0]);assert.equal(caps.x[0],range.x[0]);
  assert.equal(line.y.at(-1),100);assert.equal(line.y[0],100);assert.equal(line.line.dash,'dash');
  const history=traces.find(t=>t.name==='Type 0 2026');assert.equal(line.x[0],history.x.at(-1));assert.equal(line.y[0],history.y.at(-1));
  assert.ok(traces.every(t=>!t.xaxis));assert.equal(layout.xaxis.range[1],continuous?'2026-10-31':12);
 }
});
test('multiple columns nudge their connectors, circles and complete ranges together at all sizes',()=>{
 for(const width of [160,450,900])for(const count of [2,7]){
  const {traces}=render({width,count});const ends=traces.filter(t=>t.meta?.estimateConnector);
  assert.equal(new Set(ends.map(t=>t.x.at(-1))).size,count);
  for(let i=0;i<count;i++){
   const x=ends[i].x.at(-1),range=traces.filter(t=>t.meta?.nowcast)[i],caps=traces.filter(t=>t.meta?.rangeCaps)[i],circle=traces.filter(t=>t.meta?.currentObserved)[i];
   assert.ok(range.x.every(v=>v===x));assert.equal(caps.x[0],x);assert.equal(circle.x[0],x);
  }
 }
});
test('hidden current period and unavailable estimate preserve history and a single axis',()=>{
 for(const options of [{show:false},{error:true},{combined:true},{count:0}]){
  const {traces,layout}=render(options);assert.equal(layout.xaxis2,undefined);
  if(options.show===false||options.error)assert.ok(!traces.some(t=>t.meta?.estimateConnector||t.meta?.nowcast));
  if(options.error)assert.equal(traces.filter(t=>t.meta?.currentObserved).length,2);
  if(options.combined)assert.equal(traces.filter(t=>t.meta?.estimateConnector).length,1);
 }
});
test('nudges remain small, resize in pixels, and clamp at axis boundaries',()=>{
 for(const width of [24,160,900])for(const n of [1,2,7]){
  for(const value of [0,5,12]){
   const xs=core.currentPeriodPositions(value,[0,12],width,n);
   assert.ok(xs.every(x=>x>=0&&x<=12));assert.ok((xs.at(-1)-xs[0])*width/12<=24+1e-9);
   if(n===1)assert.equal(xs[0],value);
  }
 }
 assert.deepEqual(core.currentPeriodPositions(1,[0,12],800,0),[]);
 assert.deepEqual(core.currentPeriodPositions(1,[1,1],800,2),[1,1]);
});
module.exports={render};
