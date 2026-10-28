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
  assert.equal(line.y.at(-1),100);assert.equal(line.y[0],100);assert.equal(line.line.dash,'1px,3px');assert.equal(line.line.width,1.35);
  const history=traces.find(t=>t.name==='Type 0 2026');assert.equal(line.x[0],history.x.at(-1));assert.equal(line.y[0],history.y.at(-1));
  assert.ok(traces.every(t=>!t.xaxis));assert.equal(layout.xaxis.range[1],continuous?'2026-10-31':12);
 }
});
test('all categories share the current date for their connectors, circles and complete ranges',()=>{
 for(const width of [160,450,900])for(const count of [2,7]){
  const {traces}=render({width,count});const ends=traces.filter(t=>t.meta?.estimateConnector);
  assert.equal(new Set(ends.map(t=>t.x.at(-1))).size,1);
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
test('average capsules avoid overlap and stay within narrow and wide plot bounds',()=>{
 for(const width of [160,450,900])for(const count of [1,2,7,12]){
  const labels=Array.from({length:count},(_,i)=>({width:80,height:22,y:100+(i%3)*2}));
  const boxes=core.averageCapsuleLayout(labels,width,400);assert.equal(boxes.length,count);
  boxes.forEach((b,i)=>{
   assert.ok(b.left>=0&&b.top>=0&&b.left+b.width<=width&&b.top+b.height<=400);
   for(const p of boxes.slice(0,i))assert.ok(b.left>=p.left+p.width||p.left>=b.left+b.width||b.top>=p.top+p.height||p.top>=b.top+b.height);
  });
 }
 assert.deepEqual(core.averageCapsuleLayout([],800,400),[]);
});
test('separated average values stay centred on their own mean lines',()=>{
 const boxes=core.averageCapsuleLayout([{width:80,height:22,y:50},{width:80,height:22,y:200}],800,400);
 assert.equal(boxes[0].top+11,50);assert.equal(boxes[1].top+11,200);
});
module.exports={render};

test('Average capsules retain rounded DOM elements, track the line, and reposition after resize',()=>{
 class Element{
  constructor(){this.children=[];this.style={setProperty:(k,v)=>{this.style[k]=v}};this.offsetHeight=22}
  get offsetWidth(){return (this.textContent||'').length*6+20}
  appendChild(node){node.parent=this;this.children.push(node)}
  setAttribute(key,value){this[key]=value}
  remove(){this.parent.children=this.parent.children.filter(x=>x!==this)}
  querySelector(selector){return this.children.find(x=>selector==='.'+x.className)}
 }
 const target=new Element(),handlers=new Map();
 target._fullLayout={_size:{l:62,t:55,w:500,h:400},yaxis:{l2p:y=>400*(1-y/750)}};
 target.on=(event,fn)=>handlers.set(event,fn);target.removeListener=(event,fn)=>{if(handlers.get(event)===fn)handlers.delete(event)};
 const ctx={document:{createElement:()=>new Element()},GPCorrectness:core};vm.createContext(ctx);
 vm.runInContext(source.slice(source.indexOf('function installDemandAverageCapsules('),source.indexOf('function installDemandCardTooltip(')),ctx);
 const averages=[{category:'Urgent',average:500,colour:'#c44c5c'},{category:'Routine',average:501,colour:'#268a70'}];
 ctx.installDemandAverageCapsules(target,averages);let overlay=target.querySelector('.demand-average-overlay');
 assert.equal(overlay.children.length,2);assert.equal(overlay.children[0].textContent,'Average 500');assert.equal(overlay.children[0].className,'demand-average-pill');assert.equal(overlay.children[0].style['--average-line-y'],'11px');
 const left=overlay.children[1].style.left;target._fullLayout._size.w=200;handlers.get('plotly_afterplot')();assert.notEqual(overlay.children[1].style.left,left);
 ctx.installDemandAverageCapsules(target,averages);assert.equal(target.children.length,1);assert.equal(handlers.size,1);
});
