const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync(__dirname+'/../app.js','utf8');
const html=fs.readFileSync(__dirname+'/../index.html','utf8');
test('capacity Today label stays inside the plot in a clear top strip at every range edge',()=>{
 const begin=source.indexOf('function capacityTodayAnnotation('),end=source.indexOf('function capacityWeekBeginningLabel(',begin),ctx={};vm.createContext(ctx);vm.runInContext(source.slice(begin,end),ctx);
 for(const width of [240,320,460,736,1400])for(const count of [2,5,10,70])for(const x of [0,count/2,count-1]){
  const label=ctx.capacityTodayAnnotation(x,count,width,'2026-10-28');
  assert.equal(label.xref,'paper');assert.equal(label.y,0.985);assert.equal(label.yanchor,'top');assert.equal(label.xanchor,'center');assert.equal(label.bgcolor,'rgba(255,255,255,0.75)');assert.equal(label.bordercolor,'#206ca6');assert.equal(label.borderwidth,1);assert.equal(label.borderpad,4);assert.equal(label.font.size,10);assert.equal(label.font.family,'Segoe UI,Arial,sans-serif');
  const half=Math.min(.5,(width<480?83:130)/Math.max(100,width-80));assert.ok(label.x>=half&&label.x<=1-half);assert.match(label.text,/Today/);assert.doesNotMatch(label.text,/<br>/);
 }
 assert.match(source,/annotations.push\(capacityTodayAnnotation\(x,buckets.length,target.clientWidth\)\)/);
 assert.match(source,/annotations.push\(capacityTodayAnnotation\(todayX,dates.length,target.clientWidth\)\)/);
 assert.equal((source.match(/x0:(?:x|todayX),x1:(?:x|todayX),y0:0,y1:0\.91,line:\{color:'#257cb6',width:2\.5,dash:'dot'\}/g)||[]).length,2);
});
test('both capacity pages retain original saved models and ranges in the shared flow layout',()=>{
 for(const [page,prefix,model,chart] of [['availabilityPage','urgent','urgentDaysWeeksSwitch','availabilityChart'],['routineAvailabilityPage','routine','routineDaysWeeksSwitch','routineAvailabilityChart']]){
  const start=html.indexOf('<div id="'+page+'"'),end=html.indexOf('<div id="'+chart+'"',start),part=html.slice(start,end);
  assert.match(part,/analytics-chart-layout capacity-chart-layout/);assert.ok(part.indexOf('analytics-chart-heading')<part.indexOf('chart-display-options-panel'));
  assert.ok(part.indexOf('chart-display-options-panel')<part.indexOf('urgent-range-panel'));
  assert.match(part,new RegExp('data-chart-proxy="'+model+'" data-checked-value="week"'));
  assert.equal((html.match(new RegExp('id="'+model+'"','g'))||[]).length,1);
  assert.match(part,new RegExp('id="'+prefix+'DisplayOptions"[^>]* hidden'));
  assert.match(part,/Group unavailable appts/);assert.equal((part.match(/type="range"/g)||[]).length,2);
 }
});
