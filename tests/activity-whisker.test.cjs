const test=require('node:test'),assert=require('node:assert/strict'),core=require('../data-correctness.js');
test('two series receive symmetric twelve-pixel separation',()=>{
 const a=core.dodgeSeriesPosition(50,0,2,0,100,500),b=core.dodgeSeriesPosition(50,1,2,0,100,500);
 assert.ok(Math.abs((b-a)/100*500-12)<1e-9);assert.equal((a+b)/2,50);
});
test('seven series remain separated and inside either plot edge',()=>{
 for(const value of [0,100]){
  const points=Array.from({length:7},(_,i)=>core.dodgeSeriesPosition(value,i,7,0,100,160));
  assert.ok(points[0]/100*160>=6-1e-9);assert.ok(points.at(-1)/100*160<=154+1e-9);
  for(let i=1;i<points.length;i++)assert.ok(Math.abs((points[i]-points[i-1])/100*160-12)<1e-9);
 }
});
test('date-axis separation uses the same pixel distance and single series stays centred',()=>{
 const min=Date.parse('2023-01-01'),max=Date.parse('2026-10-31'),value=Date.parse('2026-10-26');
 const a=core.dodgeSeriesPosition(value,0,2,min,max,900),b=core.dodgeSeriesPosition(value,1,2,min,max,900);
 assert.ok(Math.abs((b-a)/(max-min)*900-12)<1e-9);
 assert.equal(core.dodgeSeriesPosition(50,0,1,0,100,500),50);
 assert.equal(core.dodgeSeriesPosition(50,0,2,0,0,500),50);
});
