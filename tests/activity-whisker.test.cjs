const test=require('node:test'),assert=require('node:assert/strict'),core=require('../data-correctness.js');
function pixels(layout,width){return layout.positions.map(x=>(layout.domain[0]+x*(layout.domain[1]-layout.domain[0]))*width)}
test('two forecasts use ten-pixel spacing entirely after the historical plot',()=>{
 const width=800,l=core.predictionLaneLayout(2,width),points=pixels(l,width);
 assert.ok(l.domain[0]>l.historyEnd);assert.ok(points[0]-4>l.historyEnd*width);
 assert.ok(Math.abs(points[1]-points[0]-10)<1e-9);
});
test('seven forecasts remain separated and visible on a smaller plot',()=>{
 const width=160,l=core.predictionLaneLayout(7,width),points=pixels(l,width);
 assert.ok(points[0]-4>l.historyEnd*width);assert.ok(points.at(-1)+4<width);
 for(let i=1;i<points.length;i++)assert.ok(Math.abs(points[i]-points[i-1]-10)<1e-9);
});
test('resize preserves marker spacing; no current period restores full history width',()=>{
 for(const width of [160,450,900]){const points=pixels(core.predictionLaneLayout(2,width),width);assert.ok(Math.abs(points[1]-points[0]-10)<1e-9)}
 const none=core.predictionLaneLayout(0,800);assert.equal(none.historyEnd,1);assert.deepEqual(none.positions,[]);
 const one=core.predictionLaneLayout(1,800);assert.deepEqual(one.positions,[0.5]);assert.ok(one.domain[0]>one.historyEnd);
});
