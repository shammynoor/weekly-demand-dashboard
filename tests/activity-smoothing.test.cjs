const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),core=require('../data-correctness.js');
test('smoothing off preserves values and does not mutate source points',()=>{
 const points=[{date:'2026-01',year:2026,count:0},{date:'2026-02',year:2026,count:90},{date:'2026-03',year:2026,count:0}],copy=JSON.stringify(points);
 assert.deepEqual(core.smoothDemandSeries(points,{level:0}).map(p=>p.smoothed),[0,90,0]);assert.equal(JSON.stringify(points),copy);
});
test('Low/Medium/High progressively smooth a spike using gentler weighted kernels',()=>{
 const points=Array.from({length:11},(_,i)=>({date:'2026-'+String(i+1).padStart(2,'0'),year:2026,count:i===5?90:0}));
 assert.equal(core.smoothDemandSeries(points,{level:1})[5].smoothed,45);
 assert.equal(core.smoothDemandSeries(points,{level:2})[5].smoothed,30);
 assert.equal(core.smoothDemandSeries(points,{level:3})[5].smoothed,22.5);
 for(const level of [1,2,3])assert.ok(core.smoothDemandSeries(points.map(p=>({...p,count:40})),{level}).every(p=>p.smoothed===40));
});
test('partial current period stays raw and does not lower completed history',()=>{
 const points=[{date:'2026-01',year:2026,count:60},{date:'2026-02',year:2026,count:60},{date:'2026-03',year:2026,count:3}];
 assert.deepEqual(core.smoothDemandSeries(points,{level:3,currentPeriod:'2026-03'}).map(p=>p.smoothed),[60,60,3]);
});
test('continuous history crosses year boundaries; superimposed history keeps years separate',()=>{
 const points=[{date:'2025-12',year:2025,count:20},{date:'2026-01',year:2026,count:100}];
 assert.deepEqual(core.smoothDemandSeries(points,{level:1,continuous:true}).map(p=>p.smoothed),[70/1.5,110/1.5]);
 assert.deepEqual(core.smoothDemandSeries(points,{level:1,continuous:false}).map(p=>p.smoothed),[20,100]);
});
test('direct appointment types remain resolvable without a redundant self mapping',()=>{
 const source=fs.readFileSync(__dirname+'/../app.js','utf8'),begin=source.indexOf('function capacityUrgencyCategory('),end=source.indexOf('function renderUrgentCategoryToggles(',begin);
 const ctx={assignments:{'Phone slot':'Routine','Legacy slot':'Detailed routine'},groupUrgencyNames:['Routine'],groupUrgencyAssignments:{'Detailed routine':'Routine'}};
 vm.createContext(ctx);vm.runInContext(source.slice(begin,end),ctx);
 assert.equal(ctx.capacityUrgencyCategory('Phone slot'),'Routine');assert.equal(ctx.capacityUrgencyCategory('Legacy slot'),'Routine');
 const refresh=source.slice(source.indexOf('function refresh(){'),source.indexOf('function installDemandCardTooltip('));
 assert.ok(refresh.includes('const group=capacityUrgencyCategory(r.slot)'));
 assert.ok(refresh.indexOf('applyDemandPurposeDefaults()')<refresh.indexOf('for(const r of records)'));
});

test('new High is gentler than the former five-period Medium',()=>{const points=Array.from({length:11},(_,i)=>({date:'2026-'+String(i+1).padStart(2,'0'),year:2026,count:i===5?90:0}));const high=core.smoothDemandSeries(points,{level:3})[5].smoothed,medium=core.smoothDemandSeries(points,{level:2})[5].smoothed;assert.ok(high>90/5);assert.ok(high<medium)});
