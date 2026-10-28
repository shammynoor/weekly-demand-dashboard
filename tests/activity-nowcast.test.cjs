const test=require('node:test'),assert=require('node:assert/strict'),core=require('../data-correctness.js');
function fixture(mode='week'){
 const daily=[],timed=[],end=mode==='month'?'2026-09-30':'2026-10-25',start=mode==='month'?'2024-10-01':'2026-05-04';
 for(let date=start;date<=end;date=core.shift(date,1)){
  const weekday=new Date(date+'T12:00:00Z').getUTCDay(),count=weekday===0||weekday===6?0:100;
  daily.push({date,count,rawCount:count});if(count)for(const minute of [540,900])timed.push({date,minute,count:50});
 }
 const currentStart=mode==='month'?'2026-10-01':'2026-10-26';
 for(let date=currentStart;date<='2026-10-28';date=core.shift(date,1)){
  const weekday=new Date(date+'T12:00:00Z').getUTCDay(),count=weekday===0||weekday===6?0:date==='2026-10-28'?50:100;daily.push({date,count,rawCount:count});
 }
 return {mode,currentDate:'2026-10-28',clock:{date:'2026-10-28',minute:720},daily,timed,coverage:[{start,end:'2026-10-28',asOf:'2026-10-28'}]};
}
test('weekly nowcast estimates the finish from weekday and time-of-day progress',()=>{
 const result=core.currentPeriodNowcast(fixture());assert.equal(result.observed,250);assert.equal(result.estimate,500);
 assert.equal(result.fraction,0.5);assert.equal(result.intradayProfile,true);assert.ok(result.low<500&&result.high>500);assert.ok(result.low>=250);assert.ok(result.samples>=8);
});
test('monthly range accounts for the month calendar and includes expected full-month bookings',()=>{
 const result=core.currentPeriodNowcast(fixture('month'));assert.ok(!result.error,result.error);assert.ok(result.samples>=6);
 let expected=0;for(let date='2026-10-01';date<='2026-10-31';date=core.shift(date,1))if(![0,6].includes(new Date(date+'T12:00:00Z').getUTCDay()))expected+=100;
 assert.ok(result.low<=expected&&result.high>=expected);assert.ok(result.estimate>=result.observed);
});
test('an identical observed count implies a lower finishing estimate later in the day',()=>{
 const input=fixture(),early=core.currentPeriodNowcast({...input,clock:{date:'2026-10-28',minute:540}}),late=core.currentPeriodNowcast({...input,clock:{date:'2026-10-28',minute:960}});
 assert.ok(early.estimate>late.estimate);assert.ok(early.high-early.low>late.high-late.low);
});
test('zero bookings early in the period still produces a broad historical range',()=>{
 const input=fixture();input.clock={date:'2026-10-26',minute:480};input.daily=input.daily.filter(r=>r.date<'2026-10-26');const result=core.currentPeriodNowcast(input);
 assert.equal(result.observed,0);assert.ok(result.estimate>0);assert.ok(result.high>result.estimate&&result.low<result.estimate);
});
test('unavailable booking times use a labelled wider fallback',()=>{
 const input=fixture(),profile=core.currentPeriodNowcast(input),fallback=core.currentPeriodNowcast({...input,timed:[]});
 assert.equal(fallback.intradayProfile,false);assert.ok(fallback.high-fallback.low>profile.high-profile.low);
});
test('stale snapshot is used instead of the wall clock; prior-period reports are rejected',()=>{
 const input=fixture(),result=core.currentPeriodNowcast({...input,clock:{date:'2026-10-27',minute:1200}});
 assert.equal(result.asOf.date,'2026-10-27');assert.equal(result.observed,200);assert.match(core.currentPeriodNowcast({...input,clock:{date:'2026-10-23',minute:1200}}).error,/predates/);
 assert.match(core.currentPeriodNowcast({...input,clock:null}).error,/snapshot/);
});
test('coverage gaps and excluded bank-holiday weeks do not become zero-booking history',()=>{
 const input=fixture(),full=core.currentPeriodNowcast(input);
 const missing=core.currentPeriodNowcast({...input,coverage:[{start:'2026-05-04',end:'2026-06-05'},{start:'2026-06-15',end:'2026-10-28',asOf:'2026-10-28'}]});
 assert.equal(missing.samples,full.samples-1);
 const excluded=core.currentPeriodNowcast({...input,excludedPeriods:['2026-06-08']});assert.equal(excluded.samples,full.samples-1);
 assert.match(core.currentPeriodNowcast({...input,coverage:[{start:'2026-10-01',end:'2026-10-28',asOf:'2026-10-28'}]}).error,/8 complete/);
});
test('range respects the bookings already observed and does not mutate input',()=>{
 const input=fixture();input.daily=input.daily.map(r=>({...r,count:r.count*1.2345}));input.timed=input.timed.map(r=>({...r,count:r.count*1.2345}));const before=JSON.stringify(input),result=core.currentPeriodNowcast(input);
 assert.ok(result.low>=result.observed);assert.ok(result.high>=result.estimate);assert.equal(JSON.stringify(input),before);assert.equal(result.rawObserved,250);
});
