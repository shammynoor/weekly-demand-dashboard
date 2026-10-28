/* Shared, DOM-free correctness rules. Also loadable by the Node regression suite. */
(function(root){
'use strict';
const shift=(iso,n)=>{const d=new Date(iso+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)};
const monday=iso=>shift(iso,-((new Date(iso+'T12:00:00Z').getUTCDay()+6)%7));
function londonDate(now=new Date()){
 const p=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(now).map(x=>[x.type,x.value]));
 return p.year+'-'+p.month+'-'+p.day;
}
function dateContext(now){const today=londonDate(now);return {today,week:monday(today),month:today.slice(0,7)}}
function validDate(value){return typeof value==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(value)&&!isNaN(Date.parse(value+'T12:00:00Z'))&&new Date(value+'T12:00:00Z').toISOString().slice(0,10)===value}
function coverageFromEntries(entries){
 const ranges=[];
 for(const entry of entries){
  const contribution=entry[1];if(contribution?.kind!=='slot')continue;
  let start='',end='';
  for(const [key] of contribution.maps[1]){
   const date=JSON.parse(key)[0];if(!validDate(date))continue;
   if(!start||date<start)start=date;if(date>end)end=date;
  }
  if(!start)continue;
  // Each cache version contains that file's modification time, not another file's.
  let asOf='';try{const value=JSON.parse(contribution.version)[2];if(value&&Number.isFinite(Date.parse(value)))asOf=londonDate(new Date(value));}catch(_){}
  ranges.push({start,end,asOf});
 }
 return ranges;
}
function completeWeeks(ranges,today,windowWeeks){
 const lastSunday=shift(monday(today),-1);
 const intervals=ranges.filter(r=>validDate(r.start)&&validDate(r.end)).map(r=>({
  start:r.start,end:[r.end,lastSunday,...(validDate(r.asOf)?[shift(r.asOf,-1)]:[])].sort()[0]
 })).filter(r=>r.start<=r.end).sort((a,b)=>a.start.localeCompare(b.start));
 const merged=[];
 for(const r of intervals){const prev=merged.at(-1);if(prev&&r.start<=shift(prev.end,1)){if(r.end>prev.end)prev.end=r.end}else merged.push({...r})}
 // The exports lack explicit coverage metadata. Use each file's unfiltered slot-date
 // span, capped at its snapshot date; never bridge an omitted file/date interval.
 const weeks=[];
 for(const range of merged){
  let day=monday(range.start);if(day<range.start)day=shift(day,7);
  for(;shift(day,4)<=range.end;day=shift(day,7))weeks.push(day);
 }
 if(!weeks.length)return [];
 const count=Math.max(1,Math.min(104,Number(windowWeeks)||52));
 const lower=shift(weeks.at(-1),-7*(count-1));
 return [...new Set(weeks.filter(w=>w>=lower))];
}
function roles(saved,appointmentNames,clinicianNames,categoryFor=()=> ''){
 const source=saved&&typeof saved==='object'?saved:{};
 const appointmentTypes=Object.create(null),clinicianGroups=Object.create(null);
 for(const name of appointmentNames){
  const old=Object.hasOwn(source.appointmentTypes||{},name)?source.appointmentTypes[name]:null;
  const fallback=/^routine(\b|$)/i.test(name.trim())?'routine':
   (/same.?day|urgent.?extra|within.?3.?day|^111.?referral$/i.test(name)||categoryFor(name)==='Urgent Care'?'urgent':'other');
  appointmentTypes[name]={id:typeof old?.id==='string'?old.id:'appointment:'+encodeURIComponent(name),kind:['urgent','routine','other'].includes(old?.kind)?old.kind:fallback};
 }
 for(const name of clinicianNames){
  const old=Object.hasOwn(source.clinicianGroups||{},name)?source.clinicianGroups[name]:null;
  clinicianGroups[name]={id:typeof old?.id==='string'?old.id:'clinician:'+encodeURIComponent(name),doctor:typeof old?.doctor==='boolean'?old.doctor:/^(doctors?|gps?|general practitioners?|medical practitioners?)$/i.test(name.trim())};
 }
 return {appointmentTypes,clinicianGroups};
}
function renameRole(metadata,dimension,oldName,newName){
 const map=metadata[dimension];if(!map||!Object.hasOwn(map,oldName))return;
 Object.defineProperty(map,newName,{value:map[oldName],writable:true,enumerable:true,configurable:true});if(oldName!==newName)delete map[oldName];
}
function smoothDemandSeries(points,{level=0,continuous=true,currentPeriod=''}={}){
 const kernels=[[1],[0.5,1,0.5],[1,1,1],[0.5,1,1,1,0.5]],kernel=kernels[Number(level)]||kernels[0],radius=(kernel.length-1)/2;
 const groups=new Map();
 for(const point of points){const key=continuous?'all':point.year;if(!groups.has(key))groups.set(key,[]);if(point.date!==currentPeriod)groups.get(key).push(point)}
 const values=new Map();
 for(const group of groups.values())for(let i=0;i<group.length;i++){
  let sum=0,weight=0;for(let offset=-radius;offset<=radius;offset++){const count=group[i+offset]?.count,w=kernel[offset+radius];if(Number.isFinite(count)){sum+=count*w;weight+=w}}
  values.set(group[i],weight?sum/weight:group[i].count);
 }
 return points.map(point=>({...point,smoothed:radius&&point.date!==currentPeriod?values.get(point):point.count}));
}
function periodBounds(key,mode){
 const start=mode==='month'?key.slice(0,7)+'-01':monday(key);
 const d=new Date(start+'T12:00:00Z');
 if(mode==='month')d.setUTCMonth(d.getUTCMonth()+1);else d.setUTCDate(d.getUTCDate()+7);
 return {start,end:shift(d.toISOString().slice(0,10),-1)};
}
function quantile(values,p){
 const sorted=values.filter(Number.isFinite).slice().sort((a,b)=>a-b);if(!sorted.length)return NaN;
 const at=(sorted.length-1)*p,lo=Math.floor(at);return sorted[lo]+(sorted[Math.ceil(at)]-sorted[lo])*(at-lo);
}
function currentPeriodNowcast({mode='week',currentDate,clock,daily=[],timed=[],coverage=[],excludedPeriods=[]}={}){
 if(!validDate(currentDate)||!clock||!validDate(clock.date)||!Number.isFinite(clock.minute))return {error:'A dated report snapshot is needed.'};
 const key=mode==='month'?currentDate.slice(0,7):monday(currentDate),bounds=periodBounds(key,mode);
 if(clock.date<bounds.start)return {error:'The latest report predates the current '+mode+'.'};
 if(clock.date>currentDate)return {error:'The report snapshot is in the future.'};
 const minute=Math.max(0,Math.min(1440,clock.minute)),dateCounts=new Map(),rawCounts=new Map();
 for(const row of daily){if(!validDate(row.date)||!Number.isFinite(row.count)||row.count<0)continue;dateCounts.set(row.date,(dateCounts.get(row.date)||0)+row.count);rawCounts.set(row.date,(rawCounts.get(row.date)||0)+(Number.isFinite(row.rawCount)?row.rawCount:row.count))}
 const ranges=coverage.filter(r=>validDate(r.start)&&validDate(r.end)).map(r=>({...r,end:[r.end,shift(bounds.start,-1),...(validDate(r.asOf)?[shift(r.asOf,-1)]:[])].sort()[0]})).filter(r=>r.start<=r.end);
 if(!ranges.length)return {error:'Report coverage is needed to identify complete history.'};
 const earliest=ranges.map(r=>r.start).sort()[0],oldest=mode==='month'?shift(bounds.start,-750):shift(bounds.start,-7*104);
 let cursor=mode==='month'?earliest.slice(0,7)+'-01':monday(earliest);
 const history=[],excluded=new Set(excludedPeriods);
 for(;cursor<bounds.start;){
  const period=periodBounds(cursor,mode),dates=[];for(let date=period.start;date<=period.end;date=shift(date,1))dates.push(date);
  const covered=dates.filter(date=>![0,6].includes(new Date(date+'T12:00:00Z').getUTCDay())).every(date=>ranges.some(r=>date>=r.start&&date<=r.end));
  if(cursor>=oldest&&covered&&!excluded.has(mode==='month'?cursor.slice(0,7):cursor))history.push({dates,total:dates.reduce((n,date)=>n+(dateCounts.get(date)||0),0)});
  cursor=shift(period.end,1);
 }
 const samples=history.slice(mode==='month'?-24:-104),minimum=mode==='month'?6:8;
 if(samples.length<minimum)return {error:'At least '+minimum+' complete covered '+(mode==='month'?'months':'weeks')+' are needed.'};
 const historyDates=new Set(samples.flatMap(p=>p.dates)),weekdaySum=Array(7).fill(0),weekdayDays=Array(7).fill(0),timeTotal=Array(7).fill(0),timeElapsed=Array(7).fill(0);
 for(const date of historyDates){const weekday=new Date(date+'T12:00:00Z').getUTCDay();weekdaySum[weekday]+=dateCounts.get(date)||0;weekdayDays[weekday]++}
 const weekdayWeight=weekdaySum.map((n,i)=>weekdayDays[i]?n/weekdayDays[i]:0);
 if(!weekdaySum.some(n=>n>0))return {error:'No historical bookings match these filters.'};
 for(const row of timed){
  if(!historyDates.has(row.date)||!Number.isFinite(row.minute)||row.minute<0||row.minute>=1440||!Number.isFinite(row.count)||row.count<0)continue;
  const weekday=new Date(row.date+'T12:00:00Z').getUTCDay();timeTotal[weekday]+=row.count;if(row.minute<=minute)timeElapsed[weekday]+=row.count;
 }
 const weekday=new Date(clock.date+'T12:00:00Z').getUTCDay(),timedCoverage=weekdaySum[weekday]>0?Math.min(1,timeTotal[weekday]/weekdaySum[weekday]):0;
 const intradayProfile=timedCoverage>=0.9&&timeTotal[weekday]>0;
 const withinDay=intradayProfile?timeElapsed[weekday]/timeTotal[weekday]:minute/1440;
 const dates=[];for(let date=bounds.start;date<=bounds.end;date=shift(date,1))dates.push(date);
 const weight=date=>weekdayWeight[new Date(date+'T12:00:00Z').getUTCDay()];
 const fullWeight=dates.reduce((n,date)=>n+weight(date),0);
 if(!(fullWeight>0))return {error:'No comparable booking pattern is available.'};
 const elapsed=dates.reduce((n,date)=>n+weight(date)*(date<clock.date?1:date===clock.date?withinDay:0),0);
 const fraction=Math.max(0,Math.min(1,elapsed/fullWeight));
 const observed=dates.filter(date=>date<=clock.date).reduce((n,date)=>n+(dateCounts.get(date)||0),0);
 const rawObserved=dates.filter(date=>date<=clock.date).reduce((n,date)=>n+(rawCounts.get(date)||0),0);
 const typical=quantile(samples.map(p=>p.total),0.5),priorExposure=Math.min(0.25,1/dates.length),prior=Math.max(1,typical*priorExposure);
 const candidates=[];
 for(const sample of samples){
  let exposure=fraction*sample.dates.reduce((n,date)=>n+weight(date),0),partial=0;
  for(const date of sample.dates){const w=weight(date);if(w<=0)continue;const share=Math.max(0,Math.min(1,exposure/w));partial+=(dateCounts.get(date)||0)*share;exposure-=w;}
  const scale=(observed+prior)/(partial+prior);
  candidates.push(observed+Math.max(0,sample.total-partial)*scale);
 }
 const estimate=Math.max(observed,quantile(candidates,0.5));
 const remaining=Math.max(0,estimate-observed),samplePenalty=Math.max(1,Math.sqrt(12/samples.length));
 // Empirical analogue spread plus a conservative uncertainty floor. This is an
 // indicative finishing range, not a calibrated statistical confidence guarantee.
 const spread=Math.max(1.645*Math.sqrt(remaining),estimate*(intradayProfile?0.25:0.35)*Math.sqrt(1-fraction))*samplePenalty;
 const low=Math.max(observed,Math.min(quantile(candidates,0.05),estimate-spread));
 const high=Math.max(low,quantile(candidates,0.95),estimate+spread);
 return {key,mode,estimate,low:Math.max(observed,Math.floor(low)),high:Math.ceil(high),observed,rawObserved,fraction,samples:samples.length,intradayProfile,timedCoverage,asOf:{date:clock.date,minute},periodStart:bounds.start,periodEnd:bounds.end};
}

// Nudge coincident current-period columns in pixels on the existing time axis.
// The caller reuses each position for the line endpoint, observation and range.
function currentPeriodPositions(value,range,width,count){
 const n=Number.isFinite(count)?Math.max(0,Math.floor(count)):0;
 const date=typeof value==='string',number=v=>date?Date.parse(v):v;
 const start=number(range[0]),end=number(range[1]),centre=number(value);
 if(!n)return [];
 if(!Number.isFinite(width)||width<=0||!(end>start))return Array(n).fill(value);
 const spacing=n>1?Math.min(6,24/(n-1),width/(n+1)):0;
 const half=(n-1)*spacing/2,scale=(end-start)/width;
 const anchor=Math.max(start+half*scale,Math.min(end-half*scale,centre));
 return Array.from({length:n},(_,i)=>{
  const position=anchor+(i-(n-1)/2)*spacing*scale;
  return date?new Date(position).toISOString():position;
 });
}

const api={currentPeriodPositions,currentPeriodNowcast,periodBounds,smoothDemandSeries,shift,monday,londonDate,dateContext,validDate,coverageFromEntries,completeWeeks,roles,renameRole};
if(typeof module!=='undefined'&&module.exports)module.exports=api;
else root.GPCorrectness=api;
})(typeof window!=='undefined'?window:globalThis);
