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
const api={shift,monday,londonDate,dateContext,validDate,coverageFromEntries,completeWeeks,roles,renameRole};
if(typeof module!=='undefined'&&module.exports)module.exports=api;
else root.GPCorrectness=api;
})(typeof window!=='undefined'?window:globalThis);
