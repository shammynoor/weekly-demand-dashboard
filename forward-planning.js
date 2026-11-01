/* Routine Forward Planning: development-only live-data forecast.
   Read-only: consumes already loaded Slot Report aggregates, never writes OneDrive. */
(function(){
'use strict';
const $=id=>document.getElementById(id);
const today=()=>{const p=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date()).map(v=>[v.type,v.value]));return p.year+'-'+p.month+'-'+p.day};
const shift=(iso,n)=>{const d=new Date(iso+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)};
const workdays=(iso,n)=>{const out=[];for(let d=iso;out.length<n;d=shift(d,1)){const w=new Date(d+'T12:00:00Z').getUTCDay();if(w>0&&w<6)out.push(d)}return out};
const dayfmt=iso=>new Intl.DateTimeFormat('en-GB',{weekday:'short',day:'numeric',month:'short',timeZone:'UTC'}).format(new Date(iso+'T12:00:00Z'));
const state={clinician:'GP',horizon:7,selection:null};
const groupOf=r=>{const type=effectiveClinicianType(r.clinician,r.person);return clinicianGroupAssignments[type]||''};
const serviceOf=r=>{const group=groupOf(r),type=effectiveClinicianType(r.clinician,r.person);if(isDoctorPlanningGroup(group))return 'GP';const label=(group+' '+type+' '+r.clinician).toLowerCase();if(/physio|physiotherap|first.contact|\bfcp\b/.test(label))return 'Physio';if(/mental.health|psycholog|counsell|\bmh\b/.test(label))return 'Mental Health';return null};
const modeOf=r=>{const v=String(consultAssignments[r.slot]||suggestConsultMode(r.slot)||'').toLowerCase();if(/phone|telephone|remote.audio/.test(v))return 'phone';if(/face|f2f|in.person|in person/.test(v))return 'f2f';return null};
const valid=r=>r&&!excludedSlotTypes.has(r.slot)&&appointmentCalculationRole(capacityUrgencyCategory(r.slot))==='routine';
const matching=r=>valid(r)&&modeOf(r)&&(state.clinician==='All'||serviceOf(r)===state.clinician)&&serviceOf(r);
function data(){
 const available=typeof availabilityRecords==='undefined'?[]:availabilityRecords;
 const booked=typeof appointmentTimeRecords==='undefined'?[]:appointmentTimeRecords;
 if(!available.length)return {error:'No Slot Report availability data loaded. Sign in and refresh the OneDrive data.'};
 const now=today(),days=workdays(now,14);
 const pools={phone:[],f2f:[]},bookings={phone:new Map(),f2f:new Map()};
 for(const r of available){
  if(!matching(r)||String(r.availability||'').trim().toLowerCase()!=='available'||r.date<now)continue;
  pools[modeOf(r)].push({date:r.date,count:Number(r.count)||0});
 }
 // Booked date is reconstructed by the existing CSV loader from booked lead time.
 const historyStart=shift(now,-84);
 for(const r of booked){
  if(!matching(r)||!r.bookedDay||r.bookedDay>=now||r.bookedDay<historyStart)continue;
  const mode=modeOf(r),day=r.bookedDay;
  bookings[mode].set(day,(bookings[mode].get(day)||0)+(Number(r.count)||0));
 }
 const rows=[];
 for(const mode of ['phone','f2f']){
  const pool=pools[mode].sort((a,b)=>a.date.localeCompare(b.date)),counts=new Map();
  for(const r of pool)counts.set(r.date,(counts.get(r.date)||0)+r.count);
  let remaining=[...counts.values()].reduce((a,b)=>a+b,0),cumulativeUnmet=0;
  const weekdaySamples=Array.from({length:5},()=>[]);
  for(let d=historyStart;d<now;d=shift(d,1)){const w=new Date(d+'T12:00:00Z').getUTCDay();if(w>=1&&w<=5)weekdaySamples[w-1].push(bookings[mode].get(d)||0)}
  const daily=weekdaySamples.map(xs=>xs.length?xs.reduce((a,b)=>a+b,0)/xs.length:0);
  const samples=weekdaySamples.reduce((a,x)=>a+x.length,0);
  for(let i=0;i<days.length;i++){
   const date=days[i],weekday=new Date(date+'T12:00:00Z').getUTCDay()-1;
   const demand=Math.round(daily[weekday]||0),before=remaining;
   let toBook=Math.min(remaining,demand);
   // First available future appointment is booked first; past-date slots expire.
   for(const key of [...counts.keys()].sort()){
    if(key<date){remaining-=counts.get(key);counts.delete(key);continue}
    if(toBook<=0)break;
    const use=Math.min(toBook,counts.get(key));counts.set(key,counts.get(key)-use);remaining-=use;toBook-=use;
    if(!counts.get(key))counts.delete(key);
   }
   cumulativeUnmet+=Math.max(0,demand-(before- Math.max(0,before-remaining-toBook)));
   const earliest=[...counts.keys()].filter(k=>k>=date&&counts.get(k)>0).sort()[0];
   const wait=earliest?Math.max(0,Math.round((new Date(earliest+'T12:00:00Z')-new Date(date+'T12:00:00Z'))/86400000)):null;
   rows.push({i,date,mode,before,demand,remaining:Math.max(0,remaining),wait,samples,unmet:cumulativeUnmet});
  }
 }
 return {rows,days,totals:{phone:pools.phone.reduce((a,b)=>a+b.count,0),f2f:pools.f2f.reduce((a,b)=>a+b.count,0)}};
}
function render(){
 const root=$('forwardPlanningRows'),note=$('forwardPlanningNote'),detail=$('forwardPlanningDetail');if(!root)return;
 for(const [id,val] of [['forwardPlanningClinician',state.clinician],['forwardPlanningHorizon',String(state.horizon)]]){const e=$(id);if(e&&e.value!==val)e.value=val}
 let model;try{model=data()}catch(e){console.error('Forward planning',e);model={error:'Unable to calculate forecast from loaded Slot Report data.'}}
 if(model.error){root.replaceChildren();note.textContent=model.error;detail.textContent='';return}
 const fmt=n=>Math.round(n).toLocaleString('en-GB');
 const rows=model.rows.filter(r=>r.i<state.horizon).sort((a,b)=>a.i-b.i||(['phone','f2f'].indexOf(a.mode)-['phone','f2f'].indexOf(b.mode)));
 const full={};for(const mode of ['phone','f2f']){const index=model.rows.find(r=>r.mode===mode&&r.remaining===0)?.i;full[mode]=index===undefined?null:index}
 root.replaceChildren();
 for(const r of rows){
  const tr=document.createElement('tr');tr.tabIndex=0;tr.style.cursor='pointer';
  const fields=[r.mode==='phone'?dayfmt(r.date):'',r.mode==='phone'?'Phone':'F2F',fmt(r.before)+' / '+fmt(r.demand),r.remaining>0?'+'+fmt(r.remaining):'0',fmt(r.remaining),r.wait===null?'No capacity':(r.wait/7).toFixed(1)+' weeks',r.wait===null?'—':'—',full[r.mode]===null?'>14 working days':Math.max(0,full[r.mode]-r.i)+' working days'];
  fields.forEach((v,i)=>{const td=document.createElement('td');td.textContent=v;if(i>1)td.style.textAlign='center';if(i===5&&r.wait===null)td.style.color='#a22b39';tr.append(td)});
  tr.addEventListener('click',()=>{state.selection=r.date+'|'+r.mode;detail.textContent=dayfmt(r.date)+' · '+(r.mode==='phone'?'Phone':'F2F')+' — '+fmt(r.before)+' available at start, '+fmt(r.demand)+' forecast new bookings, '+fmt(r.remaining)+' left at day end.'});
  root.append(tr);
 }
 const history=model.rows.filter(r=>r.i===0).map(r=>r.mode+': '+r.samples+' historic weekdays').join('; ');
 note.textContent='Live Slot Report availability and historical booked appointments · '+history+'. Only Availability = Available is counted; Blocked, Booked and Embargoed are excluded. Forecast demand is an observed weekday average over 12 weeks, not a measure of unmet demand. Waiting time is the earliest remaining available appointment date after simulated bookings; it is not a validated prediction. “All” pools distinct services and is not a patient-specific waiting time.';
 if(!state.selection)detail.textContent='Select a row to inspect the forecast.';
}
window.renderForwardPlanning=render;
document.addEventListener('DOMContentLoaded',()=>{for(const [id,key] of [['forwardPlanningClinician','clinician'],['forwardPlanningHorizon','horizon']])$(id)?.addEventListener('change',e=>{state[key]=key==='horizon'?Number(e.target.value):e.target.value;state.selection=null;render()});});
})();
