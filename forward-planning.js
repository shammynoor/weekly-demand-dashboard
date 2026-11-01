/* Routine Forward Planning: development-only live-data forecast.
   Read-only: consumes already loaded Slot Report aggregates, never writes OneDrive. */
(function(){
'use strict';
const $=id=>document.getElementById(id);
const today=()=>{const p=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date()).map(v=>[v.type,v.value]));return p.year+'-'+p.month+'-'+p.day};
const shift=(iso,n)=>{const d=new Date(iso+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)};
const workdays=(iso,n)=>{const out=[];for(let d=iso;out.length<n;d=shift(d,1)){const w=new Date(d+'T12:00:00Z').getUTCDay();if(w>0&&w<6)out.push(d)}return out};
const dayfmt=iso=>new Intl.DateTimeFormat('en-GB',{weekday:'short',day:'numeric',month:'short',timeZone:'UTC'}).format(new Date(iso+'T12:00:00Z'));
const state={clinician:'GP',horizon:14,mode:'phone',selection:null};
const groupOf=r=>{const type=effectiveClinicianType(r.clinician,r.person);return clinicianGroupAssignments[type]||''};
const serviceOf=r=>{const group=groupOf(r),type=effectiveClinicianType(r.clinician,r.person);if(isDoctorPlanningGroup(group))return 'GP';const label=(group+' '+type+' '+r.clinician).toLowerCase();if(/physio|physiotherap|first.contact|\bfcp\b/.test(label))return 'Physio';if(/mental.health|psycholog|counsell|\bmh\b/.test(label))return 'Mental Health';return null};
const modeOf=r=>{const v=String(consultAssignments[r.slot]||suggestConsultMode(r.slot)||'').toLowerCase();if(/phone|telephone|remote.audio/.test(v))return 'phone';if(/face|f2f|in.person|in person/.test(v))return 'f2f';return null};
const valid=r=>r&&!excludedSlotTypes.has(r.slot)&&appointmentCalculationRole(capacityUrgencyCategory(r.slot))==='routine';
const matching=r=>valid(r)&&modeOf(r)&&serviceOf(r)&&(state.clinician==='All'||serviceOf(r)===state.clinician);
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
  let remaining=[...counts.values()].reduce((a,b)=>a+b,0);
  const weekdaySamples=Array.from({length:5},()=>[]);
  for(let d=historyStart;d<now;d=shift(d,1)){const w=new Date(d+'T12:00:00Z').getUTCDay();if(w>=1&&w<=5)weekdaySamples[w-1].push(bookings[mode].get(d)||0)}
  const daily=weekdaySamples.map(xs=>xs.length?xs.reduce((a,b)=>a+b,0)/xs.length:0);
  const samples=weekdaySamples.reduce((a,x)=>a+x.length,0);
  for(let i=0;i<days.length;i++){
   const date=days[i],weekday=new Date(date+'T12:00:00Z').getUTCDay()-1;
   const demand=Math.round(daily[weekday]||0);
   // Expired unbooked appointment slots cannot be sold in later days.
   for(const key of [...counts.keys()])if(key<date){remaining-=counts.get(key);counts.delete(key)}
   const before=remaining;
   let toBook=Math.min(remaining,demand);
   // Allocate new bookings to the earliest future available appointment.
   for(const key of [...counts.keys()].sort()){
    if(toBook<=0)break;
    const use=Math.min(toBook,counts.get(key));counts.set(key,counts.get(key)-use);remaining-=use;toBook-=use;
    if(!counts.get(key))counts.delete(key);
   }
   const earliest=[...counts.keys()].filter(k=>k>=date&&counts.get(k)>0).sort()[0];
   const wait=earliest?Math.max(0,Math.round((new Date(earliest+'T12:00:00Z')-new Date(date+'T12:00:00Z'))/86400000)):null;
   rows.push({i,date,mode,before,demand,remaining:Math.max(0,remaining),wait,samples,unmet:Math.max(0,demand-before)});
  }
 }
 return {rows,days,totals:{phone:pools.phone.reduce((a,b)=>a+b.count,0),f2f:pools.f2f.reduce((a,b)=>a+b.count,0)}};
}
function render(){
 const note=$('forwardPlanningNote'),summary=$('forwardPlanningSummary'),capacity=$('forwardPlanningCapacityChart'),waiting=$('forwardPlanningWaitChart');
 if(!capacity||!waiting)return;
 for(const [id,val] of [['forwardPlanningClinician',state.clinician],['forwardPlanningHorizon',String(state.horizon)],['forwardPlanningMode',state.mode]]){const el=$(id);if(el&&el.value!==val)el.value=val}
 let model;try{model=data()}catch(e){console.error('Forward planning',e);model={error:'Unable to calculate the forecast from the loaded data.'}}
 if(model.error){note.textContent=model.error;summary.textContent='';capacity.replaceChildren();waiting.replaceChildren();return}
 const rows=model.rows.filter(r=>r.mode===state.mode&&r.i<state.horizon);
 const fmt=n=>Math.round(n).toLocaleString('en-GB');
 const exhausted=rows.find(r=>r.remaining===0);
 const initial=rows[0]?.before||0,latest=rows.at(-1)?.remaining||0;
 summary.textContent=exhausted?'Appointments forecast to run out on '+dayfmt(exhausted.date):'Appointments available beyond '+dayfmt(rows.at(-1).date);
 function chart(host,values,opts){
  host.replaceChildren();
  const width=900,height=230,p={l:58,r:24,t:22,b:44};
  const max=Math.max(opts.floor||0,...values.filter(v=>v!==null).map(v=>v||0),1);
  const x=i=>p.l+(rows.length<=1?0:i/(rows.length-1))*(width-p.l-p.r);
  const y=v=>height-p.b-(v/max)*(height-p.t-p.b);
  const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
  const add=(tag,attrs,text)=>{const el=document.createElementNS(svg.namespaceURI,tag);for(const [k,v] of Object.entries(attrs))el.setAttribute(k,v);if(text!==undefined)el.textContent=text;svg.append(el);return el};
  svg.setAttribute('viewBox','0 0 '+width+' '+height);svg.style.width='100%';svg.style.height='auto';svg.setAttribute('role','img');svg.setAttribute('aria-label',opts.title);
  for(let i=0;i<=4;i++){const val=max*i/4,yy=y(val);add('line',{x1:p.l,y1:yy,x2:width-p.r,y2:yy,stroke:'#e2e9ef','stroke-width':1});add('text',{x:p.l-10,y:yy+4,'text-anchor':'end',fill:'#6a7b8d','font-size':12},opts.tick(val))}
  rows.forEach((r,i)=>{if(i%Math.max(1,Math.ceil(rows.length/6))===0||i===rows.length-1)add('text',{x:x(i),y:height-15,'text-anchor':'middle',fill:'#607386','font-size':12},dayfmt(r.date).replace(/^\\w+ /,''))});
  let segment=[];
  function flush(){if(segment.length>1)add('polyline',{points:segment.join(' '),fill:'none',stroke:opts.color,'stroke-width':3,'stroke-linejoin':'round','stroke-linecap':'round'});else if(segment.length===1){const [cx,cy]=segment[0].split(',');add('circle',{cx,cy,r:3,fill:opts.color})}segment=[]}
  values.forEach((v,i)=>{if(v===null){flush();return}segment.push(x(i)+','+y(v))});flush();
  if(opts.markExhausted&&exhausted){const i=rows.indexOf(exhausted);add('line',{x1:x(i),y1:p.t,x2:x(i),y2:height-p.b,stroke:'#ba4352','stroke-width':1.5,'stroke-dasharray':'5 4'});add('text',{x:Math.min(x(i)+8,width-125),y:p.t+13,fill:'#a33342','font-size':12},'Runs out')}
  values.forEach((v,i)=>{if(v===null)return;const dot=add('circle',{cx:x(i),cy:y(v),r:5,fill:opts.color,opacity:0.01});const title=document.createElementNS(svg.namespaceURI,'title');title.textContent=dayfmt(rows[i].date)+': '+opts.tick(v);dot.append(title)});
  host.append(svg);
 }
 chart(capacity,rows.map(r=>r.remaining),{title:'Available appointments remaining by day',color:'#227fa4',tick:v=>fmt(v),markExhausted:true});
 chart(waiting,rows.map(r=>r.wait===null?null:r.wait/7),{title:'Predicted earliest appointment waiting time in weeks',color:'#8b67b3',tick:v=>v.toFixed(1)+'w'});
 const current=rows[0]?.wait,final=rows.at(-1)?.wait;
 $('forwardPlanningCapacityCaption').textContent=fmt(initial)+' currently available in the future diary · '+fmt(latest)+' projected remaining at the end of the period';
 $('forwardPlanningWaitCaption').textContent=current===null?'No available appointment to calculate a starting wait':('Earliest available appointment: '+(current/7).toFixed(1)+' weeks initially'+(final===null?' · no capacity at period end':' → '+(final/7).toFixed(1)+' weeks at period end'));
 note.textContent='Based on loaded Slot Report data. Only Availability = Available counts; blocked, booked and embargoed appointments are excluded. Forecast demand uses historical weekday booking averages, not unmet demand. Waiting time estimates assume bookings use the earliest available appointment; no capacity means the waiting time cannot be estimated.';
}
window.renderForwardPlanning=render;
document.addEventListener('DOMContentLoaded',()=>{for(const [id,key] of [['forwardPlanningClinician','clinician'],['forwardPlanningHorizon','horizon'],['forwardPlanningMode','mode']])$(id)?.addEventListener('change',e=>{state[key]=key==='horizon'?Number(e.target.value):e.target.value;state.selection=null;render()});});
})();
