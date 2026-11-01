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
 const note=$('forwardPlanningNote'),summary=$('forwardPlanningSummary'),host=$('forwardPlanningCombinedChart');
 if(!host)return;
 for(const [id,val] of [['forwardPlanningClinician',state.clinician],['forwardPlanningHorizon',String(state.horizon)]]){const el=$(id);if(el&&el.value!==val)el.value=val}
 let model;try{model=data()}catch(e){console.error('Forward planning',e);model={error:'Unable to calculate forecast from loaded data.'}}
 host.replaceChildren();
 if(model.error){note.textContent=model.error;summary.textContent='';return}
 const modes=['phone','f2f'],series=modes.map(mode=>({mode,rows:model.rows.filter(r=>r.mode===mode&&r.i<state.horizon)}));
 const dates=series[0].rows.map(r=>r.date),width=960,height=330,p={l:66,r:70,t:32,b:50};
 const fmt=n=>Math.round(n).toLocaleString('en-GB');
 const capacityMax=Math.max(1,...series.flatMap(s=>s.rows.map(r=>r.before)));
 // At the first exhausted point, extend the recent observed waiting-time slope.
 // This is an illustrative extrapolation, not a bookable appointment or validated wait.
 const projected={};
 for(const item of series){
  const idx=item.rows.findIndex(r=>r.remaining===0);
  if(idx<0)continue;
  const known=item.rows.slice(0,idx).map((r,i)=>({i,v:r.wait===null?null:r.wait/7})).filter(p=>p.v!==null);
  if(!known.length)continue;
  const last=known.at(-1),previous=known.at(-2);
  const slope=previous?Math.max(-0.5,Math.min(1.5,(last.v-previous.v)/Math.max(1,last.i-previous.i))):0;
  projected[item.mode]={index:idx,value:Math.max(0,last.v+slope*(idx-last.i))};
 }
 const waitMax=Math.max(1,...series.flatMap(s=>s.rows.map(r=>r.wait===null?0:r.wait/7)),...Object.values(projected).map(p=>p.value));
 const x=i=>p.l+i/Math.max(1,dates.length-1)*(width-p.l-p.r);
 const yCap=v=>height-p.b-v/capacityMax*(height-p.t-p.b);
 const yWait=v=>height-p.b-v/waitMax*(height-p.t-p.b);
 const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('viewBox','0 0 '+width+' '+height);svg.style.width='100%';svg.setAttribute('role','img');svg.setAttribute('aria-label','Routine forward planning: Phone and Face-to-Face appointments remaining and waiting times');
 const add=(tag,attrs,label)=>{const el=document.createElementNS(svg.namespaceURI,tag);for(const [k,v] of Object.entries(attrs))el.setAttribute(k,v);if(label!==undefined)el.textContent=label;svg.append(el);return el};
 for(let i=0;i<=4;i++){const yy=p.t+i*(height-p.t-p.b)/4;add('line',{x1:p.l,y1:yy,x2:width-p.r,y2:yy,stroke:'#e1e8ef'});add('text',{x:p.l-10,y:yy+4,'text-anchor':'end','font-size':12,fill:'#64778a'},fmt(capacityMax*(4-i)/4));add('text',{x:width-p.r+10,y:yy+4,'font-size':12,fill:'#64778a'},(waitMax*(4-i)/4).toFixed(1)) }
 add('text',{x:p.l,y:17,'font-size':13,fill:'#4d6378'},'Appointments remaining');
 add('text',{x:width-p.r,y:17,'text-anchor':'end','font-size':13,fill:'#4d6378'},'Waiting time (weeks)');
 dates.forEach((d,i)=>{if(i%Math.max(1,Math.ceil(dates.length/7))===0||i===dates.length-1)add('text',{x:x(i),y:height-17,'text-anchor':'middle','font-size':12,fill:'#64778a'},dayfmt(d).replace(/^\\w+ /,''))});
 const colors={phone:'#267ca5',f2f:'#ad7042'};
 function line(values,color,dashed,mode,label,scale){
  let segment=[];
  function flush(){if(segment.length){if(segment.length>1)add('polyline',{points:segment.join(' '),fill:'none',stroke:color,'stroke-width':dashed?2.5:3.5,'stroke-dasharray':dashed?'7 5':'none','stroke-linejoin':'round','stroke-linecap':'round'});else{const [cx,cy]=segment[0].split(',');add('circle',{cx,cy,r:3,fill:color})}}segment=[]}
  values.forEach((v,i)=>{if(v===null){flush();return}segment.push(x(i)+','+scale(v))});flush();
  values.forEach((v,i)=>{if(v===null)return;const hit=add('circle',{cx:x(i),cy:scale(v),r:7,fill:color,opacity:0.01});const title=document.createElementNS(svg.namespaceURI,'title');title.textContent=dayfmt(dates[i])+' · '+mode+' · '+label+': '+(dashed?v.toFixed(1)+' weeks':fmt(v));hit.append(title)})
 }
 for(const item of series){const color=colors[item.mode];line(item.rows.map(r=>r.remaining),color,false,item.mode,'appointments remaining',yCap);line(item.rows.map(r=>r.wait===null?null:r.wait/7),color,true,item.mode,'waiting time',yWait)}
 // Capsule labels for exhaustion, and a separately marked extrapolated wait.
 const capsule=(cx,cy,label,color,above)=>{
  const w=Math.max(122,label.length*6.6+27),h=27;
  const xx=Math.max(p.l,Math.min(width-p.r-w,cx-w/2));
  const yy=above?Math.max(p.t,cy-h-13):Math.min(height-p.b-h-8,cy+10);
  add('rect',{x:xx,y:yy,width:w,height:h,rx:13.5,fill:'#fff',stroke:'#d8e1e8','stroke-width':1.2,opacity:0.97});
  add('path',{d:'M '+(xx+11)+' '+(yy+8)+' L '+(xx+17)+' '+(yy+19)+' L '+(xx+5)+' '+(yy+19)+' Z',fill:'#d78b30'});
  add('text',{x:xx+11,y:yy+17.4,'text-anchor':'middle','font-size':9,'font-weight':700,fill:'#fff'},'!');
  add('text',{x:xx+24,y:yy+18,'font-size':12,'font-weight':600,fill:color},label);
 };
 for(const item of series){
  const exhausted=item.rows.find(r=>r.remaining===0);
  if(!exhausted)continue;
  const i=item.rows.indexOf(exhausted),color=colors[item.mode];
  add('circle',{cx:x(i),cy:yCap(0),r:5,fill:color});
  capsule(x(i),yCap(0),(item.mode==='phone'?'Phone':'F2F')+' runs out',color,item.mode==='phone');
  const proj=projected[item.mode];
  if(proj){
   const known=item.rows.map((r,j)=>({j,v:r.wait===null?null:r.wait/7})).filter(p=>p.j<i&&p.v!==null).at(-1);
   if(known)add('line',{x1:x(known.j),y1:yWait(known.v),x2:x(i),y2:yWait(proj.value),stroke:color,'stroke-width':2,'stroke-dasharray':'2 5',opacity:0.75});
   add('circle',{cx:x(i),cy:yWait(proj.value),r:5,fill:'#fff',stroke:color,'stroke-width':2});
   const label=proj.value.toFixed(1)+'w est.';
   const labelY=yWait(proj.value)+(item.mode==='phone'?-13:19);
   add('text',{x:Math.max(p.l+25,Math.min(width-p.r-25,x(i))),y:Math.max(p.t+14,Math.min(height-p.b-8,labelY)),'text-anchor':'middle','font-size':11,'font-weight':600,fill:color},label);
  }
 }
 // Display the wait (weeks) at every modelled point, with small offsets for readability.
 for(const item of series){
  const color=colors[item.mode];
  for(let i=0;i<item.rows.length;i++){
   const r=item.rows[i];if(r.wait===null)continue;
   const xx=x(i),yy=yWait(r.wait/7);
   add('circle',{cx:xx,cy:yy,r:2.7,fill:color});
   add('text',{x:xx,y:Math.max(p.t+10,Math.min(height-p.b-5,yy+(item.mode==='phone'?-12:16))),'text-anchor':'middle','font-size':10,'font-weight':600,fill:color,'paint-order':'stroke',stroke:'#fff','stroke-width':3,'stroke-linejoin':'round'},(r.wait/7).toFixed(1)+'w');
  }
 }
 host.append(svg);
 const summaries=series.map(item=>{const empty=item.rows.find(r=>r.remaining===0);return (item.mode==='phone'?'Phone':'F2F')+': '+(empty?'runs out '+dayfmt(empty.date):'capacity beyond '+dayfmt(item.rows.at(-1).date))});
 summary.textContent=summaries.join('   ·   ');
 note.textContent='Solid lines = appointments remaining (left axis); dashed lines = predicted waiting time in weeks (right axis). Phone is blue; Face-to-Face is orange. Waiting-time labels show weeks at every point. At the first exhausted point, an open-circle estimated wait extends the last observed slope; this is an illustrative projection only, not an available appointment. Waiting-time lines otherwise stop when no appointment remains. Uses live Slot Report data, only Availability = Available; excludes Blocked, Booked and Embargoed. Demand is estimated from the last 12 weeks of bookings and may understate unmet demand. Estimates assume earliest available slots are booked first.';
}
window.renderForwardPlanning=render;
document.addEventListener('DOMContentLoaded',()=>{for(const [id,key] of [['forwardPlanningClinician','clinician'],['forwardPlanningHorizon','horizon']])$(id)?.addEventListener('change',e=>{state[key]=key==='horizon'?Number(e.target.value):e.target.value;state.selection=null;render()});});
})();
