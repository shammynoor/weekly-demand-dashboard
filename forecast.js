(function(){
const $=id=>document.getElementById(id);
const shift=(date,n)=>{const d=new Date(date+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)};
const monday=d=>shift(d,-((new Date(d+'T12:00:00Z').getUTCDay()+6)%7));
const isDaily=n=>/same.?day|urgent.?extra|within.?3.?day/i.test(n);
const disabled={daily:new Set(),weekly:new Set()};
function render(mode){
const prefix=mode==='daily'?'dailyForecast':'weeklyForecast',root=$(prefix+'Chart'),filter=$(prefix+'Filters');if(!root||!filter)return;
const names=prioritySort('urgency',groupUrgencyNames.filter(isDashboardToggleOption)).filter(n=>isDaily(n)===(mode==='daily'));
if(filter.dataset.names!==names.join('|')){filter.dataset.names=names.join('|');filter.replaceChildren();for(const n of names){const row=document.createElement('label');row.className='filter-row';const sw=document.createElement('span');sw.className='switch';const inp=document.createElement('input');inp.type='checkbox';inp.checked=!disabled[mode].has(n);const track=document.createElement('span');track.className='switch-track';track.style.setProperty('--switch-on',urgencyGroupColour(n));sw.append(inp,track);row.append(sw,document.createTextNode(n));filter.append(row);inp.addEventListener('change',()=>{inp.checked?disabled[mode].delete(n):disabled[mode].add(n);render(mode)})}}
root.replaceChildren();if(!records.length){root.textContent='Load appointment data to generate a forecast.';return}
const selected=names.filter(n=>!disabled[mode].has(n));if(!selected.length){root.textContent='Select at least one urgency group.';return}
const today=dayViewLocalDate(),latest=records.reduce((x,r)=>r.bookedDay&&r.bookedDay>x?r.bookedDay:x,'');
const end=latest&&latest<today?latest:shift(today,-1),weeks=Number($(prefix+'History').value),start=shift(end,1-weeks*7);
const byDate=new Map(),allDates=new Set();
for(const r of records){if(!r.bookedDay||r.bookedDay>end||excludedSlotTypes.has(r.slot))continue;allDates.add(r.bookedDay);const n=capacityUrgencyCategory(r.slot);if(!selected.includes(n))continue;const k=r.bookedDay+'|'+n;byDate.set(k,(byDate.get(k)||0)+r.count)}
const earliest=[...allDates].sort()[0]||start;
const dates=[];for(let d=earliest;d<=end;d=shift(d,1))dates.push(d);
const recentDays=dates.filter(d=>d>=start);
const fullWeeks=[];for(let d=monday(start);d<=end;d=shift(d,7))if(d>=start&&shift(d,6)<=end)fullWeeks.push(d);
const dayOfWeek=d=>new Date(d+'T12:00:00Z').getUTCDay();
const yearPosition=d=>{const dt=new Date(d+'T12:00:00Z');const year=dt.getUTCFullYear();return (Date.UTC(year,dt.getUTCMonth(),dt.getUTCDate())-Date.UTC(year,0,1))/86400000};
const seasonDistance=(a,b)=>{const diff=Math.abs(yearPosition(a)-yearPosition(b));return Math.min(diff,365-diff)};
const ageDays=d=>(Date.parse(end+'T12:00:00Z')-Date.parse(d+'T12:00:00Z'))/86400000;
const weightedMean=(values,weights)=>{const sum=weights.reduce((a,b)=>a+b,0);return sum?values.reduce((a,v,i)=>a+v*weights[i],0)/sum:0};
const weightedQuantile=(values,weights,q)=>{const pairs=values.map((v,i)=>[v,weights[i]]).sort((a,b)=>a[0]-b[0]);const total=weights.reduce((a,b)=>a+b,0);let cumulative=0;for(const [v,w] of pairs){cumulative+=w;if(cumulative>=q*total)return v}return pairs.length?pairs[pairs.length-1][0]:0};
const predictCategory=(target,n)=>{
 const weekday=dayOfWeek(target);
 const candidates=dates.filter(d=>d>=start&&dayOfWeek(d)===weekday);
 const values=candidates.map(d=>byDate.get(d+'|'+n)||0);
 if(!values.length)return {n,mean:0,low:0,high:0,samples:0};
 const recentWeights=candidates.map(d=>Math.exp(-ageDays(d)/Math.max(35,weeks*3.5)));
 const baseline=weightedMean(values,recentWeights);
 if(mode!=='daily'){
  const weekly=fullWeeks.map(w=>Array.from({length:7},(_,j)=>byDate.get(shift(w,j)+'|'+n)||0).reduce((a,b)=>a+b,0));
  const mean=weekly.length?weekly.reduce((a,b)=>a+b,0)/weekly.length:0;
  const sorted=[...weekly].sort((a,b)=>a-b);
  return {n,mean,low:sorted[Math.floor((sorted.length-1)*.1)]||0,high:sorted[Math.ceil((sorted.length-1)*.9)]||0,samples:weekly.length};
 }
 // Seasonal factor uses same-weekday observations from all available years, including prior years.
 // A 45-day seasonal bandwidth avoids spurious month-boundary jumps.
 const seasonalDays=dates.filter(d=>dayOfWeek(d)===weekday);
 const seasonalValues=seasonalDays.map(d=>byDate.get(d+'|'+n)||0);
 const seasonalWeights=seasonalDays.map(d=>Math.exp(-.5*(seasonDistance(d,target)/45)**2));
 const seasonalEffective=(seasonalWeights.reduce((a,b)=>a+b,0)**2)/(seasonalWeights.reduce((a,b)=>a+b*b,0)||1);
 const seasonMean=weightedMean(seasonalValues,seasonalWeights);
 const globalMean=seasonalValues.reduce((a,b)=>a+b,0)/seasonalValues.length;
 // Shrink the seasonal adjustment toward neutral when comparable historical data are sparse.
 const seasonalStrength=Math.min(.65,Math.max(0,(seasonalEffective-3)/(seasonalEffective+8)));
 const factor=globalMean>0?Math.max(.65,Math.min(1.5,1+seasonalStrength*(seasonMean/globalMean-1))):1;
 // A modest trend term compares the newer and older halves of the recent window.
 const half=Math.floor(candidates.length/2),old=values.slice(0,half),newer=values.slice(half);
 const olderAvg=old.length?old.reduce((a,b)=>a+b,0)/old.length:baseline;
 const newerAvg=newer.length?newer.reduce((a,b)=>a+b,0)/newer.length:baseline;
 const trend=olderAvg>0?Math.max(.85,Math.min(1.15,1+.25*(newerAvg/olderAvg-1))):1;
 const mean=Math.max(0,baseline*factor*trend);
 // Weighted residual quantiles: indicative 80% prediction interval, widened for sparse data.
 const residuals=values.map(v=>v-baseline);
 const lo=weightedQuantile(residuals,recentWeights,.1),hi=weightedQuantile(residuals,recentWeights,.9);
 const spread=Math.sqrt(Math.max(mean,1))*1.28;
 const low=Math.max(0,mean+Math.min(lo,-spread)*Math.max(1,Math.sqrt(8/values.length)));
 const high=Math.max(mean,mean+Math.max(hi,spread)*Math.max(1,Math.sqrt(8/values.length)));
 return {n,mean,low,high,samples:values.length,seasonFactor:factor};
};
const rows=[];for(let i=0;i<(mode==='daily'?14:8);i++){
 const day=mode==='daily'?shift(end,i+1):shift(monday(end),7*(i+1));
 const parts=selected.map(n=>predictCategory(day,n));
 rows.push({day,parts,total:parts.reduce((x,p)=>x+p.mean,0),low:parts.reduce((x,p)=>x+p.low,0),high:parts.reduce((x,p)=>x+p.high,0)});
}
const legend=document.createElement('div');legend.className='forecast-legend';for(const n of selected){const item=document.createElement('span');const dot=document.createElement('i');dot.className='forecast-dot';dot.style.backgroundColor=urgencyGroupColour(n);item.append(dot,document.createTextNode(n));legend.append(item)}root.append(legend);
const max=Math.max(1,...rows.map(r=>r.high));for(const r of rows){const line=document.createElement('div');line.className='forecast-line';const label=document.createElement('span');label.textContent=new Date(r.day+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:mode==='daily'?'short':undefined,day:'numeric',month:'short',timeZone:'UTC'});const track=document.createElement('div');track.className='forecast-track';for(const p of r.parts){const seg=document.createElement('div');seg.style.width=(100*p.mean/max)+'%';seg.style.backgroundColor=urgencyGroupColour(p.n);seg.title=p.n+': '+p.mean.toFixed(1);track.append(seg)}const value=document.createElement('strong');value.textContent=Math.round(r.total).toLocaleString('en-GB');value.title='Indicative 80% prediction interval: '+Math.round(r.low)+'–'+Math.round(r.high);const bounds=document.createElement('span');bounds.className='forecast-bounds';bounds.textContent=Math.round(r.low)+'–'+Math.round(r.high);bounds.title='Indicative 80% prediction interval';const numbers=document.createElement('div');numbers.className='forecast-numbers';numbers.append(value,bounds);line.append(label,track,numbers);root.append(line)}
const note=document.createElement('p');note.className='muted';note.textContent='Historical booking data to '+end+'. '+(mode==='daily'?'Daily forecast combines recency-weighted weekday averages, a season-of-year adjustment where historical coverage permits, and a restrained recent trend. The smaller figures are indicative 80% prediction intervals, estimated from historical variation; these are not calibrated confidence guarantees.':'Weekly forecasts use complete historical Monday–Sunday weeks; the smaller figures show an indicative historical 10th–90th percentile range.')+' Zero-booking days included. Bank holidays and changes in service provision are not explicitly modelled.';root.append(note);
}
function renderHistoricalWeekday(){
 const root=$('dailyForecastHistoryChart'),summary=$('dailyForecastHistorySummary'),select=$('dailyForecastWeekday');
 if(!root||!summary||!select)return;
 root.replaceChildren();
 const weekday=Number(select.value),today=dayViewLocalDate(),end=shift(today,-1),start=shift(end,-363);
 const groups=groupUrgencyNames.filter(n=>isDashboardToggleOption(n)&&isDaily(n));
 const totals=new Map(),parts=new Map();
 for(const r of records){
  if(!r.bookedDay||r.bookedDay<start||r.bookedDay>end||excludedSlotTypes.has(r.slot))continue;
  const urgency=capacityUrgencyCategory(r.slot);
  if(!groups.includes(urgency))continue;
  const k=r.bookedDay;
  totals.set(k,(totals.get(k)||0)+(Number(r.count)||0));
  if(!parts.has(k))parts.set(k,new Map());
  const detail=parts.get(k);detail.set(urgency,(detail.get(urgency)||0)+(Number(r.count)||0));
 }
 const excludeHolidayWeeks=$('dailyForecastExcludeBankHolidayWeeks')?.checked!==false;
 const holidayWeeks=new Set();if(excludeHolidayWeeks){for(let year=Number(start.slice(0,4))-1;year<=Number(end.slice(0,4))+1;year++)for(const holiday of englandWalesBankHolidays(year))holidayWeeks.add(monday(holiday))}
 const days=[];let skippedHolidayDays=0;for(let d=start;d<=end;d=shift(d,1))if(new Date(d+'T12:00:00Z').getUTCDay()===weekday){if(excludeHolidayWeeks&&holidayWeeks.has(monday(d))){skippedHolidayDays++;continue}days.push(d)}
 if(!records.length||!groups.length||!days.length){summary.textContent=!days.length?'No matching weekdays remain after excluding bank holiday weeks.':'No historical booking data or matching urgency categories are available.';return}
 const values=days.map(d=>totals.get(d)||0),average=values.reduce((a,b)=>a+b,0)/values.length,sd=Math.sqrt(values.reduce((a,b)=>a+(b-average)**2,0)/values.length),lower=Math.max(0,average-sd),upper=average+sd;const max=Math.max(1,upper,...values);
 for(const d of days){
  const value=totals.get(d)||0,col=document.createElement('div');col.className='forecast-history-column';
  const bar=document.createElement('div');bar.className='forecast-history-bar';bar.style.height=(value/max*192)+'px';
  const date=document.createElement('div');date.className='forecast-history-date';date.textContent=new Date(d+'T12:00:00Z').toLocaleDateString('en-GB',{day:'numeric',month:'short',timeZone:'UTC'});
  const detail=parts.get(d)||new Map();
  col.title=d+': '+value+' new bookings'+groups.map(n=>'\n'+n+': '+(detail.get(n)||0)).join('');
  col.append(bar,date);root.append(col);
 }
 for(const [name,value] of [['Mean + 1 SD',upper],['Mean',average],['Mean - 1 SD',lower]]){const line=document.createElement('div');line.className='forecast-reference-line';line.style.bottom=(51+value/max*192)+'px';const label=document.createElement('span');label.textContent=name+': '+value.toFixed(1);line.append(label);root.append(line)}
 summary.textContent=days.length+' '+select.options[select.selectedIndex].text+'s · '+start+' to '+end+' · Mean '+average.toFixed(1)+' · SD '+sd.toFixed(1)+' · Minimum '+Math.min(...values)+' · Maximum '+Math.max(...values)+'. Each bar shows bookings made on that date across Same Day, Within 3 Days and Urgent Extra; zero-booking days are included.'+(excludeHolidayWeeks?' Excluded '+skippedHolidayDays+' dates in bank holiday weeks.':' Bank holiday weeks included.');
}
function open(mode){for(const p of ['dailyForecastPage','weeklyForecastPage'])$(p)?.classList.toggle('hidden',p!==(mode==='daily'?'dailyForecastPage':'weeklyForecastPage'));for(const p of ['homePage','chartPage','waitingTimesPage','waitingOverTimePage','categoryPage','availabilityPage','routineAvailabilityPage','plannedAvailabilityPage','dayViewPage','bookingTimesPage','signpostingPage','peoplePage','practiceSettingsPage'])$(p)?.classList.add('hidden');document.querySelectorAll('.dashboard-nav button.active').forEach(b=>b.classList.remove('active'));$('nav'+(mode==='daily'?'Daily':'Weekly')+'Forecast')?.classList.add('active');render(mode);if(mode==='daily')renderHistoricalWeekday()}
document.addEventListener('DOMContentLoaded',()=>{document.querySelectorAll('.dashboard-nav button').forEach(b=>{if(b.id!=='navDailyForecast'&&b.id!=='navWeeklyForecast'&&b.id!=='navForecastMenu')b.addEventListener('click',()=>{$('dailyForecastPage')?.classList.add('hidden');$('weeklyForecastPage')?.classList.add('hidden')})});$('dailyForecastWeekday')?.addEventListener('change',renderHistoricalWeekday);$('dailyForecastExcludeBankHolidayWeeks')?.addEventListener('change',renderHistoricalWeekday);for(const mode of ['daily','weekly']){$('nav'+(mode==='daily'?'Daily':'Weekly')+'Forecast')?.addEventListener('click',()=>open(mode));$(mode+'ForecastHistory')?.addEventListener('change',()=>render(mode))}});
})();