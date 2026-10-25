(function(){
const $=id=>document.getElementById(id);
const yearStyles=document.createElement('style');yearStyles.textContent='.forecast-year-band{position:absolute;top:0;bottom:0;background:rgba(70,115,155,.035);border-left:1px solid #d0dae5;pointer-events:none;z-index:0}.forecast-year-alt{background:rgba(70,115,155,.1)}.forecast-year-band span{position:absolute;top:2px;left:8px;font-size:12px;font-weight:bold;color:#526f86}.forecast-history-column{position:relative;z-index:1}.forecast-season-line{z-index:3}.forecast-prediction-band{fill-opacity:.12}';yearStyles.textContent+='.forecast-stacked-bar{display:flex;flex-direction:column-reverse;overflow:hidden;background:transparent!important}.forecast-bar-segment{width:100%;flex-shrink:0;min-height:0} .forecast-y-axis{position:sticky;left:0;top:0;width:0;align-self:stretch;z-index:5;pointer-events:none}.forecast-y-tick{position:absolute;left:0;transform:translateY(-50%);display:block;min-width:34px;padding:2px 4px;background:rgba(255,255,255,.88);font-size:10px;text-align:right;color:#47647b;font-variant-numeric:tabular-nums}';document.head.append(yearStyles);
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
 // Weighted residual quantiles: indicative 70% prediction interval, widened for sparse data.
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
const max=Math.max(1,...rows.map(r=>r.high));for(const r of rows){const line=document.createElement('div');line.className='forecast-line';const label=document.createElement('span');label.textContent=new Date(r.day+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:mode==='daily'?'short':undefined,day:'numeric',month:'short',timeZone:'UTC'});const track=document.createElement('div');track.className='forecast-track';for(const p of r.parts){const seg=document.createElement('div');seg.style.width=(100*p.mean/max)+'%';seg.style.backgroundColor=urgencyGroupColour(p.n);seg.title=p.n+': '+p.mean.toFixed(1);track.append(seg)}const value=document.createElement('strong');value.textContent=Math.round(r.total).toLocaleString('en-GB');value.title='Indicative 70% prediction interval: '+Math.round(r.low)+'–'+Math.round(r.high);const bounds=document.createElement('span');bounds.className='forecast-bounds';bounds.textContent=Math.round(r.low)+'–'+Math.round(r.high);bounds.title='Indicative 70% prediction interval';const numbers=document.createElement('div');numbers.className='forecast-numbers';numbers.append(value,bounds);line.append(label,track,numbers);root.append(line)}
const note=document.createElement('p');note.className='muted';note.textContent='Historical booking data to '+end+'. '+(mode==='daily'?'Daily forecast combines recency-weighted weekday averages, a season-of-year adjustment where historical coverage permits, and a restrained recent trend. The smaller figures are indicative 70% prediction intervals, estimated from historical variation; these are not calibrated confidence guarantees.':'Weekly forecasts use complete historical Monday–Sunday weeks; the smaller figures show an indicative historical 10th–90th percentile range.')+' Zero-booking days included. Bank holidays and changes in service provision are not explicitly modelled.';root.append(note);
}
function renderDailyUrgencyFilters(){
 const root=$('dailyForecastFilters');if(!root)return;
 const names=prioritySort('urgency',groupUrgencyNames.filter(isDashboardToggleOption)).filter(isDaily);
 if(root.dataset.names===names.join('|'))return;
 root.dataset.names=names.join('|');root.replaceChildren();
 for(const name of names){
  const label=document.createElement('label');label.className='filter-row';
  const sw=document.createElement('span');sw.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.checked=!disabled.daily.has(name);
  const track=document.createElement('span');track.className='switch-track';track.style.setProperty('--switch-on',urgencyGroupColour(name));
  sw.append(input,track);label.append(sw,document.createTextNode(name));root.append(label);
  input.addEventListener('change',()=>{if(input.checked)disabled.daily.delete(name);else disabled.daily.add(name);renderHistoricalWeekday()});
 }
}
function renderHistoricalWeekday(){
 const root=$('dailyForecastHistoryChart'),summary=$('dailyForecastHistorySummary'),select=$('dailyForecastWeekday');
 if(!root||!summary||!select)return;
 root.replaceChildren();renderDailyUrgencyFilters();
 const weekday=Number(select.value),today=dayViewLocalDate(),end=shift(today,-1),start=shift(end,1-7*Number($('dailyForecastHistory')?.value||52));
 const groups=groupUrgencyNames.filter(n=>isDashboardToggleOption(n)&&isDaily(n)&&!disabled.daily.has(n));
 const adjust=$('dailyForecastAdjustListSize')?.checked===true;const populationMonths=adjust?practiceMonths():[];const referencePopulation=populationMonths.length?listSizeAtMonth(populationMonths[populationMonths.length-1]):null;
 const normalise=(day,count)=>adjust&&referencePopulation?count*referencePopulation/Math.max(1,listSizeAtMonth(day.slice(0,7))):count;
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
 const values=days.map(d=>normalise(d,totals.get(d)||0));
 const average=values.reduce((a,b)=>a+b,0)/values.length;
 const sd=Math.sqrt(values.reduce((a,b)=>a+(b-average)**2,0)/values.length);
 const n=days.length, origin=Date.parse(days[0]+'T12:00:00Z');
 const tOf=d=>(Date.parse(d+'T12:00:00Z')-origin)/86400000/365.2425;
 const features=(d,harmonic)=>{const t=tOf(d),v=[1,t];if(harmonic)v.push(Math.cos(2*Math.PI*t),Math.sin(2*Math.PI*t));return v};
 // Fit an annual harmonic plus linear trend where at least one year of observations is available.
 // Shorter windows use a trend-only model, avoiding a poorly identified seasonal cycle.
 function fitModel(xDays,y){
  const harmonic=xDays.length>=42,degree=harmonic?4:2;
  const m=Array.from({length:degree},()=>Array(degree+1).fill(0));
  for(let i=0;i<xDays.length;i++){const x=features(xDays[i],harmonic);for(let j=0;j<degree;j++){for(let k=0;k<degree;k++)m[j][k]+=x[j]*x[k];m[j][degree]+=x[j]*y[i]}}
  // Small ridge on the trend and harmonic coefficients stabilises short/noisy samples.
  for(let j=1;j<degree;j++)m[j][j]+=j===1?.02:.1;
  for(let i=0;i<degree;i++){let best=i;for(let j=i+1;j<degree;j++)if(Math.abs(m[j][i])>Math.abs(m[best][i]))best=j;[m[i],m[best]]=[m[best],m[i]];const pivot=m[i][i];if(Math.abs(pivot)<1e-9)return null;for(let k=i;k<=degree;k++)m[i][k]/=pivot;for(let j=0;j<degree;j++)if(j!==i){const factor=m[j][i];for(let k=i;k<=degree;k++)m[j][k]-=factor*m[i][k]}}
  const c=m.map(row=>row[degree]);return d=>Math.max(0,features(d,harmonic).reduce((sum,v,i)=>sum+v*c[i],0));
 }
 const fittedModel=n>=12?fitModel(days,values):null;
 const future=[];for(let i=1;i<=13;i++)future.push(shift(days[days.length-1],7*i));
 const historicalFit=fittedModel?days.map(d=>fittedModel(d)):[];
 const projected=fittedModel?future.map(d=>fittedModel(d)):[];
 // Expanding-window, forward-looking back-tests measure actual prediction errors.
 // Hold out successive observations up to 13 weeks ahead; do not use in-sample residuals.
 const backtestErrors=[];
 if(fittedModel&&n>=28){
  const minimum=Math.max(16,Math.min(52,Math.floor(n*.5)));
  for(let cut=minimum;cut<n;cut+=Math.max(1,Math.floor((n-minimum)/32))){
   const trained=fitModel(days.slice(0,cut),values.slice(0,cut));
   if(!trained)continue;
   for(let horizon=1;horizon<=13&&cut+horizon<=n;horizon+=3)backtestErrors.push({h:horizon,e:values[cut+horizon-1]-trained(days[cut+horizon-1])});
  }
 }
 const residuals=fittedModel?values.map((v,i)=>v-historicalFit[i]):[];
 const quantile=(items,q)=>{const sorted=[...items].sort((a,b)=>a-b);if(!sorted.length)return 0;const pos=(sorted.length-1)*q,lo=Math.floor(pos),hi=Math.ceil(pos);return sorted[lo]+(sorted[hi]-sorted[lo])*(pos-lo)};
 const empirical=backtestErrors.length>=16;
 const errors=empirical?backtestErrors.map(x=>x.e):residuals;
 const intervals=projected.map((value,i)=>{
  // Signed quantiles retain asymmetric uncertainty; inflate modestly for longer horizons.
  const factor=empirical?Math.sqrt((1+(i+1)/26)/(1+7/26)):Math.sqrt(1+(i+1)/13);
  const lo=quantile(errors,.15)*factor,hi=quantile(errors,.85)*factor;
  return {low:Math.max(0,value+Math.min(lo,0)),high:Math.max(value,value+Math.max(hi,0))};
 });
 // Shared scale across weekdays: inspect every weekday in the selected historical window.
 // Keep the same ceiling when changing Monday to Friday; do not rescale to the selected day's maximum.
 const allWeekdayTotals=new Map();
 for(const r of records){
  if(!r.bookedDay||r.bookedDay<start||r.bookedDay>end||excludedSlotTypes.has(r.slot))continue;
  if(!groups.includes(capacityUrgencyCategory(r.slot)))continue;
  if(excludeHolidayWeeks&&holidayWeeks.has(monday(r.bookedDay)))continue;
  const dow=new Date(r.bookedDay+'T12:00:00Z').getUTCDay();if(dow<1||dow>5)continue;
  allWeekdayTotals.set(r.bookedDay,(allWeekdayTotals.get(r.bookedDay)||0)+(Number(r.count)||0));
 }
 const commonMaximum=Math.max(1,...Array.from(allWeekdayTotals,( [d,v] )=>normalise(d,v)));
 const dataMaximum=Math.max(commonMaximum,...values,...historicalFit,...projected,...intervals.map(v=>v.high));
 const step=Math.pow(10,Math.floor(Math.log10(dataMaximum/4||1)));
 const increments=[1,2,5,10].map(v=>v*step);
 const tickStep=increments.find(v=>v*5>=dataMaximum)||10*step;
 const max=tickStep*5;
 const allDays=days.concat(future),columns=[];
 for(let i=0;i<allDays.length;i++){
  const d=allDays[i],isFuture=i>=n,value=isFuture?0:values[i],col=document.createElement('div');
  col.className='forecast-history-column'+(isFuture?' forecast-future-column':'');
  const bar=document.createElement('div');bar.className='forecast-history-bar forecast-stacked-bar';bar.style.height=(value/max*192)+'px';
  if(!isFuture&&value>0){
   const detail=parts.get(d)||new Map();
   for(const name of groups){
    const raw=detail.get(name)||0;if(raw<=0)continue;
    const segment=document.createElement('div');segment.className='forecast-bar-segment';
    segment.style.height=(raw/(totals.get(d)||1)*100)+'%';segment.style.backgroundColor=urgencyGroupColour(name);
    segment.title=name+': '+raw+' bookings';bar.append(segment);
   }
  }
  const date=document.createElement('div');date.className='forecast-history-date';date.textContent=new Date(d+'T12:00:00Z').toLocaleDateString('en-GB',{day:'numeric',month:'short',timeZone:'UTC'});
  if(isFuture){const j=i-n;col.title=d+': projected '+projected[j].toFixed(1)+' bookings (70% prediction interval '+intervals[j].low.toFixed(1)+'–'+intervals[j].high.toFixed(1)+')'}
  else {const detail=parts.get(d)||new Map();col.title=d+': '+value.toFixed(1)+(adjust?' list-size adjusted':'')+' bookings'+groups.map(name=>'\\n'+name+': '+(detail.get(name)||0)).join('')}
  col.append(bar,date);root.append(col);columns.push(col);
 }
 // A labelled Y-axis shares exactly the same baseline and scale as the bars and forecast.
 const barRect=columns[0].querySelector('.forecast-history-bar').getBoundingClientRect();
 const baselineY=barRect.bottom-root.getBoundingClientRect().top;
 const axis=document.createElement('div');axis.className='forecast-y-axis';
 for(let i=0;i<=5;i++){
  const label=document.createElement('span');label.className='forecast-y-tick';
  label.textContent=String(Math.round(i*tickStep*100)/100);
  label.style.top=(baselineY-i*tickStep/max*192)+'px';
  axis.append(label);
 }
 root.append(axis);
 // Alternating calendar-year bands use real year boundaries, including the projected weeks.
 const yearRuns=[];for(let i=0;i<allDays.length;i++){const year=allDays[i].slice(0,4);if(!yearRuns.length||yearRuns[yearRuns.length-1].year!==year)yearRuns.push({year,start:i,end:i});else yearRuns[yearRuns.length-1].end=i}
 for(let j=0;j<yearRuns.length;j++){
  const run=yearRuns[j],first=columns[run.start],last=columns[run.end];
  const band=document.createElement('div');band.className='forecast-year-band'+(j%2?' forecast-year-alt':'');
  band.style.left=first.offsetLeft+'px';band.style.width=(last.offsetLeft+last.offsetWidth-first.offsetLeft)+'px';
  const tag=document.createElement('span');tag.textContent=run.year;band.append(tag);root.append(band);
 }
 if(fittedModel){
  const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('class','forecast-season-line');svg.setAttribute('aria-label','Seasonal regression, thirteen-week projection and 70 percent prediction interval');
  const width=root.scrollWidth,height=root.clientHeight;svg.style.width=width+'px';svg.style.height=height+'px';svg.setAttribute('viewBox','0 0 '+width+' '+height);
  const ns='http://www.w3.org/2000/svg',x=i=>columns[i].offsetLeft+columns[i].offsetWidth/2;
  const firstBar=columns[0].querySelector('.forecast-history-bar');
  const baseline=firstBar.getBoundingClientRect().bottom-root.getBoundingClientRect().top;
  const y=v=>baseline-v/max*192;
  const poly=(points,cls)=>{const el=document.createElementNS(ns,'polyline');el.setAttribute('class',cls);el.setAttribute('points',points.map(p=>p.map(v=>v.toFixed(1)).join(',')).join(' '));svg.append(el)};
  poly(historicalFit.map((v,i)=>[x(i),y(v)]),'forecast-historical-fit');
  const boundary=historicalFit.length-1;
  poly([[x(boundary),y(historicalFit[boundary])],...projected.map((v,i)=>[x(n+i),y(v)])],'forecast-projection-line');
  const upperPoints=intervals.map((v,i)=>[x(n+i),y(v.high)]),lowerPoints=intervals.map((v,i)=>[x(n+i),y(v.low)]).reverse();
  const band=document.createElementNS(ns,'polygon');band.setAttribute('class','forecast-prediction-band');band.setAttribute('points',upperPoints.concat(lowerPoints).map(p=>p.map(v=>v.toFixed(1)).join(',')).join(' '));svg.insertBefore(band,svg.firstChild);
  root.append(svg);
 }
 summary.textContent=n+' observed '+select.options[select.selectedIndex].text+'s · '+start+' to '+end+' · Mean '+average.toFixed(1)+' · SD '+sd.toFixed(1)+'. Bars show actual demand; solid pink line is fitted '+(n>=42?'annual seasonal regression plus trend':'linear trend (insufficient observations for annual seasonality)')+'. Dashed pink line projects 13 weeks; shaded area is an indicative 70% prediction interval based on '+(empirical?'expanding-window historical back-testing':'historical residuals (limited data for back-testing)')+'.'+(excludeHolidayWeeks?' Bank holiday weeks excluded from fitting.':' Bank holiday weeks included.')+(adjust?(referencePopulation?' Historical counts adjusted to latest recorded list size.':' No list-size data available; showing raw counts.'):'')+' Future bank holidays and service changes are not separately modelled.';

}
function open(mode){for(const p of ['dailyForecastPage','weeklyForecastPage'])$(p)?.classList.toggle('hidden',p!==(mode==='daily'?'dailyForecastPage':'weeklyForecastPage'));for(const p of ['homePage','chartPage','waitingTimesPage','waitingOverTimePage','categoryPage','availabilityPage','routineAvailabilityPage','plannedAvailabilityPage','dayViewPage','bookingTimesPage','signpostingPage','peoplePage','practiceSettingsPage'])$(p)?.classList.add('hidden');document.querySelectorAll('.dashboard-nav button.active').forEach(b=>b.classList.remove('active'));$('nav'+(mode==='daily'?'Daily':'Weekly')+'Forecast')?.classList.add('active');if(mode==='daily')renderHistoricalWeekday();else render(mode)}
document.addEventListener('DOMContentLoaded',()=>{document.querySelectorAll('.dashboard-nav button').forEach(b=>{if(b.id!=='navDailyForecast'&&b.id!=='navWeeklyForecast'&&b.id!=='navForecastMenu')b.addEventListener('click',()=>{$('dailyForecastPage')?.classList.add('hidden');$('weeklyForecastPage')?.classList.add('hidden')})});$('dailyForecastWeekday')?.addEventListener('change',renderHistoricalWeekday);$('dailyForecastExcludeBankHolidayWeeks')?.addEventListener('change',renderHistoricalWeekday);$('dailyForecastAdjustListSize')?.addEventListener('change',renderHistoricalWeekday);$('dailyForecastHistory')?.addEventListener('change',renderHistoricalWeekday);for(const mode of ['daily','weekly']){$('nav'+(mode==='daily'?'Daily':'Weekly')+'Forecast')?.addEventListener('click',()=>open(mode));$(mode+'ForecastHistory')?.addEventListener('change',()=>{if(mode==='daily')renderHistoricalWeekday();else render(mode)})}});
})();