"use strict";
const CLIENT_ID='ba2bc895-4213-447c-ba62-d7edec6e8173';
const SCOPES=['Files.Read','Files.ReadWrite.AppFolder'];
const REDIRECT_URI='https://shammynoor.github.io/weekly-demand-dashboard/index.html';
const GRAPH='https://graph.microsoft.com/v1.0';
const FOLDER_KEY='weekly-demand-onedrive-folder-v1';
let records=[],weeks=[],types=[],statuses=[];
const slots={selectedOptions:[],replaceChildren(){},append(){},options:[]};
const $=id=>document.getElementById(id);
let auth,activeAccount=null,selectedFolder=null,folderView={id:'root',name:'OneDrive',trail:[]},busy=false,lastFingerprint='';
let cloudReady=false, cloudSaveTimer=null, cloudVersion=null, cloudBusy=false;
let cloudFolder=null;
const CLOUD_SETTINGS_NAME='weekly-demand-categories.json';
function cloudNotice(msg){const el=$('cloudCategoryStatus');if(el)el.textContent=msg}
async function cloudFile(){const t=await token();const r=await fetch(GRAPH+'/me/drive/special/approot:/'+CLOUD_SETTINGS_NAME+':/content',{headers:{Authorization:'Bearer '+t}});if(r.status===404)return null;if(!r.ok)throw Error('Settings read failed (HTTP '+r.status+')');return r.json()}
function validCategoryNames(names){
 if(!Array.isArray(names))throw Error('Saved categories are invalid');
 const result=[];
 for(const item of names){
  if(typeof item!=='string')continue;
  const name=item.trim();
  if(name&&name.toLowerCase()!=='uncategorised'&&!result.some(x=>x.toLowerCase()===name.toLowerCase()))result.push(name);
 }
 return result;
}
function validAssignments(values,names){
 if(!values||typeof values!=='object'||Array.isArray(values))throw Error('Saved assignments are invalid');
 const result={};
 for(const [slot,category] of Object.entries(values)){
  if(typeof category==='string'&&(category==='Uncategorised'||names.includes(category)))result[slot]=category;
 }
 return result;
}
async function loadCloudCategories(){
 if(!activeAccount)return;
 cloudReady=false;
 cloudNotice('Loading all categories from OneDrive…');
 try{
  const data=await cloudFile();
  if(data){
   cloudFolder=data.folder&&typeof data.folder.id==='string'?data.folder:null;
   const names=validCategoryNames(data.categories);
   const remoteAssignments=validAssignments(data.assignments,names);
   // Older dashboard versions only saved Urgent and Routine to OneDrive.
   // On the original computer, preserve additional local categories during migration.
   const legacy=Number(data.version||1)<2;
   const localNames=categoryNames.slice();
   const localAssignments={...assignments};
   categoryNames=names.length?names:['Urgent','Routine'];
   assignments=remoteAssignments;
   if(legacy&&localNames.some(c=>!categoryNames.includes(c))){
    for(const c of localNames)if(!categoryNames.includes(c))categoryNames.push(c);
    for(const [slot,cat] of Object.entries(localAssignments)){
     if(!Object.prototype.hasOwnProperty.call(assignments,slot)||!['Urgent','Routine','Uncategorised'].includes(cat)){
      if(categoryNames.includes(cat)||cat==='Uncategorised')assignments[slot]=cat;
     }
    }
    cloudNotice('Recovered browser categories; saving all categories to OneDrive…');
    cloudReady=true;
    saveCategories(false);
    await saveCloudCategories();
   }else cloudNotice('All categories loaded from OneDrive.');
  }else cloudNotice('No cloud settings yet. Existing browser categories will be saved when you make a change.');
  cloudReady=true;
  saveCategories(false);
  renderEditor();
  refresh();
 }catch(e){
  cloudNotice('Could not load cloud categories: '+e.message+' (local settings retained)');
  cloudReady=false;
 }
}
let cloudSaveRevision=0;
function scheduleCloudSave(){
 if(!cloudReady||!activeAccount)return;
 cloudSaveRevision++;
 clearTimeout(cloudSaveTimer);
 cloudNotice('Saving all category changes to OneDrive…');
 cloudSaveTimer=setTimeout(()=>{void saveCloudCategories()},700);
}
async function saveCloudCategories(){
 if(!cloudReady||!activeAccount)return;
 if(cloudBusy)return;
 cloudBusy=true;
 try{
  do{
   const revision=cloudSaveRevision;
   const snapshot=JSON.stringify({
    version:2,
    categories:categoryNames.slice(),
    assignments:{...assignments},
    folder:selectedFolder||cloudFolder||null,
    updatedAt:new Date().toISOString()
   });
   const t=await token();
   const response=await fetch(GRAPH+'/me/drive/special/approot:/'+CLOUD_SETTINGS_NAME+':/content',{
    method:'PUT',
    headers:{Authorization:'Bearer '+t,'Content-Type':'application/json'},
    body:snapshot
   });
   if(!response.ok)throw Error('HTTP '+response.status);
   lastSavedRevision=revision;
   if(revision===cloudSaveRevision){
    cloudNotice('All categories saved to OneDrive · '+new Date().toLocaleTimeString('en-GB'));
   }
  }while(cloudSaveRevision!==revisionAtLastSaved());
 }catch(e){cloudNotice('Category cloud save failed: '+e.message)}
 finally{cloudBusy=false}
}
let lastSavedRevision=-1;
function revisionAtLastSaved(){return lastSavedRevision}
function setStatus(s){$('connectionStatus').textContent=s;$('loginStatus').textContent=s}
function setSync(s){$('syncStatus').textContent=s}
function setButtons(){const logged=!!activeAccount;$('loginScreen').classList.toggle('hidden',logged);$('dashboardShell').classList.toggle('hidden',!logged);$('signIn').disabled=logged;$('signOut').disabled=!logged;$('chooseFolder').disabled=!logged;$('refreshData').disabled=!logged||!selectedFolder||busy}
async function initAuth(){
 try{
  if(!window.msal)throw Error('Microsoft sign-in library did not load. Please refresh with Ctrl+F5 or check browser content blocking.');
  if(!window.Plotly||!window.Papa)throw Error('Chart or CSV library did not load. Please refresh with Ctrl+F5.');
  auth=new msal.PublicClientApplication({auth:{clientId:CLIENT_ID,authority:'https://login.microsoftonline.com/consumers',redirectUri:REDIRECT_URI,navigateToLoginRequestUrl:false},cache:{cacheLocation:'sessionStorage'}});
  if(typeof auth.initialize==='function')await auth.initialize();const result=await auth.handleRedirectPromise();
  activeAccount=result?.account||auth.getAllAccounts()[0]||null;
  if(activeAccount){
   auth.setActiveAccount(activeAccount);
   try{selectedFolder=JSON.parse(localStorage.getItem(FOLDER_KEY)||'null')}catch{}
   await loadCloudCategories();
   // Existing browser migrates its chosen folder to OneDrive settings once.
   if(cloudFolder){selectedFolder=cloudFolder;localStorage.setItem(FOLDER_KEY,JSON.stringify(selectedFolder))}
   else if(selectedFolder&&cloudReady){cloudFolder=selectedFolder;await saveCloudCategories()}
   setStatus('Signed in: '+(activeAccount.username||'Microsoft account'));
   $('accountLabel').textContent=activeAccount.username||'Microsoft account';
   if(selectedFolder){$('folderStatus').textContent='Source: '+selectedFolder.name;await loadFolder(false)}
   else {setSync('No source folder saved yet. Use Manage categories → Source folder setup once on the computer where you previously selected it.');$('sourceSetup').classList.remove('hidden')}
  }else setStatus('Sign in with the Hotmail account that owns your OneDrive files.');
 }catch(e){setStatus('Sign-in setup error: '+e.message)}finally{setButtons()}
}
async function token(){if(!activeAccount)throw Error('Please sign in first.');try{return (await auth.acquireTokenSilent({scopes:SCOPES,account:activeAccount})).accessToken}catch(e){if(e instanceof msal.InteractionRequiredAuthError){await auth.acquireTokenRedirect({scopes:SCOPES,account:activeAccount});throw Error('Microsoft sign-in required')}throw e}}
async function graph(url){const bearer=await token();const response=await fetch(url.startsWith('https://')?url:GRAPH+url,{headers:{Authorization:'Bearer '+bearer}});if(!response.ok){let message='HTTP '+response.status;try{const body=await response.json();message=body.error?.message||message}catch{}throw Error(message)}return response.json()}
async function children(id){let url=id==='root'?'/me/drive/root/children?$top=200&$select=id,name,folder,file,size,lastModifiedDateTime':`/me/drive/items/${encodeURIComponent(id)}/children?$top=200&$select=id,name,folder,file,size,lastModifiedDateTime`;const items=[];while(url){const r=await graph(url);items.push(...(r.value||[]));url=r['@odata.nextLink']||null}return items}
async function showFolderPicker(){$('onedrivePanel').classList.remove('hidden');folderView={id:'root',name:'OneDrive',trail:[]};$('folderPicker').classList.remove('hidden');await showFolderContents()}
async function showFolderContents(){try{$('pickerPath').textContent=[...folderView.trail.map(x=>x.name),folderView.name].join(' / ');$('pickerEntries').textContent='Loading folders…';const folders=(await children(folderView.id)).filter(x=>x.folder).sort((a,b)=>a.name.localeCompare(b.name));$('pickerEntries').replaceChildren();if(!folders.length)$('pickerEntries').textContent='No subfolders. You can select this folder.';for(const f of folders){const btn=document.createElement('button');btn.className='folder-entry';btn.textContent='📁 '+f.name;btn.onclick=async()=>{folderView={id:f.id,name:f.name,trail:[...folderView.trail,{id:folderView.id,name:folderView.name}]};await showFolderContents()};$('pickerEntries').append(btn)}}catch(e){$('pickerEntries').textContent='Could not list folders: '+e.message}}
function weekMonday(d){const x=new Date(d);x.setUTCHours(12,0,0,0);x.setUTCDate(x.getUTCDate()-((x.getUTCDay()+6)%7));return x.toISOString().slice(0,10)}
function parseDate(s){const m=/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec((s||'').trim());if(!m)return null;const mi=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].toLowerCase());if(mi<0)return null;const d=new Date(Date.UTC(Number(m[3]),mi,Number(m[1])));return d.getUTCMonth()===mi&&d.getUTCDate()===Number(m[1])?d:null}
function aggregateCSV(csv,aggregated){const parsed=Papa.parse(csv,{header:true,skipEmptyLines:'greedy'});const fields=parsed.meta.fields||[];for(const required of ['Appointment Date','Appointment Time','Booked Time to Slot Time','Slot Type','Availability'])if(!fields.includes(required))throw Error('Missing column: '+required);let day='',time='',valid=0,skipped=0;for(const r of parsed.data){if(r['Appointment Date']?.trim())day=r['Appointment Date'].trim();if(r['Appointment Time']?.trim())time=r['Appointment Time'].trim();if((r['Availability']||'').trim().toLowerCase()!=='booked')continue;const lead=Number((r['Booked Time to Slot Time']||'').replace(/,/g,''));const dt=parseDate(day);const tm=/^(\d{1,2}):(\d{2})/.exec(time);if(!dt||!tm||!Number.isFinite(lead)||!(r['Booked Time to Slot Time']||'').trim()){skipped++;continue}dt.setUTCHours(Number(tm[1]),Number(tm[2]),0,0);const booked=new Date(dt.getTime()-lead*60000);if(!Number.isFinite(booked.getTime())){skipped++;continue}const w=weekMonday(booked),slot=(r['Slot Type']||'').trim()||'Unknown';if(w<'2025-07-07')continue;const key=JSON.stringify([w,slot]);aggregated.set(key,(aggregated.get(key)||0)+1);valid++}return {valid,skipped}}
async function loadFolder(force=true){if(!selectedFolder||busy)return;busy=true;setButtons();setSync('Checking OneDrive CSV files…');try{
 const items=(await children(selectedFolder.id)).filter(f=>f.file&&/\.csv$/i.test(f.name)).sort((a,b)=>a.name.localeCompare(b.name));if(!items.length)throw Error('No CSV files found directly inside this folder.');
 const fingerprint=JSON.stringify(items.map(f=>[f.id,f.lastModifiedDateTime,f.size]));if(!force&&fingerprint===lastFingerprint){setSync(`No changes detected · ${items.length} CSV files · Checked ${new Date().toLocaleTimeString('en-GB')}`);return}
 const aggregate=new Map();let valid=0,skipped=0;for(let i=0;i<items.length;i++){const f=items[i];setSync(`Loading ${i+1} of ${items.length}: ${f.name}`);const detail=await graph(`/me/drive/items/${encodeURIComponent(f.id)}`);const url=detail['@microsoft.graph.downloadUrl'];if(!url)throw Error('OneDrive did not provide a download URL for '+f.name);const response=await fetch(url,{cache:'no-store'});if(!response.ok)throw Error(`Could not download ${f.name}: HTTP ${response.status}`);const stats=aggregateCSV(await response.text(),aggregate);valid+=stats.valid;skipped+=stats.skipped}
 records=[...aggregate].map(([key,count])=>{const [bookedWeek,slot]=JSON.parse(key);return {week:bookedWeek,bookedWeek,slot,status:'Booked',count}});weeks=[...new Set(records.map(r=>r.bookedWeek))].sort();types=[...new Set(records.map(r=>r.slot))].sort();statuses=['Booked'];
 // Assign obvious names on first encounter; retain any existing manual choices.
 // Do not guess assignments: unknown slot types stay Uncategorised until manually assigned.

 saveCategories(false);lastFingerprint=fingerprint;rebuildCategoryFilter();renderEditor();refresh();setSync(`${items.length} CSV files · ${valid.toLocaleString('en-GB')} booked rows counted · ${skipped.toLocaleString('en-GB')} skipped · Updated ${new Date().toLocaleString('en-GB')}`)
 }catch(e){setSync('Refresh failed: '+e.message)}finally{busy=false;setButtons()}}
$('signIn').onclick=async()=>{try{if(!auth){await initAuth();if(!auth)throw Error('Microsoft sign-in is still unavailable. Refresh the page (Ctrl+F5).')}await auth.loginRedirect({scopes:SCOPES,prompt:'select_account'})}catch(e){setStatus('Sign-in failed: '+e.message)}};
$('signOut').onclick=async()=>{if(!activeAccount)return;try{await auth.logoutRedirect({account:activeAccount,postLogoutRedirectUri:REDIRECT_URI})}catch(e){setStatus('Sign-out failed: '+e.message)}};
$('chooseFolder').onclick=showFolderPicker;
$('sourceSetup').onclick=showFolderPicker;
$('cancelFolder').onclick=()=>$('folderPicker').classList.add('hidden');
$('pickerUp').onclick=async()=>{if(!folderView.trail.length)return;const parent=folderView.trail.at(-1);folderView={...parent,trail:folderView.trail.slice(0,-1)};await showFolderContents()};
$('selectFolder').onclick=async()=>{selectedFolder={id:folderView.id,name:[...folderView.trail.map(x=>x.name),folderView.name].join(' / ')};localStorage.setItem(FOLDER_KEY,JSON.stringify(selectedFolder));$('folderStatus').textContent='Selected folder: '+selectedFolder.name;$('folderPicker').classList.add('hidden');$('onedrivePanel').classList.add('hidden');$('sourceSetup').classList.add('hidden');cloudFolder=selectedFolder;await saveCloudCategories();lastFingerprint='';await loadFolder(true)};
$('refreshData').onclick=()=>loadFolder(true);
setInterval(()=>{if(activeAccount&&selectedFolder&&!busy)loadFolder(false)},5*60*1000);
const holidayDates=['2025-01-01','2025-04-18','2025-04-21','2025-05-05','2025-05-26','2025-08-25','2025-12-25','2025-12-26','2026-01-01','2026-04-03','2026-04-06','2026-05-04','2026-05-25','2026-08-31','2026-12-25','2026-12-28'];
function mondayOf(dateString){const d=new Date(dateString+'T12:00:00Z');const day=d.getUTCDay();d.setUTCDate(d.getUTCDate()-((day+6)%7));return d.toISOString().slice(0,10)}
const bankHolidayWeeks=new Set(holidayDates.map(mondayOf));let excludeBankHolidays=false;
document.getElementById('excludeBH').onchange=e=>{excludeBankHolidays=!e.target.checked;refresh()};
function isoWeek(weekStart){const d=new Date(weekStart+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+3);const year=d.getUTCFullYear();const jan4=new Date(Date.UTC(year,0,4,12));jan4.setUTCDate(jan4.getUTCDate()-((jan4.getUTCDay()+6)%7));const week=1+Math.round((d-jan4)/604800000);return {year,week}}
const colours=['#1475b8','#e57a28','#159a78','#9053b4','#c64c69','#65748b'];
let current=[];let continuousMode=false;let showCurrentWeek=true;const thisWeek=mondayOf((()=>{const parts=new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());const v=Object.fromEntries(parts.map(p=>[p.type,p.value]));return `${v.year}-${v.month}-${v.day}`})());document.getElementById('currentWeekToggle').onchange=e=>{showCurrentWeek=e.target.checked;refresh()};document.getElementById('displayMode').onchange=e=>{continuousMode=e.target.checked;document.getElementById('timelineValue').textContent=continuousMode?'Continuous':'Superimposed';refresh()};function refresh(){const basis='bookedWeek';const activeWeeks=[...new Set(records.map(r=>r[basis]))].sort();const selectedWeeks=activeWeeks.filter(w=>w>='2025-07-07'&&(showCurrentWeek||w!==thisWeek)&&(!excludeBankHolidays||!bankHolidayWeeks.has(w)));const sums=new Map(selectedWeeks.map(w=>[w,0]));for(const r of records){if(!sums.has(r[basis]))continue;if(r.status!=='Booked')continue;sums.set(r[basis],sums.get(r[basis])+r.count)}current=[...sums].map(([week,count])=>({week,count,...isoWeek(week)}));const total=current.reduce((a,r)=>a+r.count,0);
const yearColours={urgent:['#8b111e','#d43846','#f08088','#6f0712'],routine:['#075d34','#199c5a','#79c99c','#a1d9b4']};
function monthPosition(date){const d=new Date(date+'T12:00:00Z');const month=d.getUTCMonth();const days=new Date(Date.UTC(d.getUTCFullYear(),month+1,0)).getUTCDate();return month+(d.getUTCDate()-1)/days}
const axis={type:'linear',range:[0,12],title:'Month of year',tickmode:'array',tickvals:Array.from({length:12},(_,i)=>i+0.5),ticktext:['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],gridcolor:'#edf1f5'};
function drawCategoryChart(target,category,palette,sharedMax){
 const enabled=true;
 const totals=new Map(selectedWeeks.map(w=>[w,0]));
 if(enabled)for(const r of records){
  if(!totals.has(r[basis]))continue;
  if(r.status!=='Booked')continue;
  if((assignments[r.slot]||'Uncategorised').trim().toLowerCase()!==category)continue;
  totals.set(r[basis],totals.get(r[basis])+r.count);
 }
 const points=[...totals].map(([date,count])=>({date,count,year:Number(date.slice(0,4)),week:isoWeek(date).week}));
 const years=[...new Set(points.map(p=>p.year))].sort();
 const latestYear=years.length?Math.max(...years):2026;
 const values=points.map(p=>p.count);
 const average=values.length?values.reduce((sum,v)=>sum+v,0)/values.length:0;
 const traces=years.map((year,i)=>{
  const data=points.filter(p=>p.year===year).sort((a,b)=>a.date.localeCompare(b.date));
  // Keep adjacent retained weeks connected, including across excluded bank-holiday weeks.
  const plotted=data;
  return {name:String(year),x:plotted.map(p=>continuousMode?p.date:monthPosition(p.date)),y:plotted.map(p=>p.count),customdata:plotted.map(p=>p.date),type:'scatter',mode:'lines+markers',connectgaps:true,opacity:year===latestYear?1:0.34,line:{color:palette[i%palette.length],width:year===latestYear?4.5:2,shape:'spline',smoothing:0.65},marker:{size:year===latestYear?5:3},hovertemplate:'%{customdata}<br><b>%{y:,} appointments</b><extra>'+year+'</extra>'};
 });
 // One flat average, calculated from every displayed week across every year.
 traces.push({name:'All-years weekly average ('+Math.round(average).toLocaleString('en-GB')+')',x:continuousMode?[selectedWeeks[0]||'2025-07-07',selectedWeeks[selectedWeeks.length-1]||'2025-07-14']:[0,12],y:[average,average],type:'scatter',mode:'lines',hoverinfo:'skip',opacity:0.58,line:{color:category==='urgent'?'#9c2633':'#23885b',width:2,dash:'dash'},showlegend:true});
 Plotly.react(target,traces,{margin:{l:58,r:12,t:32,b:65},paper_bgcolor:'white',plot_bgcolor:'white',legend:{orientation:'h',y:1.17,x:0},xaxis:continuousMode?{type:'date',title:'Booking date',tickformat:'%b %Y',gridcolor:'#edf1f5'}:axis,yaxis:{title:'Appointments',range:[0,sharedMax],dtick:sharedMax/5,gridcolor:'#edf1f5'},hovermode:'closest',annotations:enabled?[]:[{text:'Category excluded by filter',xref:'paper',yref:'paper',x:.5,y:.5,showarrow:false}]},{responsive:true,displaylogo:false});
}
let maximum=0;

for(const category of ['urgent','routine']){

 const totals=new Map(selectedWeeks.map(w=>[w,0]));
 for(const r of records){
  if(!totals.has(r[basis]))continue;
  if(r.status!=='Booked')continue;
  if((assignments[r.slot]||'Uncategorised').trim().toLowerCase()!==category)continue;
  totals.set(r[basis],totals.get(r[basis])+r.count);
 }
 for(const count of totals.values())maximum=Math.max(maximum,count);
}
const rawStep=Math.max(1,maximum/5),magnitude=Math.pow(10,Math.floor(Math.log10(rawStep)));
const tickStep=([1,2,5,10].find(v=>v*magnitude>=rawStep)||10)*magnitude;
const sharedMax=5*tickStep;
drawCategoryChart('urgentChart','urgent',yearColours.urgent,sharedMax);
drawCategoryChart('routineChart','routine',yearColours.routine,sharedMax);
}
const STORAGE_KEY='appointment-dashboard-categories-v1';let categoryNames=['Urgent','Routine'],assignments={};try{const saved=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');categoryNames=Array.isArray(saved.categories)&&saved.categories.length?saved.categories.filter(x=>typeof x==='string'&&x.trim()&&x!=='Uncategorised'):['Urgent','Routine'];assignments=saved.assignments&&typeof saved.assignments==='object'?saved.assignments:{};}catch(e){}function guessCategory(t){return /RED|YELLOW|AMBER|URGENT|SAME DAY|ACUTE/i.test(t)?'Urgent':/GREEN|FOLLOW UP|ROUTINE/i.test(t)?'Routine':null}function saveCategories(sync=true){try{localStorage.setItem(STORAGE_KEY,JSON.stringify({categories:categoryNames,assignments}));}catch(e){alert('Could not save settings in this browser. Export a backup instead.')}if(sync)scheduleCloudSave()}function allCategories(){return [...categoryNames,'Uncategorised']}function rebuildCategoryFilter(){const old=new Set([...slots.selectedOptions].map(o=>o.value));slots.replaceChildren();for(const c of allCategories()){const o=document.createElement('option');o.value=c;o.textContent=c;o.selected=old.has(c);slots.append(o)}refresh()}function renderEditor(){const names=document.getElementById('categoryNames');names.replaceChildren();for(const c of categoryNames){const div=document.createElement('div');div.className='catrow';const label=document.createElement('span');label.textContent=c+' ('+types.filter(t=>assignments[t]===c).length+')';const btn=document.createElement('button');btn.textContent='Delete';btn.style.width='auto';btn.onclick=()=>{if(!confirm('Delete category “'+c+'”? Its slot types will become Uncategorised.'))return;categoryNames=categoryNames.filter(x=>x!==c);for(const t of types)if(assignments[t]===c)assignments[t]='Uncategorised';saveCategories();rebuildCategoryFilter();renderEditor()};div.append(label,btn);names.append(div)}const list=document.getElementById('slotAssignments');list.replaceChildren();const q=document.getElementById('slotSearch').value.toLowerCase();for(const t of types.filter(t=>t.toLowerCase().includes(q))){const row=document.createElement('div');row.className='catrow';const name=document.createElement('span');name.textContent=t;const sel=document.createElement('select');for(const c of allCategories()){const opt=document.createElement('option');opt.value=c;opt.textContent=c;sel.append(opt)}sel.value=categoryNames.includes(assignments[t])?assignments[t]:'Uncategorised';sel.onchange=()=>{assignments[t]=sel.value;saveCategories();rebuildCategoryFilter();renderEditor()};row.append(name,sel);list.append(row)}document.getElementById('assignmentSummary').textContent=types.filter(t=>categoryNames.includes(assignments[t])).length+' of '+types.length+' slot types assigned to named categories.'}document.getElementById('addCategory').onclick=()=>{const el=document.getElementById('newCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||categoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This category name already exists or is reserved.');return}categoryNames.push(c);el.value='';saveCategories();rebuildCategoryFilter();renderEditor()};document.getElementById('newCategory').addEventListener('keydown',e=>{if(e.key==='Enter')document.getElementById('addCategory').click()});document.getElementById('slotSearch').addEventListener('input',renderEditor);function showPage(which){document.getElementById('chartPage').classList.toggle('hidden',which!=='chart');document.getElementById('categoryPage').classList.toggle('hidden',which!=='categories');document.getElementById('navChart').classList.toggle('active',which==='chart');document.getElementById('navCategories').classList.toggle('active',which==='categories');if(which==='chart')setTimeout(()=>{refresh();Plotly.Plots.resize('urgentChart');Plotly.Plots.resize('routineChart')},0);else renderEditor()}document.getElementById('navChart').onclick=()=>showPage('chart');document.getElementById('navCategories').onclick=()=>showPage('categories');document.getElementById('exportCategories').onclick=()=>{const blob=new Blob([JSON.stringify({categories:categoryNames,assignments},null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='appointment_slot_categories.json';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)};document.getElementById('importCategories').onclick=()=>document.getElementById('importFile').click();document.getElementById('importFile').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const data=JSON.parse(await f.text());if(!Array.isArray(data.categories)||!data.assignments||typeof data.assignments!=='object'||Array.isArray(data.assignments))throw Error('Invalid format');const names=data.categories.filter(c=>typeof c==='string'&&c.trim()&&c!=='Uncategorised');categoryNames=[...new Set(names)];assignments={};for(const t of types)if(categoryNames.includes(data.assignments[t])||data.assignments[t]==='Uncategorised')assignments[t]=data.assignments[t];saveCategories();rebuildCategoryFilter();renderEditor();alert('Category settings imported.')}catch(err){alert('Unable to import settings: '+err.message)}e.target.value=''};rebuildCategoryFilter();renderEditor();
refresh();

initAuth();
