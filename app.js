"use strict";
const CLIENT_ID='ba2bc895-4213-447c-ba62-d7edec6e8173';
const SCOPES=['Files.Read','Files.ReadWrite.AppFolder'];
const REDIRECT_URI='https://shammynoor.github.io/weekly-demand-dashboard/index.html';
const GRAPH='https://graph.microsoft.com/v1.0';
const FOLDER_KEY='weekly-demand-onedrive-folder-v1';
let records=[],weeks=[],types=[],statuses=[];
const slots={selectedOptions:[],replaceChildren(){},append(){},options:[]};
const $=id=>document.getElementById(id);
let auth,authInitPromise=null,activeAccount=null,authenticatedOnce=false,selectedFolder=null,folderView={id:'root',name:'OneDrive',trail:[]},busy=false,lastFingerprint='';
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
   clinicianCategoryNames=Array.isArray(data.clinicianCategories)?validCategoryNames(data.clinicianCategories):clinicianCategoryNames;
   clinicianAssignments=data.clinicianAssignments?validAssignments(data.clinicianAssignments,clinicianCategoryNames):clinicianAssignments;
   consultCategoryNames=Array.isArray(data.consultCategories)?validCategoryNames(data.consultCategories):[];
   consultAssignments=data.consultAssignments?validAssignments(data.consultAssignments,consultCategoryNames):{};
   // Older dashboard versions only saved Urgent and Routine to OneDrive.
   // On the original computer, preserve additional local categories during migration.
   const legacy=Number(data.version||1)<2;
   const localNames=categoryNames.slice();
   const localAssignments={...assignments};
   categoryNames=names.length?names:['Urgent','Routine'];
   ensureFollowUpCategory();
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
  }else cloudNotice('No cloud settings yet. Click Save to publish your category assignments.');
  cloudReady=true;
  saveCategories(false);
  renderEditor();
  renderClinicianEditor();
  renderConsultEditor();
  refresh();
 }catch(e){
  cloudNotice('Could not load cloud categories: '+e.message+' (local settings retained)');
  cloudReady=false;
 }
}
let categoryDirty=false;
function markCategoryDirty(){
 categoryDirty=true;
 cloudNotice('Unsaved category changes — click Save.');
}
function scheduleCloudSave(){markCategoryDirty()}
async function saveCloudCategories(options={}){
 const manual=!!options.manual;
 if(!activeAccount||!cloudReady){
  if(manual)cloudNotice('Cannot save: sign in and wait for OneDrive settings to load.');
  return false;
 }
 if(cloudBusy){
  if(manual)cloudNotice('A OneDrive save is already running. Please try again.');
  return false;
 }
 cloudBusy=true;
 const button=$('saveToOneDrive');
 if(button)button.disabled=true;
 if(manual)cloudNotice('Saving all categories and slot assignments to OneDrive…');
 const snapshot=JSON.stringify({
  version:5,
  categories:categoryNames.slice(),
  assignments:{...assignments},
  clinicianCategories:clinicianCategoryNames.slice(),
  clinicianAssignments:{...clinicianAssignments},
  consultCategories:consultCategoryNames.slice(),
  consultAssignments:{...consultAssignments},
  folder:selectedFolder||cloudFolder||null,
  updatedAt:new Date().toISOString()
 });
 try{
  const t=await token();
  const response=await fetch(GRAPH+'/me/drive/special/approot:/'+CLOUD_SETTINGS_NAME+':/content',{
   method:'PUT',
   headers:{Authorization:'Bearer '+t,'Content-Type':'application/json'},
   body:snapshot
  });
  if(!response.ok)throw Error('HTTP '+response.status);
  if(manual){
   // Only clear the dirty flag if the user has not changed anything during the save.
   const current=JSON.stringify({categories:categoryNames,assignments,clinicianCategories:clinicianCategoryNames,clinicianAssignments,consultCategories:consultCategoryNames,consultAssignments});
   const saved=JSON.parse(snapshot);
   if(current===JSON.stringify({categories:saved.categories,assignments:saved.assignments,clinicianCategories:saved.clinicianCategories,clinicianAssignments:saved.clinicianAssignments,consultCategories:saved.consultCategories,consultAssignments:saved.consultAssignments})){
    categoryDirty=false;
    cloudNotice('Saved to OneDrive successfully. Other computers will see these categories after refreshing.');
   }else{
    categoryDirty=true;
    cloudNotice('Saved, but further edits are unsaved. Click Save again.');
   }
  }
  return true;
 }catch(e){
  cloudNotice('Save failed: '+e.message+'. Changes are still stored in this browser.');
  return false;
 }finally{
  cloudBusy=false;
  if(button)button.disabled=false;
 }
}
function setStatus(s){$('connectionStatus').textContent=s;$('loginStatus').textContent=s}
function setSync(s){$('syncStatus').textContent=s}
function setButtons(){
 // Never interpret a background chart/data refresh as a sign-out.
 // Recover the active MSAL account from its persistent cache if needed.
 if(!activeAccount&&auth){
  try{activeAccount=auth.getActiveAccount()||auth.getAllAccounts()[0]||null}catch(e){}
 }
 if(activeAccount)authenticatedOnce=true;
 const logged=!!activeAccount||authenticatedOnce;
 $('loginScreen').classList.toggle('hidden',logged);
 $('dashboardShell').classList.toggle('hidden',!logged);
 $('signIn').disabled=logged;
 $('signOut').disabled=!logged;
 $('chooseFolder').disabled=!logged;
 $('refreshData').disabled=!logged||!selectedFolder||busy;
}
async function initAuth(){
 if(authInitPromise)return authInitPromise;
 authInitPromise=(async()=>{
 try{
  if(!window.msal)throw Error('Microsoft sign-in library did not load. Please refresh with Ctrl+F5 or check browser content blocking.');
  if(!window.Plotly||!window.Papa)setSync('A chart or CSV library did not load. Refresh the page to retry.');
  auth=new msal.PublicClientApplication({auth:{clientId:CLIENT_ID,authority:'https://login.microsoftonline.com/consumers',redirectUri:REDIRECT_URI,navigateToLoginRequestUrl:false},cache:{cacheLocation:'localStorage',storeAuthStateInCookie:false}});
  if(typeof auth.initialize==='function')await auth.initialize();const result=await auth.handleRedirectPromise();
  activeAccount=result?.account||auth.getActiveAccount()||auth.getAllAccounts()[0]||activeAccount||null;
  if(activeAccount){
   authenticatedOnce=true;
   auth.setActiveAccount(activeAccount);
   setButtons(); // Keep the dashboard visible while OneDrive requests are running.
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
 }catch(e){setStatus('Sign-in setup error: '+e.message);if(auth){try{activeAccount=activeAccount||auth.getActiveAccount()||auth.getAllAccounts()[0]||null}catch(_){}}}finally{setButtons()}
 })();
 return authInitPromise;
}
async function token(){
 if(!activeAccount)throw Error('Please sign in first.');
 try{return (await auth.acquireTokenSilent({scopes:SCOPES,account:activeAccount})).accessToken}
 catch(e){
  // Never automatically redirect from the dashboard on a background refresh.
  // A transient consent/token problem should not look like an unexpected logout.
  if(e instanceof msal.InteractionRequiredAuthError || /interaction_required|consent_required|login_required/i.test(String(e.errorCode||''))){
   throw Error('Microsoft needs to renew access. Your dashboard is still signed in; please refresh the page or sign out and sign in again if the problem continues.');
  }
  throw e;
 }
}
async function graph(url){const bearer=await token();const response=await fetch(url.startsWith('https://')?url:GRAPH+url,{headers:{Authorization:'Bearer '+bearer}});if(!response.ok){let message='HTTP '+response.status;try{const body=await response.json();message=body.error?.message||message}catch{}throw Error(message)}return response.json()}
async function children(id){let url=id==='root'?'/me/drive/root/children?$top=200&$select=id,name,folder,file,size,lastModifiedDateTime':`/me/drive/items/${encodeURIComponent(id)}/children?$top=200&$select=id,name,folder,file,size,lastModifiedDateTime`;const items=[];while(url){const r=await graph(url);items.push(...(r.value||[]));url=r['@odata.nextLink']||null}return items}
async function showFolderPicker(){$('onedrivePanel').classList.remove('hidden');folderView={id:'root',name:'OneDrive',trail:[]};$('folderPicker').classList.remove('hidden');await showFolderContents()}
async function showFolderContents(){try{$('pickerPath').textContent=[...folderView.trail.map(x=>x.name),folderView.name].join(' / ');$('pickerEntries').textContent='Loading folders…';const folders=(await children(folderView.id)).filter(x=>x.folder).sort((a,b)=>a.name.localeCompare(b.name));$('pickerEntries').replaceChildren();if(!folders.length)$('pickerEntries').textContent='No subfolders. You can select this folder.';for(const f of folders){const btn=document.createElement('button');btn.className='folder-entry';btn.textContent='📁 '+f.name;btn.onclick=async()=>{folderView={id:f.id,name:f.name,trail:[...folderView.trail,{id:folderView.id,name:folderView.name}]};await showFolderContents()};$('pickerEntries').append(btn)}}catch(e){$('pickerEntries').textContent='Could not list folders: '+e.message}}
function weekMonday(d){const x=new Date(d);x.setUTCHours(12,0,0,0);x.setUTCDate(x.getUTCDate()-((x.getUTCDay()+6)%7));return x.toISOString().slice(0,10)}
function parseDate(s){const m=/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec((s||'').trim());if(!m)return null;const mi=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].toLowerCase());if(mi<0)return null;const d=new Date(Date.UTC(Number(m[3]),mi,Number(m[1])));return d.getUTCMonth()===mi&&d.getUTCDate()===Number(m[1])?d:null}
function aggregateCSV(csv,aggregated){const parsed=Papa.parse(csv,{header:true,skipEmptyLines:'greedy'});const fields=parsed.meta.fields||[];for(const required of ['Appointment Date','Appointment Time','Booked Time to Slot Time','Slot Type','Availability'])if(!fields.includes(required))throw Error('Missing column: '+required);let day='',time='',valid=0,skipped=0;for(const r of parsed.data){if(r['Appointment Date']?.trim())day=r['Appointment Date'].trim();if(r['Appointment Time']?.trim())time=r['Appointment Time'].trim();if((r['Availability']||'').trim().toLowerCase()!=='booked')continue;const lead=Number((r['Booked Time to Slot Time']||'').replace(/,/g,''));const dt=parseDate(day);const tm=/^(\d{1,2}):(\d{2})/.exec(time);if(!dt||!tm||!Number.isFinite(lead)||!(r['Booked Time to Slot Time']||'').trim()){skipped++;continue}dt.setUTCHours(Number(tm[1]),Number(tm[2]),0,0);const booked=new Date(dt.getTime()-lead*60000);if(!Number.isFinite(booked.getTime())){skipped++;continue}const w=weekMonday(booked),slot=(r['Slot Type']||'').trim()||'Unknown';if(w<'2025-07-07')continue;const key=JSON.stringify([w,slot]);aggregated.set(key,(aggregated.get(key)||0)+1);valid++}return {valid,skipped}}
// Drag-and-drop category editor. Existing OneDrive JSON schema remains unchanged.
let pickedSlot=null;
function renderDropEditor(dimension){
 const cfg={
  urgency:{names:categoryNames,assign:assignments,boxes:'categoryNames',pool:'slotAssignments',search:'slotSearch',summary:'assignmentSummary',render:renderEditor},
  clinician:{names:clinicianCategoryNames,assign:clinicianAssignments,boxes:'clinicianCategoryNames',pool:'clinicianSlotAssignments',search:'clinicianSlotSearch',summary:'clinicianAssignmentSummary',render:renderClinicianEditor},
  consult:{names:consultCategoryNames,assign:consultAssignments,boxes:'consultCategoryNames',pool:'consultSlotAssignments',search:'consultSlotSearch',summary:'consultAssignmentSummary',render:renderConsultEditor}
 }[dimension];
 if(!cfg)return;
 const boxes=$(cfg.boxes),pool=$(cfg.pool);if(!boxes||!pool)return;
 boxes.replaceChildren();pool.replaceChildren();
 const query=$(cfg.search).value.trim().toLowerCase();
 const filtered=types.filter(t=>t.toLowerCase().includes(query));
 const current=t=>cfg.names.includes(cfg.assign[t])?cfg.assign[t]:'Uncategorised';
 const place=(slot,category)=>{
  if(!types.includes(slot)||!(category==='Uncategorised'||cfg.names.includes(category)))return;
  cfg.assign[slot]=category;
  pickedSlot=null;
  saveCategories();
  if(dimension==='urgency')rebuildCategoryFilter();
  cfg.render();
 };
 function makeCard(t){
  const card=document.createElement('div');card.className='slot-chip';card.draggable=true;card.tabIndex=0;
  card.setAttribute('role','button');card.setAttribute('aria-label',t+'. Drag to a category, or press Enter to select.');
  const ico=document.createElement('span');ico.className='slot-chip-icon';ico.textContent='⠿';ico.setAttribute('aria-hidden','true');
  const label=document.createElement('span');label.textContent=t;card.append(ico,label);
  card.addEventListener('dragstart',e=>{pickedSlot={dimension,slot:t};e.dataTransfer.setData('text/plain',t);e.dataTransfer.effectAllowed='move';card.classList.add('dragging')});
  card.addEventListener('dragend',()=>{card.classList.remove('dragging');document.querySelectorAll('.drop-over').forEach(el=>el.classList.remove('drop-over'))});
  card.addEventListener('click',()=>{pickedSlot={dimension,slot:t};document.querySelectorAll('.slot-chip.selected').forEach(el=>el.classList.remove('selected'));card.classList.add('selected')});
  card.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();card.click()}});
  return card;
 }
 function makeZone(name,isPool=false){
  const zone=document.createElement('div');zone.className='drop-zone'+(isPool?' unassigned-zone':'');zone.tabIndex=0;
  zone.setAttribute('role','group');zone.setAttribute('aria-label','Drop slot types into '+name);
  const head=document.createElement('div');head.className='drop-zone-header';
  const title=document.createElement('strong');title.textContent=name;
  head.append(title);
  if(!isPool){
   const rename=document.createElement('button');rename.type='button';rename.className='drop-rename';rename.textContent='Rename';
   rename.setAttribute('aria-label','Rename '+name);
   rename.onclick=e=>{
    e.stopPropagation();
    const proposed=prompt('Rename category “'+name+'”:',name);
    if(proposed===null)return;
    const next=proposed.trim();
    if(!next||next.toLowerCase()==='uncategorised'||cfg.names.some(c=>c!==name&&c.toLowerCase()===next.toLowerCase())){
     alert('Enter a unique category name (Uncategorised is reserved).');return;
    }
    if(next===name)return;
    // Update both the category and all existing assignments atomically.
    const index=cfg.names.indexOf(name);if(index<0)return;
    cfg.names[index]=next;
    for(const slot of Object.keys(cfg.assign))if(cfg.assign[slot]===name)cfg.assign[slot]=next;
    // Preserve the selected urgency filter when its category name changes.
    if(dimension==='urgency')for(const option of slots.options)if(option.value===name)option.value=next;
    saveCategories();if(dimension==='urgency')rebuildCategoryFilter();cfg.render();
   };
   head.append(rename);
   const del=document.createElement('button');del.type='button';del.className='drop-delete';del.textContent='Delete';
   del.setAttribute('aria-label','Delete '+name);
   del.onclick=e=>{e.stopPropagation();if(!confirm('Delete category “'+name+'”? Its slot types will become Uncategorised.'))return;
    if(dimension==='urgency'){categoryNames=categoryNames.filter(c=>c!==name);for(const t of types)if(assignments[t]===name)assignments[t]='Uncategorised'}
    else if(dimension==='clinician'){clinicianCategoryNames=clinicianCategoryNames.filter(c=>c!==name);for(const t of types)if(clinicianAssignments[t]===name)clinicianAssignments[t]='Uncategorised'}
    else{consultCategoryNames=consultCategoryNames.filter(c=>c!==name);for(const t of types)if(consultAssignments[t]===name)consultAssignments[t]='Uncategorised'}
    saveCategories();if(dimension==='urgency')rebuildCategoryFilter();cfg.render();
   };head.append(del);
  }
  zone.append(head);
  const body=document.createElement('div');body.className='drop-zone-body';
  const matching=filtered.filter(t=>current(t)===name);
  if(!matching.length){const hint=document.createElement('p');hint.className='drop-hint';hint.textContent=query?'No matching slot types':'Drop slot types here';body.append(hint)}
  for(const t of matching)body.append(makeCard(t));
  zone.append(body);
  zone.addEventListener('dragover',e=>{if(pickedSlot?.dimension!==dimension)return;e.preventDefault();e.dataTransfer.dropEffect='move';zone.classList.add('drop-over')});
  zone.addEventListener('dragleave',e=>{if(!zone.contains(e.relatedTarget))zone.classList.remove('drop-over')});
  zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('drop-over');if(pickedSlot?.dimension===dimension)place(pickedSlot.slot,name)});
  // Accessible alternative to drag-and-drop: select a slot, then click its destination.
  zone.addEventListener('click',e=>{if(e.target.closest('button')||e.target.closest('.slot-chip'))return;if(pickedSlot?.dimension===dimension)place(pickedSlot.slot,name)});
  zone.addEventListener('keydown',e=>{if((e.key==='Enter'||e.key===' ')&&pickedSlot?.dimension===dimension){e.preventDefault();place(pickedSlot.slot,name)}});
  return zone;
 }
 pool.append(makeZone('Uncategorised',true));
 for(const c of cfg.names)boxes.append(makeZone(c));
 $(cfg.summary).textContent=types.filter(t=>cfg.names.includes(cfg.assign[t])).length+' of '+types.length+' slot types assigned. Drag a slot icon into a category box (or select it, then click a box).';
}
async function loadFolder(force=true){if(!selectedFolder||busy)return;busy=true;setButtons();setSync('Checking OneDrive CSV files…');try{
 const items=(await children(selectedFolder.id)).filter(f=>f.file&&/\.csv$/i.test(f.name)).sort((a,b)=>a.name.localeCompare(b.name));if(!items.length)throw Error('No CSV files found directly inside this folder.');
 const fingerprint=JSON.stringify(items.map(f=>[f.id,f.lastModifiedDateTime,f.size]));if(!force&&fingerprint===lastFingerprint){setSync(`No changes detected · ${items.length} CSV files · Checked ${new Date().toLocaleTimeString('en-GB')}`);return}
 const aggregate=new Map();let valid=0,skipped=0;for(let i=0;i<items.length;i++){const f=items[i];setSync(`Loading ${i+1} of ${items.length}: ${f.name}`);const detail=await graph(`/me/drive/items/${encodeURIComponent(f.id)}`);const url=detail['@microsoft.graph.downloadUrl'];if(!url)throw Error('OneDrive did not provide a download URL for '+f.name);const response=await fetch(url,{cache:'no-store'});if(!response.ok)throw Error(`Could not download ${f.name}: HTTP ${response.status}`);const stats=aggregateCSV(await response.text(),aggregate);valid+=stats.valid;skipped+=stats.skipped}
 records=[...aggregate].map(([key,count])=>{const [bookedWeek,slot]=JSON.parse(key);return {week:bookedWeek,bookedWeek,slot,status:'Booked',count}});weeks=[...new Set(records.map(r=>r.bookedWeek))].sort();types=[...new Set(records.map(r=>r.slot))].sort();statuses=['Booked'];
 // Assign obvious names on first encounter; retain any existing manual choices.
 // Do not guess assignments: unknown slot types stay Uncategorised until manually assigned.

 saveCategories(false);lastFingerprint=fingerprint;
rebuildCategoryFilter();renderEditor();renderClinicianEditor();renderConsultEditor();refresh();setSync(`${items.length} CSV files · ${valid.toLocaleString('en-GB')} booked rows counted · ${skipped.toLocaleString('en-GB')} skipped · Updated ${new Date().toLocaleString('en-GB')}`)
 }catch(e){setSync('Refresh failed: '+e.message)}finally{busy=false;setButtons()}}
$('signIn').onclick=async()=>{try{if(!auth){await initAuth();if(!auth)throw Error('Microsoft sign-in is still unavailable. Refresh the page (Ctrl+F5).')}await auth.loginRedirect({scopes:SCOPES})}catch(e){setStatus('Sign-in failed: '+e.message)}};
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
const bankHolidayWeeks=new Set(holidayDates.map(mondayOf));let excludeBankHolidays=true;
document.getElementById('excludeBH').onchange=e=>{excludeBankHolidays=!e.target.checked;refresh()};
function isoWeek(weekStart){const d=new Date(weekStart+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+3);const year=d.getUTCFullYear();const jan4=new Date(Date.UTC(year,0,4,12));jan4.setUTCDate(jan4.getUTCDate()-((jan4.getUTCDay()+6)%7));const week=1+Math.round((d-jan4)/604800000);return {year,week}}
const colours=['#1475b8','#e57a28','#159a78','#9053b4','#c64c69','#65748b'];
let current=[];let continuousMode=false;let showCurrentWeek=true;const thisWeek=mondayOf((()=>{const parts=new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());const v=Object.fromEntries(parts.map(p=>[p.type,p.value]));return `${v.year}-${v.month}-${v.day}`})());document.getElementById('currentWeekToggle').onchange=e=>{showCurrentWeek=e.target.checked;refresh()};document.getElementById('displayMode').onchange=e=>{continuousMode=e.target.checked;document.getElementById('timelineValue').textContent=continuousMode?'Continuous':'Superimposed';refresh()};function refresh(){const basis='bookedWeek';const activeWeeks=[...new Set(records.map(r=>r[basis]))].sort();const selectedWeeks=activeWeeks.filter(w=>w>='2025-07-07'&&(showCurrentWeek||w!==thisWeek)&&(!excludeBankHolidays||!bankHolidayWeeks.has(w)));const sums=new Map(selectedWeeks.map(w=>[w,0]));for(const r of records){if(!sums.has(r[basis]))continue;if(r.status!=='Booked')continue;sums.set(r[basis],sums.get(r[basis])+r.count)}current=[...sums].map(([week,count])=>({week,count,...isoWeek(week)}));const total=current.reduce((a,r)=>a+r.count,0);
const yearColours={urgent:['#8b111e','#d43846','#f08088','#6f0712'],routine:['#075d34','#199c5a','#79c99c','#a1d9b4'],followup:['#124b9b','#2675d2','#79acee','#a9cdf6']};
function monthPosition(date){const d=new Date(date+'T12:00:00Z');const month=d.getUTCMonth();const days=new Date(Date.UTC(d.getUTCFullYear(),month+1,0)).getUTCDate();return month+(d.getUTCDate()-1)/days}
const axis={type:'linear',range:[0,12],title:'Month of year',tickmode:'array',tickvals:Array.from({length:12},(_,i)=>i+0.5),ticktext:['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],gridcolor:'#edf1f5'};
function drawCategoryChart(target,category,palette,sharedMax){
 const enabled=true;
 const totals=new Map(selectedWeeks.map(w=>[w,0]));
 if(enabled)for(const r of records){
  if(!totals.has(r[basis]))continue;
  if(r.status!=='Booked')continue;
  if((assignments[r.slot]||'Uncategorised').trim().toLowerCase().replace(/[\s-]+/g,'')!==category)continue;
  totals.set(r[basis],totals.get(r[basis])+r.count);
 }
 const points=[...totals].map(([date,count])=>({date,count,year:Number(date.slice(0,4)),week:isoWeek(date).week}));
 const years=[...new Set(points.map(p=>p.year))].sort();
 const latestYear=years.length?Math.max(...years):2026;
 const values=points.map(p=>p.count);
 const hasAppointments=values.some(v=>v>0);
 const average=values.length?values.reduce((sum,v)=>sum+v,0)/values.length:0;
 const traces=[];
 for(const [i,year] of years.entries()){
  const data=points.filter(p=>p.year===year).sort((a,b)=>a.date.localeCompare(b.date));
  const last=data[data.length-1];
  const partial=showCurrentWeek&&last?.date===thisWeek;
  // Never draw a solid segment into an incomplete current week.
  const complete=partial?data.slice(0,-1):data;
  const colour=palette[i%palette.length];
  const width=year===latestYear?4.5:2;
  const opacity=year===latestYear?1:0.34;
  const x=p=>continuousMode?p.date:monthPosition(p.date);
  traces.push({name:String(year),x:complete.map(x),y:complete.map(p=>p.count),customdata:complete.map(p=>p.date),type:'scatter',mode:'lines+markers',connectgaps:true,opacity,line:{color:colour,width,shape:'spline',smoothing:0.65},marker:{size:year===latestYear?5:3},hovertemplate:'%{customdata}<br><b>%{y:,} appointments</b><extra>'+year+'</extra>'});
  if(partial){
   const previous=complete[complete.length-1];
   const partialPoints=previous?[previous,last]:[last];
   traces.push({name:String(year)+' — partial week',x:partialPoints.map(x),y:partialPoints.map(p=>p.count),customdata:partialPoints.map(p=>p.date),type:'scatter',mode:'lines+markers',showlegend:false,opacity,line:{color:colour,width,dash:'dash'},marker:{size:partialPoints.map((_,idx)=>idx===partialPoints.length-1?10:0),symbol:partialPoints.map((_,idx)=>idx===partialPoints.length-1?'circle-open':'circle'),color:colour,line:{color:colour,width:2}},hovertemplate:'%{customdata}<br><b>%{y:,} appointments</b><br>Current week (incomplete)<extra>'+year+'</extra>'});
  }
 }
 // One flat average, calculated from every displayed week across every year.
 traces.push({name:'All-years weekly average ('+Math.round(average).toLocaleString('en-GB')+')',x:continuousMode?[selectedWeeks[0]||'2025-07-07',selectedWeeks[selectedWeeks.length-1]||'2025-07-14']:[0,12],y:[average,average],type:'scatter',mode:'lines',hoverinfo:'skip',opacity:0.58,line:{color:category==='urgent'?'#9c2633':category==='followup'?'#286cb8':'#23885b',width:2,dash:'dash'},showlegend:true});
 Plotly.react(target,traces,{margin:{l:58,r:12,t:32,b:65},paper_bgcolor:'white',plot_bgcolor:'white',legend:{orientation:'h',y:1.17,x:0},xaxis:continuousMode?{type:'date',title:'Booking date',tickformat:'%b %Y',gridcolor:'#edf1f5'}:axis,yaxis:{title:'Appointments',range:[0,sharedMax],dtick:sharedMax/5,gridcolor:'#edf1f5'},hovermode:'closest',annotations:!hasAppointments&&category==='followup'?[{text:'No follow-up appointments found. In Manage Categories, assign slot types to Follow up and click Save.',xref:'paper',yref:'paper',x:.5,y:.55,showarrow:false,align:'center',font:{size:13,color:'#286cb8'}}]:[]},{responsive:true,displaylogo:false});
}
let maximum=0;
let followupMaximum=0;

for(const category of ['urgent','routine','followup']){

 const totals=new Map(selectedWeeks.map(w=>[w,0]));
 for(const r of records){
  if(!totals.has(r[basis]))continue;
  if(r.status!=='Booked')continue;
  if((assignments[r.slot]||'Uncategorised').trim().toLowerCase().replace(/[\s-]+/g,'')!==category)continue;
  totals.set(r[basis],totals.get(r[basis])+r.count);
 }
 for(const count of totals.values()){maximum=Math.max(maximum,count);if(category==='followup')followupMaximum=Math.max(followupMaximum,count);}
}
const rawStep=Math.max(1,maximum/5),magnitude=Math.pow(10,Math.floor(Math.log10(rawStep)));
const tickStep=([1,2,5,10].find(v=>v*magnitude>=rawStep)||10)*magnitude;
const sharedMax=5*tickStep;
drawCategoryChart('urgentChart','urgent',yearColours.urgent,sharedMax);
drawCategoryChart('routineChart','routine',yearColours.routine,sharedMax);
const followupStepRaw=Math.max(1,followupMaximum/5);
const followupMagnitude=Math.pow(10,Math.floor(Math.log10(followupStepRaw)));
const followupStep=([1,2,5,10].find(v=>v*followupMagnitude>=followupStepRaw)||10)*followupMagnitude;
drawCategoryChart('followupChart','followup',yearColours.followup,5*followupStep);
}
const STORAGE_KEY='appointment-dashboard-categories-v1';let categoryNames=['Urgent','Routine','Follow up'],assignments={};try{const saved=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');categoryNames=Array.isArray(saved.categories)&&saved.categories.length?saved.categories.filter(x=>typeof x==='string'&&x.trim()&&x!=='Uncategorised'):['Urgent','Routine','Follow up'];assignments=saved.assignments&&typeof saved.assignments==='object'?saved.assignments:{};}catch(e){}let clinicianCategoryNames=[],clinicianAssignments={};try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.clinicianCategories))clinicianCategoryNames=validCategoryNames(local.clinicianCategories);if(local.clinicianAssignments)clinicianAssignments=validAssignments(local.clinicianAssignments,clinicianCategoryNames)}catch(e){}let consultCategoryNames=[],consultAssignments={};try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.consultCategories))consultCategoryNames=validCategoryNames(local.consultCategories);if(local.consultAssignments)consultAssignments=validAssignments(local.consultAssignments,consultCategoryNames)}catch(e){}function ensureFollowUpCategory(){if(!categoryNames.some(c=>c.toLowerCase().replace(/[\s-]+/g,'')==='followup'))categoryNames.push('Follow up')}ensureFollowUpCategory();function guessCategory(t){return /RED|YELLOW|AMBER|URGENT|SAME DAY|ACUTE/i.test(t)?'Urgent':/GREEN|FOLLOW UP|ROUTINE/i.test(t)?'Routine':null}function saveCategories(sync=true){try{localStorage.setItem(STORAGE_KEY,JSON.stringify({categories:categoryNames,assignments,clinicianCategories:clinicianCategoryNames,clinicianAssignments,consultCategories:consultCategoryNames,consultAssignments}));}catch(e){alert('Could not save settings in this browser. Please try again.')}if(sync)markCategoryDirty()}function allCategories(){return [...categoryNames,'Uncategorised']}function rebuildCategoryFilter(){const old=new Set([...slots.selectedOptions].map(o=>o.value));slots.replaceChildren();for(const c of allCategories()){const o=document.createElement('option');o.value=c;o.textContent=c;o.selected=old.has(c);slots.append(o)}refresh()}function renderEditor(){renderDropEditor('urgency')}document.getElementById('addCategory').onclick=()=>{const el=document.getElementById('newCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||categoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This category name already exists or is reserved.');return}categoryNames.push(c);el.value='';saveCategories();rebuildCategoryFilter();renderEditor()};document.getElementById('newCategory').addEventListener('keydown',e=>{if(e.key==='Enter')document.getElementById('addCategory').click()});document.getElementById('slotSearch').addEventListener('input',renderEditor);function showPage(which){document.getElementById('chartPage').classList.toggle('hidden',which!=='chart');document.getElementById('categoryPage').classList.toggle('hidden',which!=='categories');document.getElementById('navChart').classList.toggle('active',which==='chart');document.getElementById('navCategories').classList.toggle('active',which==='categories');if(which==='chart')setTimeout(()=>{refresh();Plotly.Plots.resize('urgentChart');Plotly.Plots.resize('routineChart');Plotly.Plots.resize('followupChart')},0);else {renderEditor();renderClinicianEditor();renderConsultEditor()}}document.getElementById('navChart').onclick=()=>showPage('chart');document.getElementById('navCategories').onclick=()=>showPage('categories');document.getElementById('saveToOneDrive').onclick=()=>{void saveCloudCategories({manual:true})};
function renderClinicianEditor(){renderDropEditor('clinician')}
$('addClinicianCategory').onclick=()=>{const el=$('newClinicianCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||clinicianCategoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This clinician category name already exists or is reserved.');return}clinicianCategoryNames.push(c);el.value='';saveCategories();renderClinicianEditor()};
$('newClinicianCategory').addEventListener('keydown',e=>{if(e.key==='Enter')$('addClinicianCategory').click()});
$('clinicianSlotSearch').addEventListener('input',renderClinicianEditor);
function renderConsultEditor(){renderDropEditor('consult')}
$('addConsultCategory').onclick=()=>{const el=$('newConsultCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||consultCategoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This consult mode already exists or is reserved.');return}consultCategoryNames.push(c);el.value='';saveCategories();renderConsultEditor()};
$('newConsultCategory').addEventListener('keydown',e=>{if(e.key==='Enter')$('addConsultCategory').click()});
$('consultSlotSearch').addEventListener('input',renderConsultEditor);
rebuildCategoryFilter();renderEditor();renderClinicianEditor();renderConsultEditor();
refresh();

initAuth();

// Show one category classification at a time; the Save button stays visible.
(function setupCategoryTabs(){
 const buttons=Array.from(document.querySelectorAll('[data-category-tab]'));
 const panels=Array.from(document.querySelectorAll('.category-tab-panel'));
 function activate(target){
  for(const button of buttons){
   const selected=button.dataset.categoryTab===target;
   button.classList.toggle('active',selected);
   button.setAttribute('aria-selected',String(selected));
   button.tabIndex=selected?0:-1;
  }
  for(const panel of panels)panel.classList.toggle('hidden',panel.id!==target);
 }
 for(const button of buttons){
  button.addEventListener('click',()=>activate(button.dataset.categoryTab));
  button.addEventListener('keydown',event=>{
   if(event.key!=='ArrowRight'&&event.key!=='ArrowLeft')return;
   event.preventDefault();
   const offset=event.key==='ArrowRight'?1:-1;
   const next=buttons[(buttons.indexOf(button)+offset+buttons.length)%buttons.length];
   activate(next.dataset.categoryTab);next.focus();
  });
 }
 activate('tabUrgency');
})();
