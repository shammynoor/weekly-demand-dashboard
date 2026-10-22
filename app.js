// Initialise the availability filter before the first category-editor render.
let selectedAvailabilityUrgency='';
"use strict";
const CLIENT_ID='ba2bc895-4213-447c-ba62-d7edec6e8173';
const SCOPES=['Files.Read','Files.ReadWrite.AppFolder'];
const REDIRECT_URI='https://shammynoor.github.io/weekly-demand-dashboard/index.html';
const GRAPH='https://graph.microsoft.com/v1.0';
const FOLDER_KEY='weekly-demand-onedrive-folder-v1';
let records=[],weeks=[],types=[],statuses=[],clinicianTypes=[],bookingTimeRecords=[];
let availabilityRecords=[];
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
   consultAssignments=data.consultAssignments?validAssignments(data.consultAssignments,consultCategoryNames):{};lastConsultFilterNames=[];
   if(Array.isArray(data.clinicianGroupCategories)){clinicianGroupNames=validCategoryNames(data.clinicianGroupCategories);clinicianGroupAssignments=data.clinicianGroupAssignments?validAssignments(data.clinicianGroupAssignments,clinicianGroupNames):{};}
   if(Array.isArray(data.groupUrgencyCategories)){
    groupUrgencyNames=validCategoryNames(data.groupUrgencyCategories);
    groupUrgencyAssignments=data.groupUrgencyAssignments?validAssignments(data.groupUrgencyAssignments,groupUrgencyNames):{};
   }
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
  renderGroupUrgencyEditor();
  renderClinicianEditor();
  renderClinicianGroupEditor();
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
  clinicianGroupCategories:clinicianGroupNames.slice(),
  clinicianGroupAssignments:{...clinicianGroupAssignments},
  consultCategories:consultCategoryNames.slice(),
  consultAssignments:{...consultAssignments},
  groupUrgencyCategories:groupUrgencyNames.slice(),
  groupUrgencyAssignments:{...groupUrgencyAssignments},
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
function showRefreshFeedback(message,loading=false){
 const el=$('refreshFeedback');if(!el)return;
 el.textContent=message;el.classList.toggle('loading',loading);
 el.setAttribute('aria-busy',String(loading));
}
function resizeVisibleCharts(){
 if(!window.Plotly||$('chartPage').classList.contains('hidden')||$('dashboardShell').classList.contains('hidden'))return;
 for(const id of ['urgentChart','routineChart','followupChart']){
  const el=$(id);
  if(el&&el.offsetWidth>0&&el.classList.contains('js-plotly-plot')){
   try{Plotly.Plots.resize(el)}catch(e){console.warn('Chart resize:',e)}
  }
 }
}
function scheduleChartResize(){
 requestAnimationFrame(()=>requestAnimationFrame(resizeVisibleCharts));
}
window.addEventListener('resize',scheduleChartResize);


function setButtons(){
 // Never interpret a background chart/data refresh as a sign-out.
 // Recover the active MSAL account from its persistent cache if needed.
 if(!activeAccount&&auth){
  try{activeAccount=auth.getActiveAccount()||auth.getAllAccounts()[0]||null}catch(e){}
 }
 if(activeAccount)authenticatedOnce=true;
 // A completed sign-in is only cleared by the explicit Sign out action.
 // A graph redraw or temporary MSAL cache lookup must never reveal login.
 const logged=!!activeAccount||authenticatedOnce||$('dashboardShell').dataset.signedIn==='true';
 if(logged){authenticatedOnce=true;$('dashboardShell').dataset.signedIn='true'}
 if(logged){
  $('loginScreen').classList.add('hidden');
  $('dashboardShell').classList.remove('hidden');
 }else if($('dashboardShell').dataset.signedIn!=='true'){
  $('loginScreen').classList.remove('hidden');
  $('dashboardShell').classList.add('hidden');
 }
 $('signIn').disabled=logged;
 $('signOut').disabled=!logged;
 $('chooseFolder').disabled=!logged;
 $('refreshData').disabled=!logged||!selectedFolder||busy;
}
async function initAuth(){
 if(authInitPromise)return authInitPromise;
 // A reload can occur while large CSVs are parsed. If MSAL has a cached
 // account, do not flash the login page while redirect handling resolves.
 try{
  const cache=localStorage.getItem('msal.account.keys');
  if(cache&&cache!=='[]'){
   $('loginScreen').classList.add('hidden');
   $('dashboardShell').classList.remove('hidden');
   setSync('Restoring Microsoft session…');
  }
 }catch(e){}
 authInitPromise=(async()=>{
 try{
  if(!window.msal)throw Error('Microsoft sign-in library did not load. Please refresh with Ctrl+F5 or check browser content blocking.');
  if(!window.Plotly||!window.Papa)setSync('A chart or CSV library did not load. Refresh the page to retry.');
  auth=new msal.PublicClientApplication({auth:{clientId:CLIENT_ID,authority:'https://login.microsoftonline.com/consumers',redirectUri:REDIRECT_URI,navigateToLoginRequestUrl:false},cache:{cacheLocation:'localStorage',storeAuthStateInCookie:false}});
  if(typeof auth.initialize==='function')await auth.initialize();const result=await auth.handleRedirectPromise();
  activeAccount=result?.account||auth.getActiveAccount()||auth.getAllAccounts()[0]||activeAccount||null;
  if(activeAccount){
   authenticatedOnce=true;
   $('dashboardShell').dataset.signedIn='true';
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
function aggregateCSV(csv,aggregated,availabilityAggregate,bookingTimeAggregate){
 const parsed=Papa.parse(csv,{header:true,skipEmptyLines:'greedy'});
 const fields=parsed.meta.fields||[];
 for(const required of ['Appointment Date','Appointment Time','Booked Time to Slot Time','Slot Type','Availability'])if(!fields.includes(required))throw Error('Missing column: '+required);
 let valid=0,skipped=0;
 const fillColumns=[
  'Appointment Date',
  'Appointment Time',
  'User Type of the Session Holder of the Session',
  'Slot Type',
  'Availability',
  'Town of the Location of the Session',
  'Full Name of the Session Holder of the Session'
 ];
 // Fill down all seven exported fields, independently, within this CSV only.
 // Never carry a value from a previous CSV file.
 const previous=Object.fromEntries(fillColumns.map(column=>[column,'']));
 const recognized=new Set(['booked','available','blocked','embargoed']);
 for(const original of parsed.data){
  const r={...original};
  const explicitStatus=String(original['Availability']??'').trim();
  const explicitSlot=String(original['Slot Type']??'').trim();
  const leadRaw=String(original['Booked Time to Slot Time']??'').trim();
  const explicitTime=String(original['Appointment Time']??'').trim();
  // Count every non-empty CSV row after independently filling down the seven columns.
  for(const column of fillColumns){
   const value=String(r[column]??'').trim();
   if(value)previous[column]=value;
   else r[column]=previous[column];
  }
  const day=String(r['Appointment Date']??'').trim();
  const time=String(r['Appointment Time']??'').trim();
  const slotType=String(r['Slot Type']??'').trim();
  const availability=String(r['Availability']??'').trim();
  const status=availability.toLowerCase();
  if(!recognized.has(status)||!slotType)continue;
  const appointmentDay=parseDate(day);
  if(!appointmentDay)continue;
  const date=appointmentDay.toISOString().slice(0,10);
  const clinician=String(r['User Type of the Session Holder of the Session']??'').trim()||'Unknown clinician';
  const akey=JSON.stringify([date,slotType,availability,clinician]);
  availabilityAggregate.set(akey,(availabilityAggregate.get(akey)||0)+1);
  if(status!=='booked')continue;
  const lead=Number(leadRaw.replace(/,/g,''));
  const tm=/^(\d{1,2}):(\d{2})/.exec(time);
  if(!tm||!leadRaw||!Number.isFinite(lead)){skipped++;continue;}
  const dt=new Date(appointmentDay.getTime());
  dt.setUTCHours(Number(tm[1]),Number(tm[2]),0,0);
  const booked=new Date(dt.getTime()-lead*60000);
  if(!Number.isFinite(booked.getTime())){skipped++;continue;}
  const w=weekMonday(booked);
  if(w<'2025-07-07')continue;
  const bookedDay=booked.toISOString().slice(0,10);
  // Booking clock time, calculated using the same lead-time arithmetic as the demand charts.
  const bookedMinute=booked.getUTCHours()*60+booked.getUTCMinutes();
  const timeKey=JSON.stringify([bookedDay,bookedMinute,slotType,clinician]);
  bookingTimeAggregate.set(timeKey,(bookingTimeAggregate.get(timeKey)||0)+1);
  const key=JSON.stringify([bookedDay,slotType,clinician]);
  aggregated.set(key,(aggregated.get(key)||0)+1);valid++;
 }
 return {valid,skipped};
}
// Drag-and-drop category editor. Existing OneDrive JSON schema remains unchanged.
let pickedSlot=null;
function renderDropEditor(dimension){
 const cfg={
  urgency:{items:types,names:categoryNames,assign:assignments,boxes:'categoryNames',pool:'slotAssignments',search:'slotSearch',summary:'assignmentSummary',render:renderEditor},
  groupUrgency:{items:categoryNames,names:groupUrgencyNames,assign:groupUrgencyAssignments,boxes:'groupUrgencyNames',pool:'groupUrgencyAssignmentsPool',search:'groupUrgencySearch',summary:'groupUrgencySummary',render:renderGroupUrgencyEditor},
  clinician:{items:clinicianTypes,names:clinicianCategoryNames,assign:clinicianAssignments,boxes:'clinicianCategoryNames',pool:'clinicianSlotAssignments',search:'clinicianSlotSearch',summary:'clinicianAssignmentSummary',render:renderClinicianEditor},
  clinicianGroup:{items:clinicianCategoryNames,names:clinicianGroupNames,assign:clinicianGroupAssignments,boxes:'clinicianGroupNames',pool:'clinicianGroupAssignmentsPool',search:'clinicianGroupSearch',summary:'clinicianGroupSummary',render:renderClinicianGroupEditor},
  consult:{items:types,names:consultCategoryNames,assign:consultAssignments,boxes:'consultCategoryNames',pool:'consultSlotAssignments',search:'consultSlotSearch',summary:'consultAssignmentSummary',render:renderConsultEditor}
 }[dimension];
 if(!cfg)return;
 const boxes=$(cfg.boxes),pool=$(cfg.pool);if(!boxes||!pool)return;
 boxes.replaceChildren();pool.replaceChildren();
 const query=$(cfg.search).value.trim().toLowerCase();
 const filtered=cfg.items.filter(t=>t.toLowerCase().includes(query));
 const current=t=>cfg.names.includes(cfg.assign[t])?cfg.assign[t]:'Uncategorised';
 const place=(slot,category)=>{
  if(!cfg.items.includes(slot)||!(category==='Uncategorised'||cfg.names.includes(category)))return;
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
  zone.setAttribute('role','group');zone.setAttribute('aria-label','Drop items into '+name);
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
    if(dimension==='clinician'&&Object.prototype.hasOwnProperty.call(clinicianGroupAssignments,name)){clinicianGroupAssignments[next]=clinicianGroupAssignments[name];delete clinicianGroupAssignments[name];}
    if(dimension==='urgency'&&Object.prototype.hasOwnProperty.call(groupUrgencyAssignments,name)){groupUrgencyAssignments[next]=groupUrgencyAssignments[name];delete groupUrgencyAssignments[name];}
    if(dimension==='clinicianGroup'){if(disabledClinicianTypes.delete(name))disabledClinicianTypes.add(next);lastClinicianFilterNames=[];}
    for(const slot of Object.keys(cfg.assign))if(cfg.assign[slot]===name)cfg.assign[slot]=next;
    // Preserve the selected urgency filter when its category name changes.
    if(dimension==='urgency'){for(const option of slots.options)if(option.value===name)option.value=next;}
    if(dimension==='groupUrgency'&&selectedAvailabilityUrgency===name)selectedAvailabilityUrgency=next;
    saveCategories();if(dimension==='urgency'){rebuildCategoryFilter();renderGroupUrgencyEditor()}if(dimension==='clinician')renderClinicianGroupEditor();cfg.render();if(dimension==='groupUrgency')renderAvailabilityChart();if(dimension==='clinicianGroup')updateClinicianChartFilter();
   };
   head.append(rename);
   const del=document.createElement('button');del.type='button';del.className='drop-delete';del.textContent='Delete';
   del.setAttribute('aria-label','Delete '+name);
   del.onclick=e=>{e.stopPropagation();if(!confirm('Delete category “'+name+'”? Its slot types will become Uncategorised.'))return;
    if(dimension==='urgency'){categoryNames=categoryNames.filter(c=>c!==name);for(const t of cfg.items)if(assignments[t]===name)assignments[t]='Uncategorised';delete groupUrgencyAssignments[name]}
    else if(dimension==='groupUrgency'){groupUrgencyNames=groupUrgencyNames.filter(c=>c!==name);for(const t of cfg.items)if(groupUrgencyAssignments[t]===name)groupUrgencyAssignments[t]='Uncategorised'}
    else if(dimension==='clinician'){clinicianCategoryNames=clinicianCategoryNames.filter(c=>c!==name);for(const t of cfg.items)if(clinicianAssignments[t]===name)clinicianAssignments[t]='Uncategorised';delete clinicianGroupAssignments[name]}
    else if(dimension==='clinicianGroup'){clinicianGroupNames=clinicianGroupNames.filter(c=>c!==name);for(const t of cfg.items)if(clinicianGroupAssignments[t]===name)clinicianGroupAssignments[t]='Uncategorised'}
    else{consultCategoryNames=consultCategoryNames.filter(c=>c!==name);for(const t of cfg.items)if(consultAssignments[t]===name)consultAssignments[t]='Uncategorised'}
    if(dimension==='clinicianGroup'){disabledClinicianTypes.delete(name);lastClinicianFilterNames=[];}
    saveCategories();if(dimension==='urgency'){rebuildCategoryFilter();renderGroupUrgencyEditor()}if(dimension==='clinician')renderClinicianGroupEditor();cfg.render();if(dimension==='groupUrgency')renderAvailabilityChart();if(dimension==='clinicianGroup')updateClinicianChartFilter();
   };head.append(del);
  }
  zone.append(head);
  const body=document.createElement('div');body.className='drop-zone-body';
  const matching=filtered.filter(t=>current(t)===name);
  if(!matching.length){const hint=document.createElement('p');hint.className='drop-hint';hint.textContent=query?'No matching items':'Drop items here';body.append(hint)}
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
 $(cfg.summary).textContent=cfg.items.filter(t=>cfg.names.includes(cfg.assign[t])).length+' of '+cfg.items.length+(dimension==='groupUrgency'?' detailed urgency categories':dimension==='clinicianGroup'?' detailed clinician categories':dimension==='clinician'?' clinician user types':' slot types')+' assigned. Drag a slot icon into a category box (or select it, then click a box).';
}
async function loadFolder(force=true){if(!selectedFolder||busy)return;busy=true;setButtons();showRefreshFeedback('Refreshing data…',true);setSync('Checking OneDrive CSV files…');try{
 const items=(await children(selectedFolder.id)).filter(f=>f.file&&/\.csv$/i.test(f.name)).sort((a,b)=>a.name.localeCompare(b.name));if(!items.length)throw Error('No CSV files found directly inside this folder.');
 const fingerprint=JSON.stringify(items.map(f=>[f.id,f.lastModifiedDateTime,f.size]));if(!force&&fingerprint===lastFingerprint){showRefreshFeedback('Already up to date');setSync(`No changes detected · ${items.length} CSV files · Checked ${new Date().toLocaleTimeString('en-GB')}`);return}
 const aggregate=new Map(),availabilityAggregate=new Map(),bookingTimeAggregate=new Map();let valid=0,skipped=0;for(let i=0;i<items.length;i++){const f=items[i];showRefreshFeedback(`Loading file ${i+1} of ${items.length}…`,true);setSync(`Loading ${i+1} of ${items.length}: ${f.name}`);const detail=await graph(`/me/drive/items/${encodeURIComponent(f.id)}`);const url=detail['@microsoft.graph.downloadUrl'];if(!url)throw Error('OneDrive did not provide a download URL for '+f.name);const response=await fetch(url,{cache:'no-store'});if(!response.ok)throw Error(`Could not download ${f.name}: HTTP ${response.status}`);const stats=aggregateCSV(await response.text(),aggregate,availabilityAggregate,bookingTimeAggregate);valid+=stats.valid;skipped+=stats.skipped}
 bookingTimeRecords=[...bookingTimeAggregate].map(([key,count])=>{const [bookedDay,minute,slot,clinician]=JSON.parse(key);return {bookedDay,minute,slot,clinician,count}});
 availabilityRecords=[...availabilityAggregate].map(([key,count])=>{const [date,slot,availability,clinician]=JSON.parse(key);return {date,slot,availability,clinician,count}});
 records=[...aggregate].map(([key,count])=>{const [bookedDay,slot,clinician]=JSON.parse(key);const bookedWeek=mondayOf(bookedDay);return {week:bookedWeek,bookedWeek,bookedDay,slot,clinician,status:'Booked',count}});weeks=[...new Set(records.map(r=>r.bookedWeek))].sort();types=[...new Set([...records.map(r=>r.slot),...availabilityRecords.map(r=>r.slot)])].sort();statuses=['Booked'];clinicianTypes=[...new Set([...records.map(r=>r.clinician),...availabilityRecords.map(r=>r.clinician)])].filter(Boolean).sort();
 // Assign obvious names on first encounter; retain any existing manual choices.
 // Do not guess assignments: unknown slot types stay Uncategorised until manually assigned.

 saveCategories(false);lastFingerprint=fingerprint;
const drawingErrors=[];
 for(const [label,draw] of [
  ['category filters',()=>rebuildCategoryFilter()],
  ['urgency categories',()=>renderEditor()],
  ['clinician categories',()=>renderClinicianEditor()],
  ['clinician groups',()=>renderClinicianGroupEditor()],
  ['consult categories',()=>renderConsultEditor()],
  ['weekly charts',()=>refresh()],
  ['availability chart',()=>renderAvailabilityChart()],
  ['routine availability chart',()=>renderRoutineAvailabilityChart()],
  ['booking times chart',()=>renderBookingTimesChart()]
 ]){
  try{draw()}catch(error){drawingErrors.push(label+': '+(error?.message||String(error)));console.error('Dashboard rendering error in '+label,error)}
 }
 showRefreshFeedback(drawingErrors.length?'Data loaded; chart issue: '+drawingErrors.join('; '):'Data refreshed successfully');
 scheduleChartResize();setSync(`${items.length} CSV files · ${valid.toLocaleString('en-GB')} booked rows counted · ${skipped.toLocaleString('en-GB')} skipped · Updated ${new Date().toLocaleString('en-GB')}`)
 }catch(e){
 const message=e?.message||String(e);
 console.error('OneDrive refresh error',e);
 showRefreshFeedback('Refresh failed: '+message);
 setSync('Refresh failed: '+message+' · You remain signed in; the existing charts have been preserved where possible.');
 }finally{busy=false;setButtons()}}
$('signIn').onclick=async()=>{if(authenticatedOnce||activeAccount)return;try{if(!auth){await initAuth();if(!auth)throw Error('Microsoft sign-in is still unavailable. Refresh the page (Ctrl+F5).')}await auth.loginRedirect({scopes:SCOPES})}catch(e){setStatus('Sign-in failed: '+e.message)}};
$('signOut').onclick=async()=>{if(!activeAccount)return;try{authenticatedOnce=false;$('dashboardShell').dataset.signedIn='false';await auth.logoutRedirect({account:activeAccount,postLogoutRedirectUri:REDIRECT_URI})}catch(e){setStatus('Sign-out failed: '+e.message)}};
$('chooseFolder').onclick=showFolderPicker;
$('sourceSetup').onclick=showFolderPicker;
$('cancelFolder').onclick=()=>$('folderPicker').classList.add('hidden');
$('pickerUp').onclick=async()=>{if(!folderView.trail.length)return;const parent=folderView.trail.at(-1);folderView={...parent,trail:folderView.trail.slice(0,-1)};await showFolderContents()};
$('selectFolder').onclick=async()=>{selectedFolder={id:folderView.id,name:[...folderView.trail.map(x=>x.name),folderView.name].join(' / ')};localStorage.setItem(FOLDER_KEY,JSON.stringify(selectedFolder));$('folderStatus').textContent='Selected folder: '+selectedFolder.name;$('folderPicker').classList.add('hidden');$('onedrivePanel').classList.add('hidden');$('sourceSetup').classList.add('hidden');cloudFolder=selectedFolder;await saveCloudCategories();lastFingerprint='';await loadFolder(true)};
$('refreshData').onclick=()=>loadFolder(true);
setInterval(()=>{if(activeAccount&&selectedFolder&&!busy)loadFolder(false)},5*60*1000);
const holidayDates=['2025-01-01','2025-04-18','2025-04-21','2025-05-05','2025-05-26','2025-08-25','2025-12-25','2025-12-26','2026-01-01','2026-04-03','2026-04-06','2026-05-04','2026-05-25','2026-08-31','2026-12-25','2026-12-28'];
function mondayOf(dateString){const d=new Date(dateString+'T12:00:00Z');const day=d.getUTCDay();d.setUTCDate(d.getUTCDate()-((day+6)%7));return d.toISOString().slice(0,10)}
// Individual clinician category switches; every named category starts enabled.
const disabledClinicianTypes=new Set();
let lastClinicianFilterNames=[];
function updateClinicianChartFilter(){
 const root=$('clinicianChartFilters');if(!root)return;
 const names=[...clinicianGroupNames];
 // Do not recreate switches during each chart refresh; retain focus and scroll position.
 if(names.length===lastClinicianFilterNames.length&&names.every((n,i)=>n===lastClinicianFilterNames[i]))return;
 lastClinicianFilterNames=names;
 root.replaceChildren();
 if(!names.length){const note=document.createElement('span');note.className='hint';note.textContent='Create clinician groups in Manage categories.';root.append(note);return;}
 for(const name of names){
  const row=document.createElement('label');row.className='filter-row';
  const switchEl=document.createElement('span');switchEl.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.checked=!disabledClinicianTypes.has(name);input.setAttribute('aria-label','Include '+name);
  const track=document.createElement('span');track.className='switch-track';
  switchEl.append(input,track);
  const title=document.createElement('span');title.textContent=name;
  row.append(switchEl,title);root.append(row);
  input.addEventListener('change',()=>{if(input.checked)disabledClinicianTypes.delete(name);else disabledClinicianTypes.add(name);refresh()});
 }
}
function clinicianMatches(clinician){
 const detailed=clinicianAssignments[clinician];
 const group=clinicianGroupAssignments[detailed];
 // Only session-holder user types mapped through a detailed category to an enabled group count.
 return clinicianGroupNames.includes(group)&&!disabledClinicianTypes.has(group);
}
// Consultation-mode switches on the Weekly Demand Chart.
const disabledConsultModes=new Set();
let lastConsultFilterNames=[];
function updateConsultChartFilter(){
 const root=$('consultChartFilters');if(!root)return;
 const names=[...consultCategoryNames];
 if(names.length===lastConsultFilterNames.length&&names.every((n,i)=>n===lastConsultFilterNames[i]))return;
 lastConsultFilterNames=names;
 root.replaceChildren();
 if(!names.length){const note=document.createElement('span');note.className='hint';note.textContent='Create consult modes in Manage categories.';root.append(note);return;}
 for(const name of names){
  const row=document.createElement('label');row.className='filter-row';
  const switchEl=document.createElement('span');switchEl.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.checked=!disabledConsultModes.has(name);input.setAttribute('aria-label','Include '+name);
  const track=document.createElement('span');track.className='switch-track';
  switchEl.append(input,track);
  const title=document.createElement('span');title.textContent=name;
  row.append(switchEl,title);root.append(row);
  input.addEventListener('change',()=>{if(input.checked)disabledConsultModes.delete(name);else disabledConsultModes.add(name);refresh()});
 }
}
function consultMatches(slot){
 const mode=consultAssignments[slot];
 return consultCategoryNames.includes(mode)&&!disabledConsultModes.has(mode);
}
const bankHolidayWeeks=new Set(holidayDates.map(mondayOf));let excludeBankHolidays=true;
document.getElementById('excludeBH').onchange=e=>{excludeBankHolidays=!e.target.checked;refresh()};
function isoWeek(weekStart){const d=new Date(weekStart+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+3);const year=d.getUTCFullYear();const jan4=new Date(Date.UTC(year,0,4,12));jan4.setUTCDate(jan4.getUTCDate()-((jan4.getUTCDay()+6)%7));const week=1+Math.round((d-jan4)/604800000);return {year,week}}
const colours=['#1475b8','#e57a28','#159a78','#9053b4','#c64c69','#65748b'];
let current=[];let continuousMode=false;let showCurrentWeek=true;let monthlyMode=false;
let shadeSeasons=document.getElementById('shadeSeasons').checked;
document.getElementById('shadeSeasons').onchange=e=>{shadeSeasons=e.target.checked;refresh()};
const todayLondon=(()=>{const parts=new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());const v=Object.fromEntries(parts.map(p=>[p.type,p.value]));return `${v.year}-${v.month}-${v.day}`})();
const thisWeek=mondayOf(todayLondon),thisMonth=todayLondon.slice(0,7);
let savedWeeklyIncludeBankHolidays=document.getElementById('excludeBH').checked;
document.getElementById('periodMode').onchange=e=>{
 monthlyMode=e.target.checked;
 const bankToggle=document.getElementById('excludeBH');
 if(monthlyMode){savedWeeklyIncludeBankHolidays=bankToggle.checked;bankToggle.checked=true;bankToggle.disabled=true;excludeBankHolidays=false;}
 else{bankToggle.disabled=false;bankToggle.checked=savedWeeklyIncludeBankHolidays;excludeBankHolidays=!bankToggle.checked;}
 document.getElementById('periodModeValue').textContent=monthlyMode?'Month':'Week';
 document.getElementById('currentPeriodLabel').textContent=monthlyMode?'Show current month':'Show current week';
 refresh();
};
document.getElementById('currentWeekToggle').onchange=e=>{showCurrentWeek=e.target.checked;refresh()};
document.getElementById('displayMode').onchange=e=>{continuousMode=e.target.checked;document.getElementById('timelineValue').textContent=continuousMode?'Continuous':'Superimposed';refresh()};
function periodKey(r){return monthlyMode?r.bookedDay?.slice(0,7):r.bookedWeek}
function includedRecord(r){return r.status==='Booked'&&clinicianMatches(r.clinician)&&consultMatches(r.slot)&&(monthlyMode||!excludeBankHolidays||!bankHolidayWeeks.has(r.bookedWeek))}
function monthOfPeriod(p){return monthlyMode?p+'-01':p}
function periodIsCurrent(p){return p===(monthlyMode?thisMonth:thisWeek)}
function weeklyUrgencyMatches(slot,category){
 const detailed=assignments[slot]||'Uncategorised';
 if(category==='urgent'){
  const group=groupUrgencyAssignments[detailed];
  return typeof group==='string'&&group.trim().toLowerCase()==='urgent';
 }
 return detailed.trim().toLowerCase().replace(/[\s-]+/g,'')===category;
}
function refresh(){updateClinicianChartFilter();updateConsultChartFilter();
const activeWeeks=[...new Set(records.filter(includedRecord).map(periodKey).filter(Boolean))].sort();
const selectedWeeks=activeWeeks.filter(w=>w>=(monthlyMode?'2025-07':'2025-07-07')&&w<=(monthlyMode?thisMonth:thisWeek)&&(showCurrentWeek||!periodIsCurrent(w)));
const sums=new Map(selectedWeeks.map(w=>[w,0]));
for(const r of records){const key=periodKey(r);if(!sums.has(key)||!includedRecord(r))continue;sums.set(key,sums.get(key)+r.count)}
current=[...sums].map(([week,count])=>({week,count,...isoWeek(monthOfPeriod(week))}));
const total=current.reduce((a,r)=>a+r.count,0);
const yearColours={urgent:['#8b111e','#d43846','#f08088','#6f0712'],routine:['#075d34','#199c5a','#79c99c','#a1d9b4'],followup:['#124b9b','#2675d2','#79acee','#a9cdf6']};
function monthPosition(date){const d=new Date(date+'T12:00:00Z');const month=d.getUTCMonth();const days=new Date(Date.UTC(d.getUTCFullYear(),month+1,0)).getUTCDate();return month+(d.getUTCDate()-1)/days}
const axis={type:'linear',range:[0,12],title:'Month of year',tickmode:'array',tickvals:Array.from({length:12},(_,i)=>i+0.5),ticktext:['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],gridcolor:'#edf1f5'};
function seasonBackgroundShapes(periods){
 const shades=['#f3f5f7','#fafbfc','#f1f3f5','#f7f8fa'];
 const shapes=[];
 if(!continuousMode){
  const groups=[[[0,2],[11,12]],[[2,5]],[[5,8]],[[8,11]]];
  groups.forEach((intervals,season)=>intervals.forEach(([x0,x1])=>shapes.push({type:'rect',xref:'x',yref:'paper',x0,x1,y0:0,y1:1,fillcolor:shades[season],line:{width:0},layer:'below'})));
 }else{
  const years=periods.map(v=>Number(v.slice(0,4)));
  const start=Math.min(...years,2025)-1,end=Math.max(...years,2025)+1;
  for(let year=start;year<=end;year++){
   const bands=[[0,`${year}-01-01`,`${year}-03-01`],[1,`${year}-03-01`,`${year}-06-01`],[2,`${year}-06-01`,`${year}-09-01`],[3,`${year}-09-01`,`${year}-12-01`],[0,`${year}-12-01`,`${year+1}-01-01`]];
   for(const [season,x0,x1] of bands)shapes.push({type:'rect',xref:'x',yref:'paper',x0,x1,y0:0,y1:1,fillcolor:shades[season],line:{width:0},layer:'below'});
  }
 }
 return shapes;
}
function drawCategoryChart(target,category,palette,sharedMax){
 const enabled=true;
 const totals=new Map(selectedWeeks.map(w=>[w,0]));
 if(enabled)for(const r of records){
  const key=periodKey(r);if(!totals.has(key))continue;
  if(!includedRecord(r))continue;
  if(!weeklyUrgencyMatches(r.slot,category))continue;
  totals.set(key,totals.get(key)+r.count);
 }
 const points=[...totals].map(([date,count])=>({date,count,year:Number(date.slice(0,4)),week:monthlyMode?Number(date.slice(5,7)):isoWeek(date).week}));
 const years=[...new Set(points.map(p=>p.year))].sort();
 const latestYear=years.length?Math.max(...years):2026;
 const values=points.map(p=>p.count);
 const hasAppointments=values.some(v=>v>0);
 const average=values.length?values.reduce((sum,v)=>sum+v,0)/values.length:0;
 const traces=[];
 for(const [i,year] of years.entries()){
  const data=points.filter(p=>p.year===year).sort((a,b)=>a.date.localeCompare(b.date));
  const last=data[data.length-1];
  const partial=showCurrentWeek&&periodIsCurrent(last?.date);
  // Never draw a solid segment into an incomplete current week.
  const complete=partial?data.slice(0,-1):data;
  const colour=palette[i%palette.length];
  const width=year===latestYear?4.5:2;
  const opacity=year===latestYear?1:0.34;
  const x=p=>continuousMode?monthOfPeriod(p.date):(monthlyMode?Number(p.date.slice(5,7))-.5:monthPosition(p.date));
  traces.push({name:String(year),x:complete.map(x),y:complete.map(p=>p.count),customdata:complete.map(p=>p.date),type:'scatter',mode:'lines+markers',connectgaps:true,opacity,line:{color:colour,width,shape:'spline',smoothing:0.65},marker:{size:year===latestYear?5:3},hovertemplate:'%{customdata}<br><b>%{y:,} appointments</b><extra>'+year+'</extra>'});
  if(partial){
   const previous=complete[complete.length-1];
   const beforePrevious=complete[complete.length-2];
   // Sample a cubic Hermite curve for the final segment. Plotly's spline
   // becomes a straight line with only two points, so supply interpolated
   // positions while keeping the actual recorded endpoint unchanged.
   const segment=[];
   if(previous){
    const toNumber=p=>continuousMode?Date.parse(monthOfPeriod(p.date)):x(p);
    const x0=toNumber(previous),x1=toNumber(last);
    const span=x1-x0;
    const slope=beforePrevious&&x0!==toNumber(beforePrevious)
     ?(last.count-beforePrevious.count)/(x1-toNumber(beforePrevious))*span
     :(last.count-previous.count);
    const endSlope=last.count-previous.count;
    for(let k=0;k<=24;k++){
     const t=k/24,t2=t*t,t3=t2*t;
     const y=(2*t3-3*t2+1)*previous.count+(t3-2*t2+t)*slope+(-2*t3+3*t2)*last.count+(t3-t2)*endSlope;
     const xpos=x0+span*t;
     segment.push({x:continuousMode?new Date(xpos).toISOString():xpos,y, date:k===24?last.date:previous.date});
    }
   }else segment.push({x:x(last),y:last.count,date:last.date});
   traces.push({name:String(year)+' — partial '+(monthlyMode?'month':'week'),x:segment.map(p=>p.x),y:segment.map(p=>p.y),customdata:segment.map(p=>p.date),type:'scatter',mode:'lines',showlegend:false,opacity,line:{color:colour,width,dash:'dot',shape:'linear'},hoverinfo:'skip'});
   traces.push({name:String(year)+' — current '+(monthlyMode?'month':'week')+' (incomplete)',x:[x(last)],y:[last.count],customdata:[last.date],type:'scatter',mode:'markers',showlegend:false,opacity,marker:{size:11,symbol:'circle-open',color:colour,line:{color:colour,width:2.5}},hovertemplate:'%{customdata}<br><b>%{y:,} appointments</b><br>Current '+(monthlyMode?'month':'week')+' (incomplete)<extra>'+year+'</extra>'});
  }
 }
 // One flat average, calculated from every displayed week across every year.
 traces.push({name:'All-years '+(monthlyMode?'monthly':'weekly')+' average ('+Math.round(average).toLocaleString('en-GB')+')',x:continuousMode?[monthOfPeriod(selectedWeeks[0]||'2025-07-07'),monthOfPeriod(selectedWeeks[selectedWeeks.length-1]||'2025-07-14')]:[0,12],y:[average,average],type:'scatter',mode:'lines',hoverinfo:'skip',opacity:0.58,line:{color:category==='urgent'?'#9c2633':category==='followup'?'#286cb8':'#23885b',width:2,dash:'dash'},showlegend:true});
 const visibleDates=selectedWeeks.map(monthOfPeriod).sort();
 const continuousRange=visibleDates.length?[visibleDates[0],visibleDates[visibleDates.length-1]]:null;
 const chartEl=$(target);
 const chartWidth=Math.max(280,Math.floor(chartEl.parentElement.getBoundingClientRect().width)-2);
 Plotly.react(target,traces,{autosize:true,width:chartWidth,margin:{l:58,r:12,t:32,b:65},paper_bgcolor:'white',plot_bgcolor:'white',shapes:shadeSeasons?seasonBackgroundShapes(selectedWeeks):[],legend:{orientation:'h',y:1.17,x:0},xaxis:continuousMode?{type:'date',title:monthlyMode?'Booking month':'Booking date',tickformat:'%b %Y',gridcolor:'#edf1f5',...(continuousRange?{range:continuousRange}: {})}:axis,yaxis:{title:'Appointments',range:[0,sharedMax],dtick:sharedMax/5,gridcolor:'#edf1f5'},hovermode:'closest',annotations:!hasAppointments&&category==='followup'?[{text:'No follow-up appointments found. In Manage Categories, assign slot types to Follow up and click Save.',xref:'paper',yref:'paper',x:.5,y:.55,showarrow:false,align:'center',font:{size:13,color:'#286cb8'}}]:[]},{responsive:true,displaylogo:false});
}
let maximum=0;
let followupMaximum=0;

for(const category of ['urgent','routine','followup']){

 const totals=new Map(selectedWeeks.map(w=>[w,0]));
 for(const r of records){
  const key=periodKey(r);if(!totals.has(key))continue;
  if(!includedRecord(r))continue;
  if(!weeklyUrgencyMatches(r.slot,category))continue;
  totals.set(key,totals.get(key)+r.count);
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
let groupUrgencyNames=[],groupUrgencyAssignments={};
let clinicianGroupNames=[],clinicianGroupAssignments={};
const STORAGE_KEY='appointment-dashboard-categories-v1';let categoryNames=['Urgent','Routine','Follow up'],assignments={};try{const saved=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');categoryNames=Array.isArray(saved.categories)&&saved.categories.length?saved.categories.filter(x=>typeof x==='string'&&x.trim()&&x!=='Uncategorised'):['Urgent','Routine','Follow up'];assignments=saved.assignments&&typeof saved.assignments==='object'?saved.assignments:{};}catch(e){}let clinicianCategoryNames=[],clinicianAssignments={};try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.clinicianCategories))clinicianCategoryNames=validCategoryNames(local.clinicianCategories);if(local.clinicianAssignments)clinicianAssignments=validAssignments(local.clinicianAssignments,clinicianCategoryNames)}catch(e){}let consultCategoryNames=[],consultAssignments={};try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.consultCategories))consultCategoryNames=validCategoryNames(local.consultCategories);if(local.consultAssignments)consultAssignments=validAssignments(local.consultAssignments,consultCategoryNames)}catch(e){}try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.groupUrgencyCategories))groupUrgencyNames=validCategoryNames(local.groupUrgencyCategories);if(local.groupUrgencyAssignments)groupUrgencyAssignments=validAssignments(local.groupUrgencyAssignments,groupUrgencyNames)}catch(e){}
try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.clinicianGroupCategories))clinicianGroupNames=validCategoryNames(local.clinicianGroupCategories);if(local.clinicianGroupAssignments)clinicianGroupAssignments=validAssignments(local.clinicianGroupAssignments,clinicianGroupNames)}catch(e){}
function ensureFollowUpCategory(){if(!categoryNames.some(c=>c.toLowerCase().replace(/[\s-]+/g,'')==='followup'))categoryNames.push('Follow up')}ensureFollowUpCategory();function guessCategory(t){return /RED|YELLOW|AMBER|URGENT|SAME DAY|ACUTE/i.test(t)?'Urgent':/GREEN|FOLLOW UP|ROUTINE/i.test(t)?'Routine':null}function saveCategories(sync=true){lastConsultFilterNames=[];try{localStorage.setItem(STORAGE_KEY,JSON.stringify({categories:categoryNames,assignments,clinicianCategories:clinicianCategoryNames,clinicianAssignments,consultCategories:consultCategoryNames,consultAssignments,groupUrgencyCategories:groupUrgencyNames,groupUrgencyAssignments,clinicianGroupCategories:clinicianGroupNames,clinicianGroupAssignments}));}catch(e){alert('Could not save settings in this browser. Please try again.')}if(sync)markCategoryDirty()}function allCategories(){return [...categoryNames,'Uncategorised']}function rebuildCategoryFilter(){const old=new Set([...slots.selectedOptions].map(o=>o.value));slots.replaceChildren();for(const c of allCategories()){const o=document.createElement('option');o.value=c;o.textContent=c;o.selected=old.has(c);slots.append(o)}}function renderEditor(){renderDropEditor('urgency');updateAvailabilityUrgencyFilter()}
function renderGroupUrgencyEditor(){renderDropEditor('groupUrgency');updateAvailabilityUrgencyFilter()}document.getElementById('addCategory').onclick=()=>{const el=document.getElementById('newCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||categoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This category name already exists or is reserved.');return}categoryNames.push(c);el.value='';saveCategories();rebuildCategoryFilter();renderEditor();renderGroupUrgencyEditor()};document.getElementById('newCategory').addEventListener('keydown',e=>{if(e.key==='Enter')document.getElementById('addCategory').click()});document.getElementById('slotSearch').addEventListener('input',renderEditor);function showPage(which){
 for(const [page,nav,name] of [['chartPage','navChart','chart'],['categoryPage','navCategories','categories'],['availabilityPage','navAvailability','availability'],['routineAvailabilityPage','navRoutineAvailability','routineAvailability'],['bookingTimesPage','navBookingTimes','bookingTimes']]){
  $(page).classList.toggle('hidden',which!==name);$(nav).classList.toggle('active',which===name);
 }
 if(which==='chart')setTimeout(()=>{refresh();scheduleChartResize()},0);
 else if(which==='availability')setTimeout(renderAvailabilityChart,0);
 else if(which==='routineAvailability')setTimeout(renderRoutineAvailabilityChart,0);
 else if(which==='bookingTimes')setTimeout(renderBookingTimesChart,0);
 else {renderEditor();renderGroupUrgencyEditor();renderClinicianEditor();renderClinicianGroupEditor();renderConsultEditor()}
}document.getElementById('navChart').onclick=()=>showPage('chart');document.getElementById('navCategories').onclick=()=>showPage('categories');document.getElementById('navAvailability').onclick=()=>showPage('availability');document.getElementById('navRoutineAvailability').onclick=()=>showPage('routineAvailability');document.getElementById('navBookingTimes').onclick=()=>showPage('bookingTimes');document.getElementById('saveToOneDrive').onclick=()=>{void saveCloudCategories({manual:true})};
function renderClinicianEditor(){renderDropEditor('clinician');updateClinicianChartFilter()}
function renderClinicianGroupEditor(){renderDropEditor('clinicianGroup')}
$('addClinicianCategory').onclick=()=>{const el=$('newClinicianCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||clinicianCategoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This clinician category name already exists or is reserved.');return}clinicianCategoryNames.push(c);el.value='';saveCategories();renderClinicianEditor();renderClinicianGroupEditor()};
$('newClinicianCategory').addEventListener('keydown',e=>{if(e.key==='Enter')$('addClinicianCategory').click()});
$('clinicianSlotSearch').addEventListener('input',renderClinicianEditor);
$('addClinicianGroupCategory').onclick=()=>{const el=$('newClinicianGroupCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||clinicianGroupNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This clinician group name already exists or is reserved.');return}clinicianGroupNames.push(c);el.value='';saveCategories();renderClinicianGroupEditor();updateClinicianChartFilter()};
$('newClinicianGroupCategory').addEventListener('keydown',e=>{if(e.key==='Enter')$('addClinicianGroupCategory').click()});
$('clinicianGroupSearch').addEventListener('input',renderClinicianGroupEditor);
function renderConsultEditor(){renderDropEditor('consult')}
$('addConsultCategory').onclick=()=>{const el=$('newConsultCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||consultCategoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This consult mode already exists or is reserved.');return}consultCategoryNames.push(c);el.value='';saveCategories();renderConsultEditor()};
$('newConsultCategory').addEventListener('keydown',e=>{if(e.key==='Enter')$('addConsultCategory').click()});
$('consultSlotSearch').addEventListener('input',renderConsultEditor);
$('addGroupUrgencyCategory').onclick=()=>{const el=$('newGroupUrgencyCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||groupUrgencyNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This group name already exists or is reserved.');return}groupUrgencyNames.push(c);el.value='';saveCategories();renderGroupUrgencyEditor();renderAvailabilityChart()};
$('newGroupUrgencyCategory').addEventListener('keydown',e=>{if(e.key==='Enter')$('addGroupUrgencyCategory').click()});
$('groupUrgencySearch').addEventListener('input',renderGroupUrgencyEditor);
rebuildCategoryFilter();renderEditor();renderGroupUrgencyEditor();renderClinicianEditor();renderClinicianGroupEditor();renderConsultEditor();
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


function updateAvailabilityUrgencyFilter(){
 const select=$('availabilityUrgencyFilter');if(!select)return;
 const names=[...groupUrgencyNames];
 if(!names.includes(selectedAvailabilityUrgency))selectedAvailabilityUrgency=names[0]||'';
 const same=select.options.length===names.length&&names.every((name,i)=>select.options[i].value===name);
 if(!same){select.replaceChildren();for(const name of names){const opt=document.createElement('option');opt.value=name;opt.textContent=name;select.append(opt)}}
 select.value=selectedAvailabilityUrgency;
}
const urgentCategoryDisabled=new Set();
let urgentWeeksBack=1,urgentWeeksForward=3;
function setupUrgentRangeControls(){
 const back=$('urgentWeeksBack'),forward=$('urgentWeeksForward');
 if(!back||!forward)return;
 back.value=-urgentWeeksBack;forward.value=urgentWeeksForward;
 const label=offset=>offset===0?'This week':offset===-1?'Last week':offset<0?Math.abs(offset)+' weeks ago':offset===1?'Next week':offset+' weeks ahead';
 const update=(moved)=>{
  let start=Number(back.value),end=Number(forward.value);
  // Always include this week and prevent handles crossing.
  if(moved==='start')start=Math.min(start,0,end);
  if(moved==='end')end=Math.max(end,0,start);
  back.value=start;forward.value=end;
  urgentWeeksBack=-start;urgentWeeksForward=end;
  $('urgentWeeksBackValue').textContent=label(start);
  $('urgentWeeksForwardValue').textContent=label(end);
  const rail=$('urgentRangeActive');
  if(rail){rail.style.left=((start+3)/7*100)+'%';rail.style.width=((end-start)/7*100)+'%'}
  renderAvailabilityChart();
 };
 back.addEventListener('input',()=>update('start'));
 forward.addEventListener('input',()=>update('end'));
 update();
}
function urgentCategoryNames(){
 return categoryNames.filter(name=>String(groupUrgencyAssignments[name]||'').trim().toLowerCase()==='urgent');
}
function renderUrgentCategoryToggles(){
 const root=$('urgentCategoryToggles');if(!root)return;
 const names=urgentCategoryNames();
 const current=[...root.querySelectorAll('input[data-filter-name]')].map(x=>x.dataset.filterName);
 if(current.length===names.length&&current.every((x,i)=>x===names[i]))return;
 root.replaceChildren();
 if(!names.length){root.textContent='Assign Detailed urgency categories to the Urgent group in Manage categories.';return}
 for(const name of names){
  const row=document.createElement('label');row.className='filter-row';
  const sw=document.createElement('span');sw.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.dataset.filterName=name;input.checked=!urgentCategoryDisabled.has(name);
  const track=document.createElement('span');track.className='switch-track';
  const title=document.createElement('span');title.textContent=name;
  const colour=urgentColour(name);track.style.setProperty('--switch-on',colour);
  sw.append(input,track);row.append(sw,title);root.append(row);
  input.addEventListener('change',()=>{if(input.checked)urgentCategoryDisabled.delete(name);else urgentCategoryDisabled.add(name);renderAvailabilityChart()});
 }
}
function urgentColour(name){
 const n=name.toLowerCase();
 if(n.includes('red'))return '#d58d8a';
 if(n.includes('extra'))return '#a58acb';
 if(n.includes('priority'))return '#ead28c';
 const palette=['#88b7b0','#8baed0','#d6a9c5','#c7aa88','#9caec1'];
 let hash=0;for(const c of n)hash=(hash*31+c.charCodeAt(0))>>>0;
 return palette[hash%palette.length];
}
function renderAvailabilityChart(){
 const target=$('availabilityChart');
 if($('availabilityPage').classList.contains('hidden')||!window.Plotly||!target)return;
 renderUrgentCategoryToggles();
 const names=urgentCategoryNames().filter(n=>!urgentCategoryDisabled.has(n));
 const monday=new Date(thisWeek+'T12:00:00Z');const start=new Date(monday);start.setUTCDate(start.getUTCDate()-7*urgentWeeksBack);
 const weekCount=urgentWeeksBack+urgentWeeksForward+1;
 const dates=[];for(let w=0;w<weekCount;w++)for(let d=0;d<7;d++){const day=new Date(start);day.setUTCDate(start.getUTCDate()+w*7+d);dates.push(day.toISOString().slice(0,10))}
 const dateSet=new Set(dates);
 const weekdays=new Set(dates.filter(d=>{const n=new Date(d+'T12:00:00Z').getUTCDay();return n!==0&&n!==6}));
 const counts=new Map();
 for(const r of availabilityRecords){
  if(!weekdays.has(r.date))continue;
  const category=String(assignments[r.slot]||'').trim();
  if(!names.includes(category))continue;
  const status=String(r.availability||'').trim().toLowerCase();
  const label=status==='booked'?category:status==='blocked'?'Blocked':status==='embargoed'?'Embargoed':status==='available'?'Available':null;
  if(!label)continue;
  const key=JSON.stringify([r.date,label]);counts.set(key,(counts.get(key)||0)+r.count);
 }
 const labels=[...names,'Blocked','Embargoed','Available'];
 const colours=Object.fromEntries(names.map(n=>[n,urgentColour(n)]));
 Object.assign(colours,{Blocked:'#59636e',Embargoed:'#aab2bb',Available:'#ffffff'});
 const totals=dates.map(date=>labels.reduce((a,n)=>a+(counts.get(JSON.stringify([date,n]))||0),0));
 const max=Math.ceil(Math.max(1,...totals)*1.2/10)*10;
 const daySummary=dates.map(date=>{
  const weekday=new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'long',timeZone:'UTC'});
  return weekday+'<br>'+labels.map(n=>n+': '+(counts.get(JSON.stringify([date,n]))||0).toLocaleString('en-GB')).join('<br>')+'<br><b>Total: '+totals[dates.indexOf(date)].toLocaleString('en-GB')+'</b>';
 });
 const traces=labels.map(name=>({type:'bar',name,showlegend:false,x:dates,y:dates.map(date=>weekdays.has(date)?counts.get(JSON.stringify([date,name]))||0:null),marker:{color:colours[name],line:{color:name==='Available'?'#ffffff':'#e2e7ec',width:name==='Available'?0:1.1}},customdata:daySummary,hovertemplate:'%{customdata}<extra></extra>'}));
 for(const name of labels)traces.push({type:'scatter',mode:'markers',name,x:[null],y:[null],marker:{symbol:'circle',size:14,color:colours[name],line:{color:name==='Available'?'#aeb8c3':colours[name],width:name==='Available'?1.5:0}},hoverinfo:'skip',showlegend:true});
 const weekLabels=Array.from({length:weekCount},(_,i)=>{const offset=i-urgentWeeksBack;return offset===0?'This Week':offset===-1?'Last Week':offset<0?Math.abs(offset)+' weeks ago':offset===1?'Next Week':offset+' weeks'});
 const annotations=weekLabels.map((text,i)=>({x:dates[i*7+2],xref:'x',y:1.045,yref:'paper',text,showarrow:false,font:{size:12,color:'#444444'},xanchor:'center',yanchor:'bottom'}));
 const shapes=Array.from({length:weekCount},(_,i)=>({type:'line',xref:'x',yref:'paper',x0:i*7+5.5,x1:i*7+5.5,y0:0,y1:1,line:{color:'#9ca6b1',width:1.3}}));
 if(dateSet.has(todayLondon)){
  const x=dates.indexOf(todayLondon);
  shapes.unshift({type:'rect',xref:'x',yref:'paper',x0:-.5,x1:x,y0:0,y1:1,fillcolor:'rgba(224,117,117,0.20)',line:{width:0},layer:'below'});
  shapes.push({type:'line',xref:'x',yref:'paper',x0:x,x1:x,y0:0,y1:1,line:{color:'#4b5563',width:2,dash:'dot'}});
  annotations.push({x,xref:'x',y:1.005,yref:'paper',text:'Today',showarrow:false,font:{size:11,color:'#374151'},xanchor:'left',yanchor:'bottom',xshift:5});
 }
 dates.forEach((date,i)=>{if(weekdays.has(date)&&totals[i])annotations.push({x:date,xref:'x',y:totals[i],yref:'y',text:String(totals[i]),showarrow:false,font:{size:10,color:'#818b96'},yshift:4,yanchor:'bottom'})});
 const note=$('availabilityNote');if(note)note.textContent=names.length?'':'Select an Urgent detailed urgency category to display appointments.';
 Plotly.react(target,traces,{
  autosize:true,height:510,barmode:'stack',bargap:0.015,margin:{l:62,r:12,t:66,b:105},paper_bgcolor:'white',plot_bgcolor:'#e2e7ec',
  legend:{orientation:'h',x:0.5,y:-0.23,xanchor:'center',yanchor:'top',font:{size:12,color:'#444444'},bgcolor:'rgba(255,255,255,0)'},
  xaxis:{type:'category',categoryorder:'array',categoryarray:dates,tickmode:'array',tickvals:dates,ticktext:dates.map(date=>weekdays.has(date)?new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'short',timeZone:'UTC'}):''),tickangle:-90,tickfont:{size:12,color:'#444444'},showgrid:false,range:[-.5,dates.length-.5],fixedrange:true},
  yaxis:{title:'Appointment slots',range:[0,max],gridcolor:'#dce2e8'},annotations,shapes,hovermode:'closest'
 },{responsive:true,displaylogo:false});
}
window.addEventListener('resize',()=>{if(!$('availabilityPage').classList.contains('hidden'))renderAvailabilityChart()});
setupUrgentRangeControls();

// Independent filters for Routine Appointment Availability.
const routineDisabledClinicianGroups=new Set();
const routineDisabledConsultModes=new Set();
function renderRoutineAvailabilityFilters(){
 const build=(id,names,disabled,empty)=>{
  const root=$(id);if(!root)return;
  const previous=[...root.querySelectorAll('input')].map(el=>el.getAttribute('data-filter-name'));
  if(previous.length===names.length&&names.every((name,i)=>name===previous[i]))return;
  root.replaceChildren();
  if(!names.length){const note=document.createElement('span');note.className='hint';note.textContent=empty;root.append(note);return}
  for(const name of names){
   const row=document.createElement('label');row.className='filter-row';
   const switchEl=document.createElement('span');switchEl.className='switch';
   const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.setAttribute('data-filter-name',name);
   input.checked=!disabled.has(name);input.setAttribute('aria-label','Include '+name);
   const track=document.createElement('span');track.className='switch-track';
   const title=document.createElement('span');title.textContent=name;
   switchEl.append(input,track);row.append(switchEl,title);root.append(row);
   input.addEventListener('change',()=>{if(input.checked)disabled.delete(name);else disabled.add(name);renderRoutineAvailabilityChart()});
  }
 };
 build('routineClinicianFilters',clinicianGroupNames,routineDisabledClinicianGroups,'Create clinician groups in Manage categories.');
 build('routineConsultFilters',consultCategoryNames,routineDisabledConsultModes,'Create consult modes in Manage categories.');
}
function routineClinicianMatches(clinician){
 const detailed=clinicianAssignments[clinician];
 const group=clinicianGroupAssignments[detailed];
 return clinicianGroupNames.includes(group)&&!routineDisabledClinicianGroups.has(group);
}
function routineConsultMatches(slot){
 const mode=consultAssignments[slot];
 return consultCategoryNames.includes(mode)&&!routineDisabledConsultModes.has(mode);
}
function renderRoutineAvailabilityChart(){
 const target=$('routineAvailabilityChart');
 if(!target||$('routineAvailabilityPage').classList.contains('hidden')||!window.Plotly)return;
 renderRoutineAvailabilityFilters();
 // Monday of last week through Sunday six weeks after the current week: eight weeks.
 const start=new Date(thisWeek+'T12:00:00Z');start.setUTCDate(start.getUTCDate()-7);
 const weeks=[];
 for(let i=0;i<8;i++){const d=new Date(start);d.setUTCDate(start.getUTCDate()+7*i);weeks.push(d.toISOString().slice(0,10));}
 const byWeek=new Map();
 let matched=0;
 for(const r of availabilityRecords){
  const detailed=String(assignments[r.slot]||'').trim();
  const group=String(groupUrgencyAssignments[detailed]||'').trim();
  if(group.toLowerCase()!=='routine'||!routineClinicianMatches(r.clinician)||!routineConsultMatches(r.slot))continue;
  const week=mondayOf(r.date);
  if(!weeks.includes(week))continue;
  const status=String(r.availability||'').trim().toLowerCase();
  const label=({booked:'Booked',blocked:'Blocked',embargoed:'Embargoed',available:'Available'})[status];
  if(!label)continue;
  matched+=r.count;
  const key=JSON.stringify([week,label]);byWeek.set(key,(byWeek.get(key)||0)+r.count);
 }
 const labels=['Booked','Blocked','Embargoed','Available'];
 const colours={'Booked':'#a7d7b4','Blocked':'#59636e','Embargoed':'#aab2bb','Available':'#ffffff'};
 const dateText=d=>new Date(d+'T12:00:00Z').toLocaleDateString('en-GB',{day:'numeric',month:'short',timeZone:'UTC'});
 const weekLabels=weeks.map(w=>dateText(w));
 const summaries=weeks.map((week,i)=>'<b>Week commencing '+dateText(week)+'</b><br>'+labels.map(label=>label+': '+(byWeek.get(JSON.stringify([week,label]))||0).toLocaleString('en-GB')).join('<br>')+'<br><b>Total: '+labels.reduce((sum,label)=>sum+(byWeek.get(JSON.stringify([week,label]))||0),0).toLocaleString('en-GB')+'</b>');
 const plotBackground='#e0e4e8';
 const traces=labels.map(label=>({type:'bar',name:label,showlegend:false,x:weekLabels,y:weeks.map(w=>byWeek.get(JSON.stringify([w,label]))||0),marker:{color:colours[label],line:{color:label==='Available'?'rgba(0,0,0,0)':plotBackground,width:label==='Available'?0:1.5}},customdata:summaries,hovertemplate:'%{customdata}<extra></extra>'}));
 // Scatter traces provide circular legend symbols instead of Plotly bar squares.
 for(const label of labels)traces.push({type:'scatter',mode:'markers',name:label,x:[null],y:[null],marker:{symbol:'circle',size:14,color:colours[label],line:{color:label==='Available'?'#9ba7b2':colours[label],width:label==='Available'?1.5:0}},showlegend:true,hoverinfo:'skip'});
 const totals=weeks.map(w=>labels.reduce((sum,label)=>sum+(byWeek.get(JSON.stringify([w,label]))||0),0));
 const max=Math.max(1,...totals);
 const annotations=totals.flatMap((total,i)=>total?[{x:weekLabels[i],y:total,text:total.toLocaleString('en-GB'),showarrow:false,yshift:5,font:{size:11,color:'#768392'}}]:[]);
 const weekTitles=['Last Week','This Week','Next Week','2 weeks','3 weeks','4 weeks','5 weeks','6 weeks'];
 for(let i=0;i<weeks.length;i++)annotations.push({xref:'x',yref:'paper',x:i,y:1.105,text:weekTitles[i],showarrow:false,font:{family:'Arial',size:12,color:'#384858'}});
 // Each category occupies a full week from x=index-0.5 to x=index+0.5.
 // Place Today at 5%, 25%, 45%, 65% or 85% of its week for Mon–Fri.
 const dayIndex=(new Date(todayLondon+'T12:00:00Z').getUTCDay()+6)%7;
 const weekIndex=weeks.indexOf(thisWeek);
 const todayX=weekIndex<0?null:weekIndex-0.5+(Math.min(dayIndex,4)*0.2+0.05);
 const shapes=todayX===null?[]:[
  {type:'rect',xref:'x',yref:'paper',x0:-0.5,x1:todayX,y0:0,y1:1,fillcolor:'rgba(128,136,146,0.50)',line:{width:0},layer:'below'},
  {type:'line',xref:'x',yref:'paper',x0:todayX,x1:todayX,y0:0,y1:1,line:{color:'#344355',width:2,dash:'dot'},layer:'above'}
 ];
 for(let i=0;i<weeks.length-1;i++)shapes.push({type:'line',xref:'x',yref:'paper',x0:i+0.5,x1:i+0.5,y0:0,y1:1,line:{color:'#98a2ad',width:1,dash:'dot'},layer:'above'});
 if(todayX!==null)annotations.push({xref:'x',yref:'paper',x:todayX,y:1.02,text:'Today',showarrow:false,xanchor:'left',font:{size:12,color:'#344355'}});
 const note=$('routineAvailabilityNote');
 if(note)note.textContent=matched?'':"No matching Routine urgency-group appointments found. Check Manage categories → Urgency Group.";
 Plotly.react(target,traces,{
  height:510,autosize:true,barmode:'stack',bargap:0.12,
  margin:{l:185,r:24,t:85,b:35},paper_bgcolor:'#fff',plot_bgcolor:plotBackground,
  legend:{orientation:'v',x:-0.075,y:0.5,xanchor:'right',yanchor:'middle',font:{size:13,color:'#344355'},traceorder:'normal'},
  xaxis:{type:'category',categoryorder:'array',categoryarray:weekLabels,showticklabels:false,showgrid:false,showline:true,linecolor:'#374350',linewidth:3,zeroline:false},
  yaxis:{title:'Appointment slots',range:[0,Math.ceil(max*1.18/10)*10],gridcolor:'#dce2e8'},
  annotations,shapes,hovermode:'closest'
 },{responsive:true,displaylogo:false});
}
window.addEventListener('resize',()=>{if(!$('routineAvailabilityPage').classList.contains('hidden'))renderRoutineAvailabilityChart()});


window.addEventListener('unhandledrejection',event=>{
 const message=event.reason?.message||String(event.reason||'Unknown background error');
 console.error('Unhandled dashboard error',event.reason);
 if($('dashboardShell')&&!$('dashboardShell').classList.contains('hidden')){
  setSync('Background error: '+message+' · You are still signed in.');
 }
});
window.addEventListener('error',event=>{
 if($('dashboardShell')&&!$('dashboardShell').classList.contains('hidden')){
  setSync('Dashboard error: '+event.message+' · You are still signed in.');
 }
});


// Booking Times: independently switchable Urgency Groups, Clinician Groups and Consult modes.
const bookingTimesDisabledUrgency=new Set();
const bookingTimesDisabledClinician=new Set();
const bookingTimesDisabledConsult=new Set();
function bookingTimesTogglePanel(id,names,disabled){
 const root=$(id);if(!root)return;
 const current=[...root.querySelectorAll('input[data-filter-name]')].map(input=>input.dataset.filterName);
 if(current.length===names.length&&current.every((name,i)=>name===names[i]))return;
 root.replaceChildren();
 if(!names.length){const span=document.createElement('span');span.className='hint';span.textContent='Set up categories in Manage categories.';root.append(span);return}
 for(const name of names){
  const label=document.createElement('label');label.className='filter-row';
  const switchWrap=document.createElement('span');switchWrap.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');
  input.dataset.filterName=name;input.checked=!disabled.has(name);input.setAttribute('aria-label','Include '+name);
  const track=document.createElement('span');track.className='switch-track';
  const caption=document.createElement('span');caption.textContent=name;
  switchWrap.append(input,track);label.append(switchWrap,caption);root.append(label);
  input.addEventListener('change',()=>{if(input.checked)disabled.delete(name);else disabled.add(name);renderBookingTimesChart()});
 }
}
function renderBookingTimesChart(){
 const chart=$('bookingTimesChart');
 if(!chart||$('bookingTimesPage').classList.contains('hidden')||!window.Plotly)return;
 bookingTimesTogglePanel('bookingTimesUrgencyFilters',groupUrgencyNames,bookingTimesDisabledUrgency);
 bookingTimesTogglePanel('bookingTimesClinicianFilters',clinicianGroupNames,bookingTimesDisabledClinician);
 bookingTimesTogglePanel('bookingTimesConsultFilters',consultCategoryNames,bookingTimesDisabledConsult);
 const bins=Array(96).fill(0);
 let count=0;
 for(const r of bookingTimeRecords){
  const detailedUrgency=assignments[r.slot];
  const urgency=groupUrgencyAssignments[detailedUrgency];
  const detailedClinician=clinicianAssignments[r.clinician];
  const clinician=clinicianGroupAssignments[detailedClinician];
  const consult=consultAssignments[r.slot];
  if(!groupUrgencyNames.includes(urgency)||bookingTimesDisabledUrgency.has(urgency))continue;
  if(!clinicianGroupNames.includes(clinician)||bookingTimesDisabledClinician.has(clinician))continue;
  if(!consultCategoryNames.includes(consult)||bookingTimesDisabledConsult.has(consult))continue;
  if(!Number.isFinite(r.minute)||r.minute<0||r.minute>=1440)continue;
  bins[Math.floor(r.minute/15)]+=r.count;count+=r.count;
 }
 // Preserve the full-day denominator so the 50% marker means half of all bookings.
 const times=Array.from({length:96},(_,i)=>String(Math.floor(i/4)).padStart(2,'0')+':'+String((i%4)*15).padStart(2,'0'));
 let running=0;
 const cumulative=bins.map(n=>{running+=n;return count?running/count*100:0});
 const visibleStart=24; // 06:00
 const visibleEnd=84;   // 21:00 (exclusive for 15-minute bins)
 const visibleTimes=times.slice(visibleStart,visibleEnd);
 const visibleBins=bins.slice(visibleStart,visibleEnd);
 const visibleCumulative=cumulative.slice(visibleStart,visibleEnd);
 const tickvals=visibleTimes.filter((_,i)=>i%8===0);
 const medianBin=count?cumulative.findIndex(p=>p>=50):-1;
 const medianVisible=medianBin>=visibleStart&&medianBin<visibleEnd;
 const medianTime=medianVisible?times[medianBin]:null;
 const shapes=medianVisible?[{type:'line',xref:'x',yref:'y2',x0:medianTime,x1:medianTime,y0:0,y1:50,line:{color:'#636d78',width:1.8,dash:'dot'}}]:[];
 const annotations=medianVisible?[{xref:'x',yref:'y2',x:medianTime,y:50,text:'50% · '+medianTime,showarrow:false,xanchor:'left',yanchor:'bottom',xshift:6,font:{size:12,color:'#475569'}}]:[];
 const traces=[
  {type:'bar',name:'Bookings per 15 minutes',x:visibleTimes,y:visibleBins,marker:{color:'#8fb4d3',line:{width:0}},hovertemplate:'%{x}–15 min later<br>%{y:,} bookings<extra></extra>',yaxis:'y'},
  {type:'scatter',mode:'lines',name:'Cumulative %',x:visibleTimes,y:visibleCumulative,yaxis:'y2',line:{color:'#bc756e',width:3,shape:'hv'},hovertemplate:'%{x}<br>%{y:.1f}% booked by this time<extra></extra>'}
 ];
 Plotly.react(chart,traces,{
  height:520,autosize:true,margin:{l:24,r:85,t:70,b:65},
  paper_bgcolor:'#fff',plot_bgcolor:'#f1f3f5',bargap:0.04,
  legend:{orientation:'h',x:0,y:1.13,xanchor:'left',yanchor:'bottom'},
  xaxis:{type:'category',categoryorder:'array',categoryarray:visibleTimes,tickmode:'array',tickvals,ticktext:tickvals,title:'Calculated booking time (06:00–21:00)',showgrid:false},
  yaxis:{visible:false,rangemode:'tozero',showgrid:false,zeroline:false},
  yaxis2:{title:'Cumulative bookings (%)',overlaying:'y',side:'right',range:[0,100],ticksuffix:'%',showgrid:true,gridcolor:'#d9e1e8',zeroline:false},
  shapes,annotations,hovermode:'x unified'
 },{responsive:true,displaylogo:false});
 const note=$('bookingTimesNote');
 if(note)note.textContent=count?count.toLocaleString('en-GB')+' booked appointments across the loaded CSV files and selected categories.':'No bookings match these filters. Check category assignments in Manage categories.';
}
window.addEventListener('resize',()=>{if(!$('bookingTimesPage').classList.contains('hidden'))renderBookingTimesChart()});
