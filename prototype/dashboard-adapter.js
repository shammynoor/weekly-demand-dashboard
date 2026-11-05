// Integration boundary for the existing analytics engine. Definitions hoist above startup.
let iqSources=[],iqContributions=new Map(),iqSettingsRevision=0,iqUploading=false;
function initAuth(){
 window.addEventListener('load',()=>void iqStart(),{once:true});
 window.addEventListener('message',event=>{if(event.source===window.parent&&event.origin===location.origin&&event.data?.type==='iq-refresh')void refreshDashboardData(true);});
}
async function iqStart(){
 if(!iqBridge){document.body.textContent='Open this dashboard through your Darwin IQ practice workspace.';return;}
 try{
  await iqBridge.ensure();
  activeAccount={homeAccountId:iqBridge.user.id+':'+iqBridge.practice.id,username:iqBridge.user.email,name:iqBridge.user.email};
  isOwner=iqBridge.manage;folderAccessVerified=true;folderAccessDenied=false;authenticatedOnce=true;
  selectedFolder={id:iqBridge.practice.id,name:iqBridge.practice.name};
  applyAccessMode();setButtons();iqCreateUploadPanel();
  await loadCloudCategories({strict:true});await loadFolder(true);startDashboardIdleSession();
  $('signIn').onclick=()=>{};$('chooseFolder').onclick=()=>{};$('sourceSetup').onclick=()=>{};
  $('refreshData').onclick=()=>void refreshDashboardData(true);
 }catch(error){document.body.replaceChildren();const message=document.createElement('p');message.textContent='The practice dashboard could not load: '+error.message;document.body.append(message);}
}
async function cloudFile(){
 const data=await iqBridge.readSettings();iqSettingsRevision=data.revision;
 if(data.settings)return data.settings;
 const defaults=JSON.parse(analyticsSettingsSnapshot());defaults.practiceSettings={name:iqBridge.practice.name,anchors:{}};return defaults;
}
async function saveCloudCategories(){
 if(!isOwner||!cloudReady||cloudBusy)return false;
 cloudBusy=true;const serial=cloudChangeSerial;
 try{
  const snapshot=JSON.parse(analyticsSettingsSnapshot());delete snapshot.folder;validateCloudCategories(snapshot);
  iqSettingsRevision=await iqBridge.saveSettings(snapshot,iqSettingsRevision);
  categoryDirty=cloudChangeSerial!==serial;cloudNotice(categoryDirty?'Saving latest changes…':'Practice settings saved.');
  if(categoryDirty){clearTimeout(cloudSaveTimer);cloudSaveTimer=setTimeout(()=>void saveCloudCategories(),850);}
  return true;
 }catch(error){categoryDirty=true;cloudNotice('Settings not saved: '+error.message);return false;}
 finally{cloudBusy=false;}
}
async function openCSVCache(){return null;}
async function graph(){throw Error('Microsoft storage is disabled in this prototype.');}
async function token(){throw Error('Microsoft authentication is disabled in this prototype.');}
function showFolderPicker(){iqUploadElement?.scrollIntoView({behavior:'smooth'});}
async function logoutDashboard(){await iqBridge?.signOut();}
function returnToMicrosoftSignIn(){void logoutDashboard();}
async function renderPracticeUsers(){
 const status=$('practiceUsersStatus');if(!status)return;
 status.replaceChildren(document.createTextNode('Invitations, approvals and suspensions are managed in Users & approvals. '));
 const button=document.createElement('button');button.type='button';button.textContent='Manage users & approvals';button.style.width='auto';button.onclick=()=>iqBridge.users();status.append(button);
 const body=$('practiceUsersBody');if(body)body.replaceChildren();
}
async function removePracticeUser(){iqBridge.users();}
async function loadFolder(force=true){
 if(!selectedFolder||busy)return;busy=true;setButtons();showRefreshFeedback('Loading practice reports…',true);
 try{
 await iqBridge.ensure();
 iqSources=await iqBridge.list();
 const items=iqSources.map(source=>({id:source.source_key,name:source.kind==='slot'?'Slot Report '+source.source_name:'Signposting Auto Report '+source.source_name,lastModifiedDateTime:source.clinical_at,size:0}));
 const fingerprint=JSON.stringify(iqSources.map(s=>[s.source_key,s.revision]));
 const aggregate=new Map(),availabilityAggregate=new Map(),bookingTimeAggregate=new Map(),appointmentTimeAggregate=new Map(),waitingAggregate=new Map(),waitingTrendAggregate=new Map(),peopleAggregate=new Map(),newSlotTypeStats=new Map();
 let valid=0,skipped=0,newSlotFiles=0,newSignpostingFiles=0,downloaded=0,reused=0;
 const newSignposting=[],pendingCache=new Map();
 const newRapidHealth=await iqBridge.rapidSummary();
 for(const source of iqSources){
  let entry=iqContributions.get(source.source_key);
  if(!entry||entry.revision!==source.revision){entry={revision:source.revision,value:await iqBridge.read(source.source_key)};downloaded++;}else reused++;
  iqContributions.set(source.source_key,entry);
  const contribution={...entry.value,version:JSON.stringify([source.source_key,source.source_name,source.clinical_at,0])};pendingCache.set(source.source_key,contribution);
  if(contribution.kind==='signposting'){newSignposting.push(...contribution.events);newSignpostingFiles++;}
  else if(contribution.kind==='slot'){newSlotFiles++;valid+=contribution.valid;skipped+=contribution.skipped;const destinations=[aggregate,availabilityAggregate,bookingTimeAggregate,peopleAggregate,newSlotTypeStats,waitingAggregate,waitingTrendAggregate,appointmentTimeAggregate];for(let m=0;m<8;m++)mergeCSVCounts(destinations[m],contribution.maps[m]);}
 }
 for(const key of iqContributions.keys())if(!pendingCache.has(key))iqContributions.delete(key);
 updateHomeReportDates(items);iqRenderSources();
 /* REPORT_FINALIZATION */
function updateHomeReportDates(items){
 const display=kind=>{const rows=iqSources.filter(row=>row.kind===kind);if(!rows.length)return 'No report uploaded';const latest=Math.max(...rows.map(row=>Date.parse(row.clinical_at)));return new Date(latest).toLocaleString('en-GB',{dateStyle:'medium',timeStyle:'short',timeZone:'Europe/London'});};
 $('homeSlotModified').textContent=display('slot');$('homeSignpostingModified').textContent=display('signposting');
 $('homeRapidHealthModified').textContent=iqBridge.rapidDate()||'No report uploaded';
}
let iqUploadElement=null,iqSourceSelect=null,iqUploadStatus=null,iqDateInput=null;
function iqCreateUploadPanel(){
 if(!isOwner)return;
 const panel=document.createElement('section');panel.className='iq-import-panel';iqUploadElement=panel;
 const heading=document.createElement('h3');heading.textContent='Upload practice reports';
 const description=document.createElement('p');description.textContent='Choose or drop one Slot Report, Signposting Report or Rapid Health CSV. For a refreshed export, select its existing report source below to replace it. Rapid Health adds only newer request minutes. You can also import your existing settings JSON.';
 const input=document.createElement('input');input.type='file';input.accept='.csv,.json';input.setAttribute('aria-label','Upload practice CSV or settings JSON');
 const sourceLabel=document.createElement('label');sourceLabel.textContent='Report source to replace (Slot or Signposting)';iqSourceSelect=document.createElement('select');sourceLabel.append(iqSourceSelect);
 const dateLabel=document.createElement('label');dateLabel.textContent='Clinical export date and time (defaults to the file modified time)';iqDateInput=document.createElement('input');iqDateInput.type='datetime-local';dateLabel.append(iqDateInput);
 iqUploadStatus=document.createElement('p');iqUploadStatus.className='iq-import-status';iqUploadStatus.setAttribute('role','status');
 const list=document.createElement('ul');list.id='iqSources';list.className='iq-source-list';
 panel.append(heading,description,input,sourceLabel,dateLabel,iqUploadStatus,list);$('homePage').append(panel);
 input.addEventListener('change',()=>{if(input.files[0])void replaceExistingCSV(input.files[0]);input.value='';});iqRenderSources();
}
function iqRenderSources(){
 if(!iqSourceSelect)return;const selected=iqSourceSelect.value;iqSourceSelect.replaceChildren();const initial=document.createElement('option');initial.value='';initial.textContent='New source (uses the uploaded filename)';iqSourceSelect.append(initial);
 const list=$('iqSources');list.replaceChildren();
 for(const source of iqSources){
  const option=document.createElement('option');option.value=source.source_key;option.textContent=(source.kind==='slot'?'Slot: ':'Signposting: ')+source.source_name;iqSourceSelect.append(option);
  const row=document.createElement('li'),text=document.createElement('span'),remove=document.createElement('button');text.textContent=option.textContent+' · '+source.row_count.toLocaleString('en-GB')+' records · '+new Date(source.clinical_at).toLocaleString('en-GB');remove.type='button';remove.textContent='Remove';remove.onclick=async()=>{if(iqUploading)return;if(!confirm('Remove '+source.source_name+' from this practice’s analytics?'))return;try{await iqBridge.remove(source.source_key,source.revision);await loadFolder(true);}catch(error){iqUploadStatus.textContent=error.message;}};row.append(text,remove);list.append(row);
 }
 if([...iqSourceSelect.options].some(option=>option.value===selected))iqSourceSelect.value=selected;
}
async function replaceExistingCSV(file){
 if(iqUploading)return;
 if(!isOwner){csvDropStatus('Practice administrator access is required.',true);return;}
 iqUploading=true;csvDropBusy=true;
 try{
  if(!file||!file.size||file.size>50*1024*1024||! /\.(csv|json)$/i.test(file.name))throw Error('Choose one non-empty CSV or settings JSON up to 50 MB.');
  await iqBridge.ensure();
  const text=await file.text();
  if(/\.json$/i.test(file.name)){
   const settings=parseSettingsJSON(text);delete settings.folder;if(!confirm('Replace this practice’s shared settings with '+file.name+'? Personal display preferences are kept separately.'))return;
   iqSettingsRevision=await iqBridge.saveSettings(settings,iqSettingsRevision);await loadCloudCategories({strict:true});await loadFolder(true);iqUploadStatus.textContent='Practice settings imported.';return;
  }
  const headers=Papa.parse(text,{header:true,preview:1}).meta.fields||[];
  if(headers.includes('request_date')&&headers.includes('request_type')){await iqBridge.importRapid(file);await loadFolder(true);iqUploadStatus.textContent='Rapid Health import finished.';return;}
  const kind=headers.includes('Appointment Date')&&headers.includes('Availability')?'slot':headers.includes('Code Term')&&headers.includes('Date')?'signposting':null;
  if(!kind)throw Error('This is not a recognised Slot, Signposting or Rapid Health export.');
  const contribution=parseCSVContribution(text,kind==='slot'?'Slot Report.csv':'Signposting Auto Report.csv');
  const count=kind==='slot'?contribution.maps[1].size&&[...contribution.maps[1].values()].reduce((a,b)=>a+b,0):contribution.events.length;
  if(!count)throw Error('No valid records were found. The existing report has been preserved.');
  let existing=iqSources.find(row=>row.source_key===iqSourceSelect?.value);
  if(existing&&existing.kind!==kind)throw Error('Choose an existing report source of the same type.');
  if(!existing)existing=iqSources.find(row=>row.kind===kind&&row.source_name.toLowerCase()===file.name.toLowerCase());
  const clinical=iqDateInput?.value?new Date(iqDateInput.value):new Date(file.lastModified);
  if(!Number.isFinite(clinical.getTime()))throw Error('Set the export date and time.');
  const message=existing?'Replace '+existing.source_name+' with this refreshed export? Its previous counts will be replaced.':'Add '+file.name+' as a new report source? Use a new source only for a separate, non-overlapping reporting period; refreshed exports should replace an existing source.';
  if(!confirm(message))return;
  iqUploadStatus.textContent='Processing and saving '+file.name+'…';
  await iqBridge.save({existing,kind,name:existing?.source_name||file.name,clinical_at:clinical.toISOString(),count,contribution});
  await loadFolder(true);iqUploadStatus.textContent='Saved '+count.toLocaleString('en-GB')+' records for '+iqBridge.practice.name+'.';
 }catch(error){if(iqUploadStatus)iqUploadStatus.textContent='Upload failed: '+error.message;csvDropStatus('Upload failed: '+error.message,true);}
 finally{iqUploading=false;csvDropBusy=false;}
}
