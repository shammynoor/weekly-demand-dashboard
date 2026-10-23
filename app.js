// Initialise the availability filter before the first category-editor render.
let selectedAvailabilityUrgency='';
"use strict";
const CLIENT_ID='ba2bc895-4213-447c-ba62-d7edec6e8173';
const SCOPES=['Files.Read','Files.ReadWrite','Files.ReadWrite.AppFolder'];
const REDIRECT_URI='https://shammynoor.github.io/weekly-demand-dashboard/index.html';
const GRAPH='https://graph.microsoft.com/v1.0';
const FOLDER_KEY='weekly-demand-onedrive-folder-v1';
// A specific-person, view-only OneDrive share; not an authentication credential.
// Only accounts explicitly granted access in OneDrive can read this folder.
const SHARED_FOLDER_LINK='https://1drv.ms/f/c/07c73540cffc666d/IgDtAqPOppU4RI0NK-J0JwZwASLOpdds9ogr0rNNS18_X78?e=poJClm';
const OWNER_DRIVE_ID='07c73540cffc666d';
const SHARED_SETTINGS_FILE='weekly-demand-categories.json';
const SHARED_TOKEN='u!'+btoa(unescape(encodeURIComponent(SHARED_FOLDER_LINK))).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
let isOwner=false,sharedDriveId='',sharedFolderId='';
let folderAccessVerified=false,folderAccessDenied=false;
function sharedPath(){return '/shares/'+SHARED_TOKEN+'/driveItem'}
function sourceItemUrl(id){return isOwner?'/me/drive/items/'+encodeURIComponent(id):'/drives/'+encodeURIComponent(sharedDriveId)+'/items/'+encodeURIComponent(id)}
function applyAccessMode(){
 document.body.classList.toggle('viewer-mode',!isOwner);
 for(const id of ['navCategories','navPeople','chooseFolder','sourceSetup']){const el=document.getElementById(id);if(el)el.hidden=!isOwner}
 const label=document.getElementById('accessRoleLabel');if(label)label.textContent=isOwner?'Administrator':'View only';
}
async function identifyOwner(){
 const drive=await graph('/me/drive?$select=id');
 isOwner=String(drive.id||'').toLowerCase()===OWNER_DRIVE_ID.toLowerCase();
 applyAccessMode();
}
async function locateSharedFolder(){
 // Access must be confirmed by Microsoft Graph for this signed-in account.
 const root=await graph(sharedPath());
 if(!root.folder)throw Error('The sharing link does not point to a folder.');
 sharedDriveId=root.parentReference?.driveId||'';
 sharedFolderId=root.id||'';
 if(!sharedDriveId||!sharedFolderId)throw Error('Microsoft did not return the shared folder drive and item identifiers.');
 return {id:sharedFolderId,name:root.name||'Shared dashboard folder',shared:true};
}
async function readSharedSettings(){
 // Resolve the file from the folder's children, which works with shared personal OneDrive.
 // A sharing-link path followed by :/filename:/content is not consistently supported.
 const folder=await locateSharedFolder();
 const files=await children(folder.id);
 const settings=files.find(item=>item.file&&item.name.toLowerCase()===SHARED_SETTINGS_FILE.toLowerCase());
 if(!settings)return null;
 const detail=await graph(sourceItemUrl(settings.id));
 const downloadUrl=detail['@microsoft.graph.downloadUrl'];
 if(!downloadUrl)throw Error('Microsoft did not provide a download link for shared settings.');
 const response=await fetch(downloadUrl,{cache:'no-store'});
 if(!response.ok)throw Error('Shared settings download failed (HTTP '+response.status+')');
 const data=await response.json();
 if(!data||!Array.isArray(data.categories)||!data.assignments||typeof data.assignments!=='object')
  throw Error('Shared settings JSON does not contain valid categories and slot assignments.');
 if(!Object.keys(data.assignments).length)
  throw Error('Shared settings file contains no slot type assignments. Ask the administrator to check Slot settings.');
 return data;
}

let records=[],weeks=[],types=[],statuses=[],clinicianTypes=[],bookingTimeRecords=[];
let availabilityRecords=[];
let peopleRecords=[];
let peopleLocum={};
let peopleClinicianType={};
// One customisable colour per Clinician Group (never per detailed type).
let clinicianGroupColours={};
let clinicianTypeColours={};
let urgencyTypeColours={};

// Curated 18-colour spectrum: cohesive, deep pastel shades rather than unrestricted RGB.
const DASHBOARD_COLOUR_PALETTE=['#bd5473','#c75b65','#c9775a','#c98d4d','#c5a34a','#b4ad55','#8ea65c','#68a37a','#52a48f','#4e9e9e','#4e9eae','#568fb7','#6583bb','#7977b8','#9272b4','#aa70a8','#b86d93','#a75b80'];
function showDashboardColourPalette(anchor,input){
 document.querySelectorAll('.dashboard-colour-popover').forEach(el=>el.remove());
 const popup=document.createElement('div');popup.className='dashboard-colour-popover';popup.setAttribute('role','group');popup.setAttribute('aria-label','Choose from 18 coordinated colours');
 DASHBOARD_COLOUR_PALETTE.forEach((hex,i)=>{
  const b=document.createElement('button');b.type='button';b.className='dashboard-colour-swatch';b.style.backgroundColor=hex;b.title='Colour '+(i+1);b.setAttribute('aria-label','Select colour '+(i+1));
  if(input.value.toLowerCase()===hex)b.classList.add('selected');
  b.addEventListener('click',event=>{event.stopPropagation();input.value=hex;input.dispatchEvent(new Event('change',{bubbles:true}));popup.remove()});popup.append(b);
 });
 document.body.append(popup);
 const rect=anchor.getBoundingClientRect();const width=popup.offsetWidth;
 popup.style.left=Math.max(8,Math.min(window.innerWidth-width-8,rect.left))+'px';
 popup.style.top=(rect.bottom+popup.offsetHeight+10<window.innerHeight?rect.bottom+7:Math.max(8,rect.top-popup.offsetHeight-7))+'px';
 const dismiss=event=>{if(!popup.contains(event.target)&&!anchor.contains(event.target)){popup.remove();document.removeEventListener('pointerdown',dismiss,true)}};
 setTimeout(()=>document.addEventListener('pointerdown',dismiss,true),0);
}
function installDashboardColourPicker(wrap,input){
 // Keep the existing input/change handlers and saved-settings flow; replace native picker UI only.
 input.tabIndex=-1;input.setAttribute('aria-hidden','true');
 wrap.setAttribute('role','button');wrap.tabIndex=0;
 wrap.addEventListener('click',event=>{event.preventDefault();event.stopPropagation();showDashboardColourPalette(wrap,input)});
 wrap.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();showDashboardColourPalette(wrap,input)}});
}

const TYPE_PALETTE=['#d4777e','#9978c8','#e4b94f','#72a9d5','#6eaea6','#cf8eb7','#82b789','#c79c70'];
function typeColour(name,overrides){if(/^#[0-9a-f]{6}$/i.test(overrides[name]||''))return overrides[name];let hash=0;for(const c of String(name).toLowerCase())hash=(hash*31+c.charCodeAt(0))>>>0;return TYPE_PALETTE[hash%TYPE_PALETTE.length]}
function urgencyTypeColour(name){return typeColour(name,urgencyTypeColours)}
function clinicianTypeColour(name){return typeColour(name,clinicianTypeColours)}
function typeColourPicker(type,kind,onChanged){const colour=kind==='urgency'?urgencyTypeColours:clinicianTypeColours;const current=kind==='urgency'?urgencyTypeColour(type):clinicianTypeColour(type);const wrap=document.createElement('label');wrap.className='people-colour-control people-type-colour-control';wrap.title='Change colour for '+type;wrap.draggable=false;const dot=document.createElement('span');dot.className='clinician-colour-dot';dot.style.backgroundColor=current;const input=document.createElement('input');input.className='people-colour-picker';input.type='color';input.value=current;input.setAttribute('aria-label','Choose colour for '+type);input.addEventListener('change',()=>{colour[type]=input.value;onChanged()});wrap.append(dot,input);installDashboardColourPicker(wrap,input);return wrap}

const CLINICIAN_GROUP_PALETTE=['#78a9d1','#85bb99','#b49bd3','#dea37e','#80bdbf','#d493b0','#b8ae75','#8badd1','#b1a0ca','#c89e80'];
function clinicianGroupColour(group){
 if(!group||!clinicianGroupNames.includes(group))return '#aab5c1';
 if(/^#[0-9a-f]{6}$/i.test(clinicianGroupColours[group]||''))return clinicianGroupColours[group];
 const index=clinicianGroupNames.indexOf(group);
 return CLINICIAN_GROUP_PALETTE[index%CLINICIAN_GROUP_PALETTE.length];
}
function clinicianColourDot(group){
 const dot=document.createElement('span');dot.className='clinician-colour-dot';
 dot.style.backgroundColor=clinicianGroupColour(group);dot.title=group+' group colour';return dot;
}
function normalisePersonName(value){return String(value??'').trim().replace(/\\s+/g,' ')}
function personKey(name){return normalisePersonName(name).toLocaleLowerCase('en-GB')}
function suggestedLocum(type){
 // Conservative suggestions only: explicit employment descriptors, not clinical role.
 return /\\b(locum|agency|bank|temporary|temp staff)\\b/i.test(String(type||''))?'yes':'no';
}
function suggestedClinicianType(type){
 const assigned=clinicianAssignments[type];
 if(assigned&&clinicianCategoryNames.includes(assigned))return assigned;
 const exact=clinicianCategoryNames.find(n=>n.toLocaleLowerCase('en-GB')===String(type||'').trim().toLocaleLowerCase('en-GB'));
 return exact||'';
}
// Avoid scanning the entire people table for every appointment on every chart refresh.
let peopleTypeLookup=new Map();
function rebuildPeopleTypeLookup(){
 peopleTypeLookup=new Map(peopleRecords.map(p=>[personKey(p.name),p.type]));
}
function effectiveClinicianType(clinician,person){
 const key=personKey(person);
 if(key){
  const explicit=peopleClinicianType[key];
  if(explicit!==undefined){
   // Explicit Unassigned means use the original CSV user-type mapping.
   if(explicit&&clinicianCategoryNames.includes(explicit))return explicit;
  }else{
   const suggested=suggestedClinicianType(peopleTypeLookup.get(key)||clinician);
   if(suggested)return suggested;
  }
 }
 return clinicianAssignments[clinician]||'';
}
function refreshClinicianCharts(){
 if(!$('chartPage')?.classList.contains('hidden'))scheduleDemandRefresh();
 if(!$('availabilityPage')?.classList.contains('hidden'))renderAvailabilityChart();
 if(!$('routineAvailabilityPage')?.classList.contains('hidden'))renderRoutineAvailabilityChart();
 if(!$('bookingTimesPage')?.classList.contains('hidden'))renderBookingTimesChart();
}
// Slot settings uses the same compact group/type organiser as People settings.
let urgencyGroupColours={};
const URGENCY_GROUP_PALETTE=['#cf7e89','#83afca','#85b69b','#bd9cc9','#d3ad71','#79b9b8','#d49a83','#9daccf'];
function urgencyGroupColour(group){
 if(/^#[0-9a-f]{6}$/i.test(urgencyGroupColours[group]||''))return urgencyGroupColours[group];
 const i=groupUrgencyNames.indexOf(group);
 return i>=0?URGENCY_GROUP_PALETTE[i%URGENCY_GROUP_PALETTE.length]:'#aab5c1';
}
function renderSlotGroupBoxes(){
 const root=$('slotGroupBoxes');if(!root)return;
 root.replaceChildren();
 const groups=[...groupUrgencyNames];
 const make=(tag,cls,value)=>{const element=document.createElement(tag);if(cls)element.className=cls;if(value!==undefined)element.textContent=value;return element};
 const valid=(name,existing)=>name&&name.toLocaleLowerCase('en-GB')!=='uncategorised'&&!existing.some(x=>x.toLocaleLowerCase('en-GB')===name.toLocaleLowerCase('en-GB'));
 const changed=()=>{
  saveCategories();rebuildCategoryFilter();renderSlotGroupBoxes();renderUrgentCategoryToggles();
  // Maintain compatibility with existing slot type → detailed urgency editor.
  renderEditor();renderGroupUrgencyEditor();renderSlotTable();
  if(!$('availabilityPage').classList.contains('hidden'))renderAvailabilityChart();
  if(!$('routineAvailabilityPage').classList.contains('hidden'))renderRoutineAvailabilityChart();
  if(!$('chartPage').classList.contains('hidden'))scheduleDemandRefresh();
  if(!$('bookingTimesPage').classList.contains('hidden'))renderBookingTimesChart();
  const status=$('slotGroupSaveStatus');if(status)status.textContent='Changes are saving automatically to OneDrive.';
 };
 const editor=(initial,placeholder,onSave,onCancel)=>{
  const wrap=make('div','people-inline-editor');
  const input=make('input','people-group-add-input');input.value=initial;input.placeholder=placeholder;
  const save=make('button','people-mini-action','Save');save.type='button';
  const cancel=make('button','people-mini-action','Cancel');cancel.type='button';
  save.addEventListener('click',()=>onSave(input.value.trim()));
  cancel.addEventListener('click',onCancel||renderSlotGroupBoxes);
  input.addEventListener('keydown',event=>{if(event.key==='Enter')save.click();if(event.key==='Escape')cancel.click()});
  wrap.append(input,save,cancel);return {wrap,input};
 };
 const renameType=(oldName,newName)=>{
  if(oldName===newName){renderSlotGroupBoxes();return}
  if(!valid(newName,categoryNames)){alert('Please enter a unique detailed urgency type.');return}
  categoryNames=categoryNames.map(name=>name===oldName?newName:name);
  if(urgencyTypeColours[oldName]){urgencyTypeColours[newName]=urgencyTypeColours[oldName];delete urgencyTypeColours[oldName]}
  groupUrgencyAssignments[newName]=groupUrgencyAssignments[oldName]||'Uncategorised';
  delete groupUrgencyAssignments[oldName];
  for(const slot of Object.keys(assignments))if(assignments[slot]===oldName)assignments[slot]=newName;
  if(urgentCategoryDisabled.has(oldName)){urgentCategoryDisabled.delete(oldName);urgentCategoryDisabled.add(newName)}
  changed();
 };
 const addDrop=(list,card,group)=>{
  list.addEventListener('dragover',event=>{
   if(event.dataTransfer.types.includes('application/x-slot-urgency-type')||event.dataTransfer.types.includes('text/plain')){
    event.preventDefault();event.dataTransfer.dropEffect='move';card.classList.add('people-drop-target');
   }
  });
  list.addEventListener('dragleave',event=>{if(!list.contains(event.relatedTarget))card.classList.remove('people-drop-target')});
  list.addEventListener('drop',event=>{
   event.preventDefault();card.classList.remove('people-drop-target');
   const type=event.dataTransfer.getData('application/x-slot-urgency-type');
   if(!categoryNames.includes(type)||groupUrgencyAssignments[type]===group)return;
   groupUrgencyAssignments[type]=group;changed();
  });
 };
 const renderGroup=(group,isUnassigned=false)=>{
  const card=make('section','people-group-card slot-group-card');
  const types=categoryNames.filter(type=>isUnassigned?!groups.includes(groupUrgencyAssignments[type]):groupUrgencyAssignments[type]===group).sort((a,b)=>a.localeCompare(b));
  const header=make('div','people-group-heading');
  const title=make('h3','people-group-title',isUnassigned?'Unassigned':group);
  const count=make('span','people-group-count',types.length+' types');
  if(!isUnassigned){
   const colour=make('label','people-colour-control people-group-colour-control');colour.title='Change colour for '+group;
   const dot=make('span','clinician-colour-dot');dot.style.backgroundColor=urgencyGroupColour(group);
   const picker=make('input','people-colour-picker');picker.type='color';picker.value=urgencyGroupColour(group);
   picker.setAttribute('aria-label','Choose colour for urgency group '+group);
   picker.addEventListener('change',()=>{urgencyGroupColours[group]=picker.value;changed()});
   colour.append(dot,picker);installDashboardColourPicker(colour,picker);header.append(colour);
  }
  header.append(title,count);
  if(!isUnassigned){
   const actions=make('div','people-group-actions');
   const rename=make('button','people-icon-action','✎');rename.type='button';rename.title='Rename '+group;rename.setAttribute('aria-label','Rename urgency group '+group);
   rename.addEventListener('click',()=>{
    const {wrap,input}=editor(group,'Urgency group',value=>{
     if(value===group){renderSlotGroupBoxes();return}
     if(!valid(value,groupUrgencyNames)){alert('Please enter a unique urgency group.');return}
     groupUrgencyNames=groupUrgencyNames.map(name=>name===group?value:name);
     if(urgencyGroupColours[group]){urgencyGroupColours[value]=urgencyGroupColours[group];delete urgencyGroupColours[group]}
     for(const type of Object.keys(groupUrgencyAssignments))if(groupUrgencyAssignments[type]===group)groupUrgencyAssignments[type]=value;
     if(bookingTimesDisabledUrgency.has(group)){bookingTimesDisabledUrgency.delete(group);bookingTimesDisabledUrgency.add(value)}
     changed();
    });header.replaceChildren(wrap);input.focus();input.select();
   });
   const remove=make('button','people-icon-action people-delete-action','×');remove.type='button';remove.title='Delete '+group;remove.setAttribute('aria-label','Delete urgency group '+group);
   remove.addEventListener('click',()=>{
    if(!confirm('Delete urgency group "'+group+'"? Its '+types.length+' detailed types will become unassigned; no slot types or appointment records will be deleted.'))return;
    groupUrgencyNames=groupUrgencyNames.filter(name=>name!==group);delete urgencyGroupColours[group];
    for(const type of Object.keys(groupUrgencyAssignments))if(groupUrgencyAssignments[type]===group)groupUrgencyAssignments[type]='Uncategorised';
    bookingTimesDisabledUrgency.delete(group);changed();
   });actions.append(rename,remove);header.append(actions);
  }
  card.append(header);
  const list=make('div','people-group-types');addDrop(list,card,isUnassigned?'Uncategorised':group);
  for(const type of types){
   const row=make('div','people-group-type');row.draggable=true;row.title='Drag to another urgency group';
   row.addEventListener('dragstart',event=>{
    event.dataTransfer.setData('application/x-slot-urgency-type',type);event.dataTransfer.setData('text/plain',type);
    event.dataTransfer.effectAllowed='move';row.classList.add('people-dragging');
   });
   row.addEventListener('dragend',()=>row.classList.remove('people-dragging'));
   const grip=make('span','people-drag-grip','⠿');grip.setAttribute('aria-hidden','true');
   const label=make('span','people-type-label',type);
   const edit=make('button','people-icon-action','✎');edit.type='button';edit.title='Rename '+type;edit.setAttribute('aria-label','Rename urgency type '+type);
   edit.addEventListener('click',()=>{
    const {wrap,input}=editor(type,'Urgency type',value=>renameType(type,value));row.replaceChildren(wrap);input.focus();input.select();
   });
   const remove=make('button','people-icon-action people-delete-action','×');remove.type='button';remove.title='Delete '+type;remove.setAttribute('aria-label','Delete urgency type '+type);
   remove.addEventListener('click',()=>{
    if(!confirm('Delete detailed urgency type "'+type+'"? Slot types assigned to it will become uncategorised.'))return;
    categoryNames=categoryNames.filter(name=>name!==type);delete groupUrgencyAssignments[type];delete urgencyTypeColours[type];
    for(const slot of Object.keys(assignments))if(assignments[slot]===type)assignments[slot]='Uncategorised';
    urgentCategoryDisabled.delete(type);changed();
   });
   row.append(grip,typeColourPicker(type,'urgency',changed),label,edit,remove);list.append(row);
  }
  if(!types.length)list.append(make('p','people-empty-group','Drop a type here'));
  card.append(list);
  if(!isUnassigned){
   const add=make('button','people-add-type-button','+ Add type');add.type='button';
   add.addEventListener('click',()=>{
    const {wrap,input}=editor('','New urgency type',value=>{
     if(!valid(value,categoryNames)){alert('Please enter a unique detailed urgency type.');return}
     categoryNames.push(value);groupUrgencyAssignments[value]=group;changed();
    });add.replaceWith(wrap);input.focus();
   });card.append(add);
  }
  root.append(card);
 };
 for(const group of groups)renderGroup(group);
 const add=make('button','people-add-group-button','+ Add Group');add.type='button';
 add.addEventListener('click',()=>{
  const {wrap,input}=editor('','New urgency group',value=>{
   if(!valid(value,groupUrgencyNames)){alert('Please enter a unique urgency group.');return}
   groupUrgencyNames.push(value);changed();
  });add.replaceWith(wrap);input.focus();
 });root.append(add);
 if(categoryNames.some(type=>!groups.includes(groupUrgencyAssignments[type])))renderGroup('Uncategorised',true);
}
function renderPeopleGroups(){
 const root=$('peopleGroupBoxes');if(!root)return;
 root.replaceChildren();
 const groups=[...clinicianGroupNames];
 const el=(tag,cls,text)=>{const e=document.createElement(tag);if(cls)e.className=cls;if(text!==undefined)e.textContent=text;return e};
 const nameValid=(name,existing)=>name&&name.toLowerCase()!=='uncategorised'&&!existing.some(x=>x.toLowerCase()===name.toLowerCase());
 const changed=()=>{
  saveCategories();
  renderPeopleGroups();
  renderPeopleGroupFilter();
  renderPeople();
  renderClinicianEditor();
  renderClinicianGroupEditor();
  updateClinicianChartFilter();
  refreshClinicianCharts();
  const status=$('peopleSaveStatus');if(status)status.textContent='Changes are saving automatically to OneDrive.';
 };
 const editor=(initial,placeholder,onSave)=>{
  const wrap=el('div','people-inline-editor');
  const input=el('input','people-group-add-input');input.value=initial;input.placeholder=placeholder;
  const save=el('button','people-mini-action','Save');save.type='button';
  const cancel=el('button','people-mini-action','Cancel');cancel.type='button';
  save.addEventListener('click',()=>onSave(input.value.trim()));
  cancel.addEventListener('click',()=>renderPeopleGroups());
  input.addEventListener('keydown',e=>{if(e.key==='Enter')save.click();if(e.key==='Escape')cancel.click()});
  wrap.append(input,save,cancel);
  return {wrap,input};
 };
 const renameType=(oldName,newName)=>{
  if(newName===oldName){renderPeopleGroups();return}
  if(!nameValid(newName,clinicianCategoryNames)){alert('Please enter a unique detailed clinician type.');return}
  clinicianCategoryNames=clinicianCategoryNames.map(n=>n===oldName?newName:n);
  if(clinicianTypeColours[oldName]){clinicianTypeColours[newName]=clinicianTypeColours[oldName];delete clinicianTypeColours[oldName]}
  clinicianGroupAssignments[newName]=clinicianGroupAssignments[oldName];
  delete clinicianGroupAssignments[oldName];
  for(const key of Object.keys(clinicianAssignments))if(clinicianAssignments[key]===oldName)clinicianAssignments[key]=newName;
  for(const key of Object.keys(peopleClinicianType))if(peopleClinicianType[key]===oldName)peopleClinicianType[key]=newName;
  changed();
 };
 for(const group of groups){
  const card=el('section','people-group-card');
  const assigned=clinicianCategoryNames.filter(t=>clinicianGroupAssignments[t]===group).sort((a,b)=>a.localeCompare(b));
  const header=el('div','people-group-heading');
  const title=el('h3','people-group-title',group);
  const count=el('span','people-group-count',assigned.length+' types');
  const actions=el('div','people-group-actions');
  const rename=el('button','people-icon-action','✎');rename.type='button';rename.title='Rename group '+group;rename.setAttribute('aria-label','Rename clinician group '+group);
  rename.addEventListener('click',()=>{
   const {wrap,input}=editor(group,'Clinician group',value=>{
    if(value===group){renderPeopleGroups();return}
    if(!nameValid(value,clinicianGroupNames)){alert('Please enter a unique clinician group name.');return}
    clinicianGroupNames=clinicianGroupNames.map(n=>n===group?value:n);
    if(clinicianGroupColours[group]){clinicianGroupColours[value]=clinicianGroupColours[group];delete clinicianGroupColours[group]}
    for(const type of Object.keys(clinicianGroupAssignments))if(clinicianGroupAssignments[type]===group)clinicianGroupAssignments[type]=value;
    for(const disabled of [disabledClinicianTypes,urgentClinicianDisabled,routineDisabledClinicianGroups,bookingTimesDisabledClinician]){
     if(disabled.has(group)){disabled.delete(group);disabled.add(value)}
    }
    changed();
   });
   header.replaceChildren(wrap);input.focus();input.select();
  });
  const remove=el('button','people-icon-action people-delete-action','×');remove.type='button';remove.title='Delete group '+group;remove.setAttribute('aria-label','Delete clinician group '+group);
  remove.addEventListener('click',()=>{
   const count=clinicianCategoryNames.filter(t=>clinicianGroupAssignments[t]===group).length;
   if(!confirm('Delete clinician group "'+group+'"? '+count+' detailed clinician type(s) will become unassigned, but the types and people will not be deleted.'))return;
   clinicianGroupNames=clinicianGroupNames.filter(n=>n!==group);
   delete clinicianGroupColours[group];
   for(const type of Object.keys(clinicianGroupAssignments))if(clinicianGroupAssignments[type]===group)clinicianGroupAssignments[type]='Uncategorised';
   for(const disabled of [disabledClinicianTypes,urgentClinicianDisabled,routineDisabledClinicianGroups,bookingTimesDisabledClinician])disabled.delete(group);
   changed();
  });
  actions.append(rename,remove);
  const colourWrap=el('label','people-colour-control people-group-colour-control');
  colourWrap.title='Change colour for '+group+' group';
  const dot=clinicianColourDot(group);
  const picker=el('input','people-colour-picker');
  picker.type='color';picker.value=clinicianGroupColour(group);
  picker.setAttribute('aria-label','Choose colour for '+group+' clinician group');
  picker.addEventListener('change',()=>{
   clinicianGroupColours[group]=picker.value;
   saveCategories();renderPeopleGroups();renderPeople();
   renderUrgentSideFilters();renderRoutineAvailabilityFilters();refreshClinicianCharts();
  });
  colourWrap.append(dot,picker);installDashboardColourPicker(colourWrap,picker);
  header.append(colourWrap,title,count,actions);
  card.append(header);
  const list=el('div','people-group-types');
  list.dataset.group=group;
  list.addEventListener('dragover',event=>{if(event.dataTransfer.types.includes('text/plain')){event.preventDefault();event.dataTransfer.dropEffect='move';card.classList.add('people-drop-target')}});
  list.addEventListener('dragleave',event=>{if(!list.contains(event.relatedTarget))card.classList.remove('people-drop-target')});
  list.addEventListener('drop',event=>{
   event.preventDefault();card.classList.remove('people-drop-target');
   const type=event.dataTransfer.getData('text/plain');
   if(!clinicianCategoryNames.includes(type)||clinicianGroupAssignments[type]===group)return;
   clinicianGroupAssignments[type]=group;changed();
  });
  for(const type of assigned){
   const row=el('div','people-group-type');
   row.draggable=true;row.title='Drag to another clinician group';
   row.addEventListener('dragstart',event=>{event.dataTransfer.setData('text/plain',type);event.dataTransfer.effectAllowed='move';row.classList.add('people-dragging')});
   row.addEventListener('dragend',()=>row.classList.remove('people-dragging'));
   const grip=el('span','people-drag-grip','⠿');grip.setAttribute('aria-hidden','true');
   const label=el('span','people-type-label',type);
   const edit=el('button','people-icon-action','✎');edit.type='button';edit.title='Rename '+type;edit.setAttribute('aria-label','Rename '+type);
   edit.addEventListener('click',()=>{
    const {wrap,input}=editor(type,'Clinician type',value=>renameType(type,value));
    row.replaceChildren(wrap);input.focus();input.select();
   });
   const remove=el('button','people-icon-action people-delete-action','×');remove.type='button';remove.title='Delete '+type;remove.setAttribute('aria-label','Delete '+type);
   remove.addEventListener('click',()=>{
    if(!confirm('Delete clinician type "'+type+'"? Any people or CSV user types assigned to it will become unassigned.'))return;
    clinicianCategoryNames=clinicianCategoryNames.filter(n=>n!==type);delete clinicianTypeColours[type];
    delete clinicianGroupAssignments[type];
    for(const key of Object.keys(clinicianAssignments))if(clinicianAssignments[key]===type)clinicianAssignments[key]='Uncategorised';
    for(const key of Object.keys(peopleClinicianType))if(peopleClinicianType[key]===type)peopleClinicianType[key]='';
    changed();
   });
   row.append(grip,typeColourPicker(type,'clinician',changed),label,edit,remove);list.append(row);
  }
  if(!assigned.length)list.append(el('p','people-empty-group','Drop a type here'));
  card.append(list);
  const add=el('button','people-add-type-button','+ Add type');add.type='button';
  add.addEventListener('click',()=>{
   const {wrap,input}=editor('','New detailed type',value=>{
    if(!nameValid(value,clinicianCategoryNames)){alert('Please enter a unique detailed clinician type.');return}
    clinicianCategoryNames.push(value);clinicianGroupAssignments[value]=group;changed();
   });
   add.replaceWith(wrap);input.focus();
  });
  card.append(add);root.append(card);
 }
 const add=el('button','people-add-group-button','+ Add Group');add.type='button';
 add.addEventListener('click',()=>{
  const {wrap,input}=editor('','New clinician group',value=>{
   if(!nameValid(value,clinicianGroupNames)){alert('Please enter a unique clinician group.');return}
   clinicianGroupNames.push(value);changed();
  });
  add.replaceWith(wrap);input.focus();
 });
 root.append(add);
}
function renderSlotTableGroupFilter(){
 const select=$('slotGroupFilter');if(!select)return;
 const old=select.value||'';
 select.replaceChildren();
 for(const [value,label] of [['','All urgency groups'],...groupUrgencyNames.map(g=>[g,g])]){
  const option=document.createElement('option');option.value=value;option.textContent=label;select.append(option);
 }
 select.value=groupUrgencyNames.includes(old)?old:'';
}
function renderSlotTable(){
 const body=$('slotTableBody');if(!body)return;
 renderSlotTableGroupFilter();
 const query=String($('slotTableSearch')?.value||'').trim().toLowerCase();
 const filter=$('slotGroupFilter')?.value||'';
 const unassignedOnly=!!$('slotUnassignedOnly')?.checked;
 const rows=types.filter(slot=>{
  const assigned=categoryNames.includes(assignments[slot])?assignments[slot]:'';
  const group=assigned&&groupUrgencyNames.includes(groupUrgencyAssignments[assigned])?groupUrgencyAssignments[assigned]:'';
  if(unassignedOnly&&assigned)return false;
  if(!unassignedOnly&&filter&&group!==filter)return false;
  return !query||(slot+' '+assigned+' '+group).toLowerCase().includes(query);
 }).sort((a,b)=>a.localeCompare(b,undefined,{sensitivity:'base'}));
 body.replaceChildren();
 for(const slot of rows){
  const tr=document.createElement('tr');
  const slotCell=document.createElement('td');slotCell.textContent=slot;
  const assignedCell=document.createElement('td');assignedCell.className='people-assigned-cell';
  const select=document.createElement('select');select.className='people-clinician-select';select.setAttribute('aria-label','Assigned type for '+slot);
  for(const name of ['',...categoryNames]){const option=document.createElement('option');option.value=name;option.textContent=name||'Unassigned';select.append(option)}
  select.value=categoryNames.includes(assignments[slot])?assignments[slot]:'';
  select.addEventListener('change',()=>{
   assignments[slot]=select.value||'Uncategorised';
   saveCategories();rebuildCategoryFilter();renderEditor();renderSlotTable();
   updateAvailabilityUrgencyFilter();
   if(!$('availabilityPage').classList.contains('hidden'))renderAvailabilityChart();
   if(!$('routineAvailabilityPage').classList.contains('hidden'))renderRoutineAvailabilityChart();
   if(!$('chartPage').classList.contains('hidden'))scheduleDemandRefresh();
   if(!$('bookingTimesPage').classList.contains('hidden'))renderBookingTimesChart();
   const status=$('slotGroupSaveStatus');if(status)status.textContent='Changes are saving automatically to OneDrive.';
  });
  assignedCell.append(select);
  const groupCell=document.createElement('td');groupCell.className='people-clinician-group-cell';
  const assigned=categoryNames.includes(assignments[slot])?assignments[slot]:'';
  const group=assigned&&groupUrgencyNames.includes(groupUrgencyAssignments[assigned])?groupUrgencyAssignments[assigned]:'';
  const badge=document.createElement('span');badge.className='people-clinician-group-badge';
  const dot=document.createElement('span');dot.className='clinician-colour-dot';dot.style.backgroundColor=group?urgencyGroupColour(group):'#b6c0ca';
  badge.append(dot,document.createTextNode(group||'Unassigned'));groupCell.append(badge);
  tr.append(slotCell,assignedCell,groupCell);body.append(tr);
 }
 const count=$('slotTableCount');if(count)count.textContent=rows.length+' of '+types.length+' slot types shown';
}
function renderPeopleGroupFilter(){
 const select=$('peopleGroupFilter');if(!select)return;
 const previous=select.value||'';
 select.replaceChildren();
 for(const [value,label] of [['','All Clinician Groups'],...clinicianGroupNames.map(name=>[name,name])]){
  const option=document.createElement('option');option.value=value;option.textContent=label;select.append(option);
 }
 select.value=clinicianGroupNames.includes(previous)?previous:'';
}
function effectivePeopleClinicianType(person){
 const key=personKey(person.name);
 const saved=peopleClinicianType[key];
 // An explicit empty selection means unassigned; don't replace it with a suggestion.
 return saved!==undefined?(clinicianCategoryNames.includes(saved)?saved:''):suggestedClinicianType(person.type);
}
function renderPeople(){
 const body=$('peopleTableBody'),count=$('peopleCount');if(!body)return;
 const query=String($('peopleSearch')?.value||'').trim().toLowerCase();
 const group=$('peopleGroupFilter')?.value||'';
 const unassignedOnly=!!$('peopleUnassignedOnly')?.checked;
 const rows=peopleRecords.filter(p=>{
  const type=effectivePeopleClinicianType(p);
  const assignedGroup=type?clinicianGroupAssignments[type]||'':'';
  if(unassignedOnly){if(assignedGroup&&assignedGroup!=='Uncategorised')return false;}
  else if(group&&assignedGroup!==group)return false;
  return !query||(p.name+' '+p.type+' '+type+' '+assignedGroup).toLowerCase().includes(query);
 })
  .sort((a,b)=>a.type.localeCompare(b.type,undefined,{sensitivity:'base'})||a.name.localeCompare(b.name,undefined,{sensitivity:'base'}));
 body.replaceChildren();
 for(const person of rows){
  const key=personKey(person.name);
  const tr=document.createElement('tr');
  const name=document.createElement('td');name.textContent=person.name;
  const type=document.createElement('td');type.textContent=person.type||'Unknown';
  const assignedCell=document.createElement('td');assignedCell.className='people-assigned-cell';
  const assigned=document.createElement('select');assigned.className='people-clinician-select';
  assigned.setAttribute('aria-label','Assigned Clinician Type for '+person.name);
  const unassigned=document.createElement('option');unassigned.value='';unassigned.textContent='Unassigned';assigned.append(unassigned);
  for(const value of clinicianCategoryNames){
   const option=document.createElement('option');option.value=value;option.textContent=value;assigned.append(option);
  }
  const stored=peopleClinicianType[key];
  assigned.value=stored!==undefined&&clinicianCategoryNames.includes(stored)?stored:suggestedClinicianType(person.type);
  assigned.addEventListener('change',()=>{
   peopleClinicianType[key]=assigned.value;
   saveCategories();
   refreshClinicianCharts();
   renderPeople();
   const status=$('peopleSaveStatus');if(status)status.textContent='Changes are saving automatically to OneDrive.';
  });
  assignedCell.append(assigned);
  const groupCell=document.createElement('td');groupCell.className='people-clinician-group-cell';
  const selectedType=effectivePeopleClinicianType(person);
  const groupName=selectedType?clinicianGroupAssignments[selectedType]||'':'';
  const validGroup=clinicianGroupNames.includes(groupName);
  const groupBadge=document.createElement('span');groupBadge.className='people-clinician-group-badge';
  const groupDot=document.createElement('span');groupDot.className='clinician-colour-dot';
  groupDot.style.backgroundColor=validGroup?clinicianGroupColour(groupName):'#b6c0ca';
  groupBadge.append(groupDot,document.createTextNode(validGroup?groupName:'Unassigned'));
  groupCell.append(groupBadge);
  const locumCell=document.createElement('td');locumCell.className='people-locum-cell';
  const locum=document.createElement('select');locum.className='people-locum-select';
  locum.setAttribute('aria-label','Locum status for '+person.name);
  for(const [value,label] of [['no','No'],['yes','Yes']]){const option=document.createElement('option');option.value=value;option.textContent=label;locum.append(option)}
  locum.value=Object.prototype.hasOwnProperty.call(peopleLocum,key)?peopleLocum[key]:suggestedLocum(person.type);
  locum.addEventListener('change',()=>{
   peopleLocum[key]=locum.value;
   saveCategories();
   const status=$('peopleSaveStatus');if(status)status.textContent='Changes are saving automatically to OneDrive.';
  });
  locumCell.append(locum);tr.append(name,type,assignedCell,groupCell,locumCell);body.append(tr);
 }
 if(count)count.textContent=rows.length.toLocaleString('en-GB')+' of '+peopleRecords.length.toLocaleString('en-GB')+' people';
}

const slots={selectedOptions:[],replaceChildren(){},append(){},options:[]};
const $=id=>document.getElementById(id);
let auth,authInitPromise=null,activeAccount=null,authenticatedOnce=false,selectedFolder=null,folderView={id:'root',name:'OneDrive',trail:[]},busy=false,lastFingerprint='';
let cloudReady=false, cloudSaveTimer=null, cloudVersion=null, cloudBusy=false;
let cloudFolder=null;
const CLOUD_SETTINGS_NAME='weekly-demand-categories.json';
function cloudNotice(msg){const el=$('cloudCategoryStatus');if(el)el.textContent=msg}
async function cloudFile(){if(!isOwner)return readSharedSettings();try{const shared=await readSharedSettings();if(shared)return shared}catch(e){cloudNotice('Shared settings not available yet: '+e.message)}const t=await token();const r=await fetch(GRAPH+'/me/drive/special/approot:/'+CLOUD_SETTINGS_NAME+':/content',{headers:{Authorization:'Bearer '+t}});if(r.status===404)return null;if(!r.ok)throw Error('Settings read failed (HTTP '+r.status+')');return r.json()}
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
   peopleLocum=data.peopleLocum&&typeof data.peopleLocum==='object'&&!Array.isArray(data.peopleLocum)?Object.fromEntries(Object.entries(data.peopleLocum).filter(([key,value])=>typeof key==='string'&&(value==='yes'||value==='no'))):peopleLocum;
   peopleClinicianType=data.peopleClinicianType&&typeof data.peopleClinicianType==='object'&&!Array.isArray(data.peopleClinicianType)?data.peopleClinicianType:peopleClinicianType;
   if(data.clinicianTypeColours&&typeof data.clinicianTypeColours==='object')clinicianTypeColours=Object.fromEntries(Object.entries(data.clinicianTypeColours).filter(([k,v])=>/^#[0-9a-f]{6}$/i.test(v)));
   if(data.urgencyTypeColours&&typeof data.urgencyTypeColours==='object')urgencyTypeColours=Object.fromEntries(Object.entries(data.urgencyTypeColours).filter(([k,v])=>/^#[0-9a-f]{6}$/i.test(v)));
   if(data.clinicianGroupColours&&typeof data.clinicianGroupColours==='object'&&!Array.isArray(data.clinicianGroupColours))clinicianGroupColours=Object.fromEntries(Object.entries(data.clinicianGroupColours).filter(([k,v])=>typeof k==='string'&&/^#[0-9a-f]{6}$/i.test(v)));
   cloudFolder=data.folder&&typeof data.folder.id==='string'?data.folder:null;
   const names=validCategoryNames(data.categories);
   const remoteAssignments=validAssignments(data.assignments,names);
   clinicianCategoryNames=Array.isArray(data.clinicianCategories)?validCategoryNames(data.clinicianCategories):clinicianCategoryNames;
   clinicianAssignments=data.clinicianAssignments?validAssignments(data.clinicianAssignments,clinicianCategoryNames):clinicianAssignments;
   consultCategoryNames=Array.isArray(data.consultCategories)?validCategoryNames(data.consultCategories):[];
   consultAssignments=data.consultAssignments?validAssignments(data.consultAssignments,consultCategoryNames):{};lastConsultFilterNames=[];
   if(Array.isArray(data.clinicianGroupCategories)){clinicianGroupNames=validCategoryNames(data.clinicianGroupCategories);clinicianGroupAssignments=data.clinicianGroupAssignments?validAssignments(data.clinicianGroupAssignments,clinicianGroupNames):{};}
   if(data.urgencyGroupColours&&typeof data.urgencyGroupColours==='object'&&!Array.isArray(data.urgencyGroupColours))urgencyGroupColours=Object.fromEntries(Object.entries(data.urgencyGroupColours).filter(([key,value])=>typeof key==='string'&&/^#[0-9a-f]{6}$/i.test(value)));
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
  }else if(!isOwner)throw Error('No shared settings file. The owner must sign in and publish settings first.');
  else cloudNotice('No cloud settings yet. Changes will save automatically.');
  cloudReady=true;
  saveCategories(false);
  if(categoryDirty)markCategoryDirty();
  renderEditor();
  renderGroupUrgencyEditor();
  renderClinicianEditor();
  renderClinicianGroupEditor();
  renderConsultEditor();
  refresh();
  renderPeopleGroups();
  renderSlotGroupBoxes();
  renderPeople();
 }catch(e){
  cloudNotice('Could not load cloud categories: '+e.message+' (local settings retained)');
  cloudReady=false;
 }
}
let categoryDirty=false, cloudChangeSerial=0, cloudRetryTimer=null;
function markCategoryDirty(){
 if(!isOwner)return;
 categoryDirty=true;cloudChangeSerial++;
 cloudNotice('Saving settings automatically…');
 clearTimeout(cloudSaveTimer);
 if(activeAccount&&cloudReady)cloudSaveTimer=setTimeout(()=>{void saveCloudCategories()},850);
 else cloudNotice('Changes saved in this browser; OneDrive will sync when connected.');
}
function scheduleCloudSave(){markCategoryDirty()}

async function saveCloudCategories(options={}){
 if(!isOwner)return false;
 const manual=!!options.manual;
 if(!activeAccount||!cloudReady){
  if(manual)cloudNotice('Cannot save: sign in and wait for OneDrive settings to load.');
  return false;
 }
 if(cloudBusy){return false;}
 cloudBusy=true;
 const saveSerial=cloudChangeSerial;
 if(manual)cloudNotice('Saving all categories and slot assignments to OneDrive…');
 const snapshot=JSON.stringify({
  version:6,
  peopleLocum:{...peopleLocum},
  peopleClinicianType:{...peopleClinicianType},
  clinicianGroupColours:{...clinicianGroupColours},clinicianTypeColours:{...clinicianTypeColours},urgencyTypeColours:{...urgencyTypeColours},
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
  urgencyGroupColours:{...urgencyGroupColours},
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
  // Publish the same settings inside the shared CSV folder for read-only viewers.
  // Never save through a viewer's delegated Microsoft Graph session.
  if(selectedFolder?.id){
   const sharedSave=await fetch(GRAPH+'/me/drive/items/'+encodeURIComponent(selectedFolder.id)+':/'+encodeURIComponent(SHARED_SETTINGS_FILE)+':/content',{
    method:'PUT',headers:{Authorization:'Bearer '+t,'Content-Type':'application/json'},body:snapshot
   });
   if(!sharedSave.ok)throw Error('Shared settings publishing failed (HTTP '+sharedSave.status+')');
  }
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
  categoryDirty=true;
  cloudNotice('OneDrive autosave failed: '+e.message+'. Changes remain in this browser; retrying automatically.');
  clearTimeout(cloudRetryTimer);
  cloudRetryTimer=setTimeout(()=>{if(categoryDirty&&cloudReady&&activeAccount)void saveCloudCategories()},10000);
  return false;
 }finally{
  cloudBusy=false;
  if(categoryDirty&&cloudReady&&activeAccount&&!cloudRetryTimer){
   clearTimeout(cloudSaveTimer);
   cloudSaveTimer=setTimeout(()=>{void saveCloudCategories()},350);
  }
  if(!categoryDirty){clearTimeout(cloudRetryTimer);cloudRetryTimer=null;}
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
 if(logged)authenticatedOnce=true;
 // Never reveal the dashboard until Microsoft has verified folder access.
 if(logged&&folderAccessVerified&&!folderAccessDenied){
  $('dashboardShell').dataset.signedIn='true';
  $('loginScreen').classList.add('hidden');
  $('dashboardShell').classList.remove('hidden');
 }else if($('dashboardShell').dataset.signedIn!=='true'){
  $('loginScreen').classList.remove('hidden');
  $('dashboardShell').classList.add('hidden');
 }
 $('signIn').disabled=logged&&!folderAccessDenied;
 $('signIn').textContent=folderAccessDenied?'Sign in with a different Microsoft account':'Sign in with Microsoft';
 $('signOut').disabled=!logged;
 $('chooseFolder').disabled=!logged||!isOwner;
 $('refreshData').disabled=!logged||!selectedFolder||busy;
}
async function initAuth(){
 if(authInitPromise)return authInitPromise;
 // A reload can occur while large CSVs are parsed. If MSAL has a cached
 // account, do not flash the login page while redirect handling resolves.
 try{
  const cache=localStorage.getItem('msal.account.keys');
  if(cache&&cache!=='[]'){
   setStatus('Checking your Microsoft account and OneDrive access…');
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
   auth.setActiveAccount(activeAccount);
   folderAccessVerified=false;folderAccessDenied=false;
   setStatus('Checking your OneDrive folder permissions…');
   setButtons();
   await identifyOwner();
   if(isOwner){
    try{selectedFolder=JSON.parse(localStorage.getItem(FOLDER_KEY)||'null')}catch{}
   }else{
    selectedFolder=await locateSharedFolder();
   }
   await loadCloudCategories();
   if(!isOwner&&!cloudReady)throw Error('Shared settings could not be loaded. '+($('cloudCategoryStatus')?.textContent||'Ask the owner to publish them.'));
   folderAccessVerified=true;folderAccessDenied=false;setButtons();
   // Existing browser migrates its chosen folder to OneDrive settings once.
   if(isOwner){if(cloudFolder){selectedFolder=cloudFolder;localStorage.setItem(FOLDER_KEY,JSON.stringify(selectedFolder))}
   else if(selectedFolder&&cloudReady){cloudFolder=selectedFolder;await saveCloudCategories()}
   // Publish shared settings even if this account has not made edits recently.
   if(selectedFolder&&cloudReady)await saveCloudCategories();}
   setStatus('Signed in: '+(activeAccount.username||'Microsoft account'));
   $('accountLabel').textContent=activeAccount.username||'Microsoft account';
   if(selectedFolder){$('folderStatus').textContent='Source: '+selectedFolder.name;await loadFolder(false)}
   else if(isOwner){setSync('No source folder saved yet. Use Slot settings → Source folder setup.');$('sourceSetup').classList.remove('hidden')}
  }else setStatus('Sign in with the Hotmail account that owns your OneDrive files.');
 }catch(e){
  const message=String(e?.message||e);
  // Only a confirmed 401/403 or a sharing-link denial is an access rejection.
  // Network outages and missing settings should not be mislabelled as unauthorised.
  const permissionError=/\b(401|403)\b|access.?denied|permission.?denied|not authorised|not authorized|sharing link.*(invalid|access)|does not have access|itemNotFound|item not found/i.test(message);
  if(!isOwner&&permissionError){
   folderAccessDenied=true;folderAccessVerified=false;
   setStatus('You do not have access to this dashboard. Ask the folder owner to share the OneDrive folder with your Microsoft account, then sign in again.');
  }else{
   folderAccessVerified=false;
   setStatus('Unable to open dashboard: '+message);
  }
  if(auth){try{activeAccount=activeAccount||auth.getActiveAccount()||auth.getAllAccounts()[0]||null}catch(_){}}
 }finally{setButtons()}
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
async function children(id){if(!isOwner&&id===sharedFolderId){const items=[];let url=sharedPath()+'/children?$top=200';while(url){const r=await graph(url);items.push(...(r.value||[]));url=r['@odata.nextLink']||null}return items;}let url=id==='root'?'/me/drive/root/children?$top=200&$select=id,name,folder,file,size,lastModifiedDateTime':`/me/drive/items/${encodeURIComponent(id)}/children?$top=200&$select=id,name,folder,file,size,lastModifiedDateTime`;const items=[];while(url){const r=await graph(url);items.push(...(r.value||[]));url=r['@odata.nextLink']||null}return items}
async function showFolderPicker(){$('onedrivePanel').classList.remove('hidden');folderView={id:'root',name:'OneDrive',trail:[]};$('folderPicker').classList.remove('hidden');await showFolderContents()}
async function showFolderContents(){try{$('pickerPath').textContent=[...folderView.trail.map(x=>x.name),folderView.name].join(' / ');$('pickerEntries').textContent='Loading folders…';const folders=(await children(folderView.id)).filter(x=>x.folder).sort((a,b)=>a.name.localeCompare(b.name));$('pickerEntries').replaceChildren();if(!folders.length)$('pickerEntries').textContent='No subfolders. You can select this folder.';for(const f of folders){const btn=document.createElement('button');btn.className='folder-entry';btn.textContent='📁 '+f.name;btn.onclick=async()=>{folderView={id:f.id,name:f.name,trail:[...folderView.trail,{id:folderView.id,name:folderView.name}]};await showFolderContents()};$('pickerEntries').append(btn)}}catch(e){$('pickerEntries').textContent='Could not list folders: '+e.message}}
function weekMonday(d){const x=new Date(d);x.setUTCHours(12,0,0,0);x.setUTCDate(x.getUTCDate()-((x.getUTCDay()+6)%7));return x.toISOString().slice(0,10)}
function parseDate(s){const m=/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec((s||'').trim());if(!m)return null;const mi=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].toLowerCase());if(mi<0)return null;const d=new Date(Date.UTC(Number(m[3]),mi,Number(m[1])));return d.getUTCMonth()===mi&&d.getUTCDate()===Number(m[1])?d:null}
function aggregateCSV(csv,aggregated,availabilityAggregate,bookingTimeAggregate,peopleAggregate){
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
  const person=normalisePersonName(r['Full Name of the Session Holder of the Session']);
  if(person&&peopleAggregate){const key=JSON.stringify([person,clinician]);peopleAggregate.set(key,(peopleAggregate.get(key)||0)+1)}
  const akey=JSON.stringify([date,slotType,availability,clinician,person]);
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
  const timeKey=JSON.stringify([bookedDay,bookedMinute,slotType,clinician,person]);
  bookingTimeAggregate.set(timeKey,(bookingTimeAggregate.get(timeKey)||0)+1);
  const key=JSON.stringify([bookedDay,slotType,clinician,person]);
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
    saveCategories();if(dimension==='urgency'){rebuildCategoryFilter();renderGroupUrgencyEditor()}if(dimension==='clinician'){renderClinicianGroupEditor();renderPeople()}cfg.render();if(dimension==='groupUrgency')renderAvailabilityChart();if(dimension==='clinicianGroup')updateClinicianChartFilter();
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
    saveCategories();if(dimension==='urgency'){rebuildCategoryFilter();renderGroupUrgencyEditor()}if(dimension==='clinician'){renderClinicianGroupEditor();renderPeople()}cfg.render();if(dimension==='groupUrgency')renderAvailabilityChart();if(dimension==='clinicianGroup')updateClinicianChartFilter();
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

// Only replace an existing CSV in the selected source folder. Never create files,
// and never write to the separate app-folder settings JSON.
let csvDropBusy=false,csvDragDepth=0;
function csvDropStatus(message,isError=false){
 const status=$('csvDropStatus');if(status){status.textContent=message;status.classList.toggle('csv-drop-error',isError)}
}
function csvDropOverlay(show){
 const overlay=$('csvDropOverlay');if(overlay)overlay.classList.toggle('hidden',!show);
}
function csvUploadModal(show){
 const modal=$('csvUploadModal');if(modal)modal.classList.toggle('hidden',!show);
}
function csvFileDrag(event){
 return [...(event.dataTransfer?.types||[])].includes('Files');
}
function validateReplacementCSV(text){
 const parsed=Papa.parse(text,{header:true,preview:15,skipEmptyLines:'greedy'});
 if(parsed.errors?.length)throw Error('CSV parsing failed: '+parsed.errors[0].message);
 const columns=(parsed.meta.fields||[]).map(x=>String(x||'').replace(/^\uFEFF/,'').trim());
 const required=['Appointment Date','Appointment Time','Booked Time to Slot Time',
  'Slot Type','Availability','User Type of the Session Holder of the Session',
  'Full Name of the Session Holder of the Session','Town of the Location of the Session'];
 const missing=required.filter(x=>!columns.includes(x));
 if(missing.length)throw Error('CSV is incompatible. Missing column(s): '+missing.join(', '));
 if(!parsed.data.length)throw Error('CSV has no appointment rows.');
 return {columns,previewRows:parsed.data.length};
}
function awaitCSVConfirmation(fileName,match){
 return new Promise(resolve=>{
  const modal=$('csvUploadModal'),confirm=$('csvUploadConfirm'),cancel=$('csvUploadCancel');
  if(!modal||!confirm||!cancel){resolve(false);return}
  $('csvUploadMessage').textContent='Replace the existing OneDrive CSV "'+match.name+'" with the dropped file "'+fileName+'"? This will overwrite the existing file. Your people and slot settings will not be changed.';
  csvUploadModal(true);
  const finish=accepted=>{
   csvUploadModal(false);confirm.removeEventListener('click',yes);cancel.removeEventListener('click',no);
   resolve(accepted);
  };
  const yes=()=>finish(true),no=()=>finish(false);
  confirm.addEventListener('click',yes);cancel.addEventListener('click',no);
 });
}
async function replaceExistingCSV(file){
 if(csvDropBusy)return;
 if(!isOwner){csvDropStatus('View-only accounts cannot replace CSV files.',true);return;}if(!activeAccount||!selectedFolder){csvDropStatus('Sign in and select your OneDrive CSV folder first.',true);return}
 if(!file||!/\.csv$/i.test(file.name)){csvDropStatus('Rejected: only CSV files can be uploaded.',true);return}
 if(file.size===0||file.size>200*1024*1024){csvDropStatus('Rejected: CSV must be non-empty and no larger than 200 MB.',true);return}
 csvDropBusy=true;csvDropStatus('Checking '+file.name+' against OneDrive…');
 try{
  const existing=(await children(selectedFolder.id)).filter(item=>item.file&&/\.csv$/i.test(item.name));
  // Case-insensitive name matching is necessary because OneDrive filenames are case-insensitive.
  const matches=existing.filter(item=>item.name.toLocaleLowerCase('en-GB')===file.name.toLocaleLowerCase('en-GB'));
  if(matches.length!==1)throw Error('Rejected: no existing CSV with exactly this filename in the selected OneDrive folder.');
  const match=matches[0];
  const text=await file.text();
  validateReplacementCSV(text);
  // Check the original export's columns as well, so incompatible replacements
  // cannot silently change the source schema.
  const detail=await graph('/me/drive/items/'+encodeURIComponent(match.id));
  if(detail.parentReference?.id&&detail.parentReference.id!==selectedFolder.id)
   throw Error('The existing file is not in the selected OneDrive folder.');
  const originalUrl=detail['@microsoft.graph.downloadUrl'];
  if(!originalUrl)throw Error('Could not verify the original OneDrive CSV.');
  const originalResponse=await fetch(originalUrl,{cache:'no-store'});
  if(!originalResponse.ok)throw Error('Could not check the original CSV.');
  const originalText=await originalResponse.text();
  const originalSchema=validateReplacementCSV(originalText).columns;
  const newSchema=validateReplacementCSV(text).columns;
  if(originalSchema.length!==newSchema.length||originalSchema.some((column,i)=>column!==newSchema[i]))
   throw Error('Rejected: the CSV column headings or their order do not match the existing OneDrive file.');
  csvDropStatus('Validated '+file.name+'. Waiting for confirmation…');
  if(!await awaitCSVConfirmation(file.name,match)){csvDropStatus('Upload cancelled. No files were changed.');return}
  csvDropStatus('Uploading '+file.name+' to OneDrive…');
  const bearer=await token();
  // PUT by existing item ID, never by path: this can only replace the matched file.
  const upload=await fetch(GRAPH+'/me/drive/items/'+encodeURIComponent(match.id)+'/content',{
   method:'PUT',headers:{Authorization:'Bearer '+bearer,'Content-Type':'text/csv'},
   body:file
  });
  if(!upload.ok){
   let message='HTTP '+upload.status;
   try{const response=await upload.json();message=response.error?.message||message}catch{}
   throw Error('OneDrive rejected the replacement: '+message);
  }
  csvDropStatus('Replaced '+match.name+' successfully. Refreshing charts…');
  // This refresh reads the current saved categories and people assignments
  // without changing them; the settings JSON is never overwritten here.
  lastFingerprint='';
  await loadFolder(true);
  csvDropStatus('OneDrive CSV replaced: '+match.name+'. Dashboard refresh completed.');
 }catch(error){
  console.error('CSV replacement error',error);
  csvDropStatus(error?.message||String(error),true);
 }finally{csvDropBusy=false}
}
function initCSVDragDrop(){
 window.addEventListener('dragenter',event=>{
  if(!csvFileDrag(event))return;
  event.preventDefault();csvDragDepth++;
  if(!csvDropBusy)csvDropOverlay(true);
 });
 window.addEventListener('dragover',event=>{
  if(!csvFileDrag(event))return;
  event.preventDefault();event.dataTransfer.dropEffect='copy';
 });
 window.addEventListener('dragleave',event=>{
  if(!csvFileDrag(event))return;
  csvDragDepth=Math.max(0,csvDragDepth-1);
  if(!csvDragDepth)csvDropOverlay(false);
 });
 window.addEventListener('drop',event=>{
  // Internal group/type drags carry text data, not files. Never intercept them.
  if(!csvFileDrag(event))return;
  event.preventDefault();csvDragDepth=0;csvDropOverlay(false);
  const files=[...(event.dataTransfer?.files||[])];
  if(files.length!==1){csvDropStatus('Rejected: drop exactly one CSV file at a time.',true);return}
  void replaceExistingCSV(files[0]);
 });
}
initCSVDragDrop();

let signpostingRecords=[],signpostingFileCount=0;
function isSignpostingFile(name){return /signposting[\s_-]*auto[\s_-]*report/i.test(String(name||''))&&/\.csv$/i.test(name)}
function parseSignpostingDate(raw){
 const value=String(raw||'').trim();
 const m=/^(\d{1,2})[-\/ ]([A-Za-z]{3,9})[-\/ ,]*(\d{4})$/.exec(value);
 if(m){const month=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].slice(0,3).toLowerCase());const day=Number(m[1]),year=Number(m[3]);const dt=new Date(Date.UTC(year,month,day));if(month>=0&&dt.getUTCFullYear()===year&&dt.getUTCMonth()===month&&dt.getUTCDate()===day)return dt.toISOString().slice(0,10)}
 const iso=/^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
 if(iso){const dt=new Date(value+'T00:00:00Z');if(!isNaN(dt)&&dt.toISOString().slice(0,10)===value)return value}
 return '';
}
function parseSignpostingCSV(csv){
 const parsed=Papa.parse(csv,{header:true,skipEmptyLines:false,bom:true});
 const required=['Anonymised Identifier','Age','Gender','Date',"Consultation Details' Time",'Code Term'];
 const headers=(parsed.meta.fields||[]).map(h=>h.trim());
 if(!required.every(col=>headers.includes(col)))throw Error('Signposting CSV is missing expected columns: '+required.filter(col=>!headers.includes(col)).join(', '));
 const previous=Object.fromEntries(required.map(k=>[k,'']));
 const results=[];
 for(const raw of parsed.data){
  // Fill down EVERY column independently, including identifiers and demographics.
  // Completely empty rows are ignored so they cannot become duplicate events.
  const allFields=Object.keys(raw).filter(k=>k!=='__parsed_extra');
  if(!allFields.some(k=>String(raw[k]??'').trim()))continue;
  const row={};
  for(const key of allFields){
   const value=String(raw[key]??'').trim();
   if(value)previous[key]=value;
   row[key]=value||previous[key]||'';
  }
  const date=parseSignpostingDate(row.Date),destination=String(row['Code Term']||'').trim();
  if(date&&destination)results.push({date,destination,time:String(row["Consultation Details' Time"]||'').trim()});
 }
 return results;
}
function renderSignpostingChart(){
 const target=$('signpostingChart');if(!target||$('signpostingPage').classList.contains('hidden'))return;
 const status=$('signpostingStatus');
 if(!signpostingFileCount){Plotly.purge(target);status.textContent='No Signposting Auto Report CSV found in the selected OneDrive folder.';return}
 if(!signpostingRecords.length){Plotly.purge(target);status.textContent='Signposting file loaded, but no valid dated signposting records were found.';return}
 const counts=new Map(),dates=new Set(),destinations=new Set();
 for(const record of signpostingRecords){dates.add(record.date);destinations.add(record.destination);const key=JSON.stringify([record.date,record.destination]);counts.set(key,(counts.get(key)||0)+1)}
 const sortedDates=[...dates].sort(),sortedDestinations=[...destinations].sort((a,b)=>a.localeCompare(b));
 const traces=sortedDestinations.map((destination,i)=>({
  type:'scatter',mode:'lines',name:destination,
  x:sortedDates,y:sortedDates.map(date=>counts.get(JSON.stringify([date,destination]))||0),
  line:{width:2.4,shape:'linear'},
  hovertemplate:'<b>%{x|%d %b %Y}</b><br>'+destination.replace(/&/g,'&amp;').replace(/</g,'&lt;')+': %{y:,}<extra></extra>',
  connectgaps:false
 }));
 Plotly.react(target,traces,{
  autosize:true,height:535,margin:{l:62,r:24,t:28,b:65},
  paper_bgcolor:'#ffffff',plot_bgcolor:'#ffffff',
  xaxis:{type:'date',title:'Date',tickformat:'%b %Y',gridcolor:'#edf1f5'},
  yaxis:{title:'Signposting events',rangemode:'tozero',gridcolor:'#edf1f5'},
  legend:{orientation:'h',y:1.18,x:0},hovermode:'x unified',
  hoverlabel:{bgcolor:'#ffffff',bordercolor:'#dce5ed',font:{family:'Arial',size:12,color:'#182b3b'},align:'left'}
 },{responsive:true,displaylogo:false,displayModeBar:false});
 status.textContent=signpostingRecords.length.toLocaleString('en-GB')+' signposting records · '+sortedDestinations.length+' destinations · '+signpostingFileCount+' CSV file'+(signpostingFileCount===1?'':'s')+'. Empty cells filled from preceding rows.';
}
async function loadFolder(force=true){if(!selectedFolder||busy)return;busy=true;setButtons();showRefreshFeedback('Refreshing data…',true);setSync('Checking OneDrive CSV files…');try{
 const items=(await children(selectedFolder.id)).filter(f=>f.file&&/\.csv$/i.test(f.name)).sort((a,b)=>a.name.localeCompare(b.name));if(!items.length)throw Error('No CSV files found directly inside this folder.');
 const fingerprint=JSON.stringify(items.map(f=>[f.id,f.lastModifiedDateTime,f.size]));if(!force&&fingerprint===lastFingerprint){showRefreshFeedback('Already up to date');setSync(`No changes detected · ${items.length} CSV files · Checked ${new Date().toLocaleTimeString('en-GB')}`);return}
 const aggregate=new Map(),availabilityAggregate=new Map(),bookingTimeAggregate=new Map(),peopleAggregate=new Map();let valid=0,skipped=0;const newSignposting=[];let newSignpostingFiles=0,newSlotFiles=0;for(let i=0;i<items.length;i++){const f=items[i];showRefreshFeedback(`Loading file ${i+1} of ${items.length}…`,true);setSync(`Loading ${i+1} of ${items.length}: ${f.name}`);const detail=await graph(sourceItemUrl(f.id));const url=detail['@microsoft.graph.downloadUrl'];if(!url)throw Error('OneDrive did not provide a download URL for '+f.name);const response=await fetch(url,{cache:'no-store'});if(!response.ok)throw Error(`Could not download ${f.name}: HTTP ${response.status}`);const csvText=await response.text();
 if(isSignpostingFile(f.name)){newSignposting.push(...parseSignpostingCSV(csvText));newSignpostingFiles++;continue}
 if(!/slot[\s_-]*report/i.test(f.name))continue;
 newSlotFiles++;
 const stats=aggregateCSV(csvText,aggregate,availabilityAggregate,bookingTimeAggregate,peopleAggregate);valid+=stats.valid;skipped+=stats.skipped}
 signpostingRecords=newSignposting;signpostingFileCount=newSignpostingFiles;
 // Each dated signposting event is also a booked appointment for the main
 // demand/availability calculations. Its destination is its Slot Type.
 // Keep the dedicated Signposting page sourced from the original event list.
 // No clinician/person is invented: the report has no session holder field.
 // The event date is used as its booking AND appointment date.
 const signpostingClinician='Signposting (no clinician recorded)';
 for(const event of newSignposting){
  const slot=event.destination,date=event.date;
  const availabilityKey=JSON.stringify([date,slot,'Booked',signpostingClinician,'']);
  availabilityAggregate.set(availabilityKey,(availabilityAggregate.get(availabilityKey)||0)+1);
  const bookedWeek=mondayOf(date);
  if(bookedWeek>='2025-07-07'){
   const key=JSON.stringify([date,slot,signpostingClinician,'']);
   aggregate.set(key,(aggregate.get(key)||0)+1);
   const tm=/^(\\d{1,2}):(\\d{2})/.exec(event.time||'');
   if(tm&&Number(tm[1])<24&&Number(tm[2])<60){
    const minute=Number(tm[1])*60+Number(tm[2]);
    const timeKey=JSON.stringify([date,minute,slot,signpostingClinician,'']);
    bookingTimeAggregate.set(timeKey,(bookingTimeAggregate.get(timeKey)||0)+1);
   }
  }
 }
 const peopleByName=new Map();
 for(const [key,count] of peopleAggregate){
  const [name,type]=JSON.parse(key);
  const id=name.toLocaleLowerCase('en-GB');
  const entry=peopleByName.get(id)||{name,types:new Map()};
  entry.types.set(type,(entry.types.get(type)||0)+count);
  peopleByName.set(id,entry);
 }
 peopleRecords=[...peopleByName.values()].map(({name,types})=>({name,type:[...types].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]))[0][0]})).sort((a,b)=>a.type.localeCompare(b.type)||a.name.localeCompare(b.name));
 rebuildPeopleTypeLookup();
 renderPeople();renderSlotTable();
 bookingTimeRecords=[...bookingTimeAggregate].map(([key,count])=>{const [bookedDay,minute,slot,clinician,person]=JSON.parse(key);return {bookedDay,minute,slot,clinician,person,count}});
 availabilityRecords=[...availabilityAggregate].map(([key,count])=>{const [date,slot,availability,clinician,person]=JSON.parse(key);return {date,slot,availability,clinician,person,count}});
 records=[...aggregate].map(([key,count])=>{const [bookedDay,slot,clinician,person]=JSON.parse(key);const bookedWeek=mondayOf(bookedDay);return {week:bookedWeek,bookedWeek,bookedDay,slot,clinician,person,status:'Booked',count}});weeks=[...new Set(records.map(r=>r.bookedWeek))].sort();types=[...new Set([...records.map(r=>r.slot),...availabilityRecords.map(r=>r.slot)])].sort();statuses=['Booked'];clinicianTypes=[...new Set([...records.map(r=>r.clinician),...availabilityRecords.map(r=>r.clinician)])].filter(Boolean).sort();
 // Assign obvious names on first encounter; retain any existing manual choices.
 // Do not guess assignments: unknown slot types stay Uncategorised until manually assigned.

 saveCategories(false);renderSlotTable();lastFingerprint=fingerprint;
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
  ['booking times chart',()=>renderBookingTimesChart()],
  ['signposting chart',()=>renderSignpostingChart()]
 ]){
  try{draw()}catch(error){drawingErrors.push(label+': '+(error?.message||String(error)));console.error('Dashboard rendering error in '+label,error)}
 }
 showRefreshFeedback(drawingErrors.length?'Data loaded; chart issue: '+drawingErrors.join('; '):(!isOwner?'Data refreshed · shared slot assignments: '+Object.keys(assignments).length:'Data refreshed successfully'));
 scheduleChartResize();setSync(`${newSlotFiles} Slot Report CSV files · ${newSignpostingFiles} Signposting CSV files · ${valid.toLocaleString('en-GB')} booked Slot Report rows + ${newSignposting.length.toLocaleString('en-GB')} signposting events · ${skipped.toLocaleString('en-GB')} skipped · ${Object.keys(assignments).length} assigned slot types · Updated ${new Date().toLocaleString('en-GB')}`)
 }catch(e){
 const message=e?.message||String(e);
 console.error('OneDrive refresh error',e);
 showRefreshFeedback('Refresh failed: '+message);
 setSync('Refresh failed: '+message+' · You remain signed in; the existing charts have been preserved where possible.');
 }finally{busy=false;setButtons()}}
$('signIn').onclick=async()=>{if(folderAccessDenied&&auth&&activeAccount){
 try{folderAccessDenied=false;folderAccessVerified=false;authenticatedOnce=false;$('dashboardShell').dataset.signedIn='false';await auth.logoutRedirect({account:activeAccount,postLogoutRedirectUri:REDIRECT_URI});}catch(e){setStatus('Account switch failed: '+e.message)}return;
 }if(authenticatedOnce||activeAccount)return;try{if(!auth){await initAuth();if(!auth)throw Error('Microsoft sign-in is still unavailable. Refresh the page (Ctrl+F5).')}await auth.loginRedirect({scopes:SCOPES})}catch(e){setStatus('Sign-in failed: '+e.message)}};
$('signOut').onclick=async()=>{if(!activeAccount)return;try{authenticatedOnce=false;$('dashboardShell').dataset.signedIn='false';await auth.logoutRedirect({account:activeAccount,postLogoutRedirectUri:REDIRECT_URI})}catch(e){setStatus('Sign-out failed: '+e.message)}};
$('chooseFolder').onclick=()=>{if(isOwner)showFolderPicker()};
$('sourceSetup').onclick=()=>{if(isOwner)showFolderPicker()};
$('cancelFolder').onclick=()=>$('folderPicker').classList.add('hidden');
$('pickerUp').onclick=async()=>{if(!folderView.trail.length)return;const parent=folderView.trail.at(-1);folderView={...parent,trail:folderView.trail.slice(0,-1)};await showFolderContents()};
$('selectFolder').onclick=async()=>{if(!isOwner)return;selectedFolder={id:folderView.id,name:[...folderView.trail.map(x=>x.name),folderView.name].join(' / ')};localStorage.setItem(FOLDER_KEY,JSON.stringify(selectedFolder));$('folderStatus').textContent='Selected folder: '+selectedFolder.name;$('folderPicker').classList.add('hidden');$('onedrivePanel').classList.add('hidden');$('sourceSetup').classList.add('hidden');cloudFolder=selectedFolder;await saveCloudCategories();lastFingerprint='';await loadFolder(true)};
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
 if(names.length===lastClinicianFilterNames.length&&names.every((n,i)=>n===lastClinicianFilterNames[i])){root.querySelectorAll('input[data-filter-name]').forEach(input=>input.nextElementSibling?.style.setProperty('--switch-on',clinicianGroupColour(input.dataset.filterName)));return;}
 lastClinicianFilterNames=names;
 root.replaceChildren();
 if(!names.length){const note=document.createElement('span');note.className='hint';note.textContent='Create clinician groups in Manage categories.';root.append(note);return;}
 for(const name of names){
  const row=document.createElement('label');row.className='filter-row';
  const switchEl=document.createElement('span');switchEl.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.checked=!disabledClinicianTypes.has(name);input.dataset.filterName=name;input.setAttribute('aria-label','Include '+name);
  const track=document.createElement('span');track.className='switch-track';track.style.setProperty('--switch-on',clinicianGroupColour(name));
  switchEl.append(input,track);
  const title=document.createElement('span');title.textContent=name;
  row.append(switchEl,title);root.append(row);
  input.addEventListener('change',()=>{if(input.checked)disabledClinicianTypes.delete(name);else disabledClinicianTypes.add(name);scheduleDemandRefresh()});
 }
}
let clinicianMatchCache=new Map();
function clinicianMatches(clinician,person){
 const key=clinician+'\\u0000'+person;
 if(clinicianMatchCache.has(key))return clinicianMatchCache.get(key);
 const detailed=effectiveClinicianType(clinician,person);
 const group=clinicianGroupAssignments[detailed];
 // Only session-holder user types mapped through a detailed category to an enabled group count.
 const matches=clinicianGroupNames.includes(group)&&!disabledClinicianTypes.has(group);
 clinicianMatchCache.set(key,matches);return matches;
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
  const track=document.createElement('span');track.className='switch-track';track.style.setProperty('--switch-on','#596675');
  switchEl.append(input,track);
  const title=document.createElement('span');title.textContent=name;
  row.append(switchEl,title);root.append(row);
  input.addEventListener('change',()=>{if(input.checked)disabledConsultModes.delete(name);else disabledConsultModes.add(name);scheduleDemandRefresh()});
 }
}
function consultMatches(slot){
 const mode=consultAssignments[slot];
 return consultCategoryNames.includes(mode)&&!disabledConsultModes.has(mode);
}
const bankHolidayWeeks=new Set(holidayDates.map(mondayOf));let excludeBankHolidays=true;
document.getElementById('excludeBH').onchange=e=>{excludeBankHolidays=!e.target.checked;scheduleDemandRefresh()};
function isoWeek(weekStart){const d=new Date(weekStart+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+3);const year=d.getUTCFullYear();const jan4=new Date(Date.UTC(year,0,4,12));jan4.setUTCDate(jan4.getUTCDate()-((jan4.getUTCDay()+6)%7));const week=1+Math.round((d-jan4)/604800000);return {year,week}}
const colours=['#1475b8','#e57a28','#159a78','#9053b4','#c64c69','#65748b'];
let current=[];let continuousMode=false;let showCurrentWeek=true;let monthlyMode=false;
let shadeSeasons=document.getElementById('shadeSeasons').checked;
document.getElementById('shadeSeasons').onchange=e=>{shadeSeasons=e.target.checked;scheduleDemandRefresh()};
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
document.getElementById('currentWeekToggle').onchange=e=>{showCurrentWeek=e.target.checked;scheduleDemandRefresh()};
document.getElementById('displayMode').onchange=e=>{continuousMode=e.target.checked;document.getElementById('timelineValue').textContent=continuousMode?'Continuous':'Superimposed';scheduleDemandRefresh()};
function periodKey(r){return monthlyMode?r.bookedDay?.slice(0,7):r.bookedWeek}
let demandFilterCache=null;
function includedRecord(r){
 if(demandFilterCache&&demandFilterCache.has(r))return demandFilterCache.get(r);
 const result=r.status==='Booked'&&clinicianMatches(r.clinician,r.person)&&consultMatches(r.slot)&&(monthlyMode||!excludeBankHolidays||!bankHolidayWeeks.has(r.bookedWeek));
 if(demandFilterCache)demandFilterCache.set(r,result);
 return result;
}
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
let demandRefreshQueued=false;
function scheduleDemandRefresh(){
 if(demandRefreshQueued)return;
 demandRefreshQueued=true;
 requestAnimationFrame(()=>{demandRefreshQueued=false;refresh()});
}
function refresh(){
 updateClinicianChartFilter();updateConsultChartFilter();
 demandFilterCache=new WeakMap();
 clinicianMatchCache=new Map();
const activeWeekSet=new Set();
for(const r of records){if(includedRecord(r)){const key=periodKey(r);if(key)activeWeekSet.add(key)}}
const activeWeeks=[...activeWeekSet].sort();
const selectedWeeks=activeWeeks.filter(w=>w>=(monthlyMode?'2025-07':'2025-07-07')&&w<=(monthlyMode?thisMonth:thisWeek)&&(showCurrentWeek||!periodIsCurrent(w)));
const sums=new Map(selectedWeeks.map(w=>[w,0]));
const demandCategoryTotals=Object.fromEntries(['urgent','routine','followup'].map(c=>[c,new Map(selectedWeeks.map(w=>[w,0]))]));
for(const r of records){
 const key=periodKey(r);
 if(!sums.has(key)||!includedRecord(r))continue;
 sums.set(key,sums.get(key)+r.count);
 for(const category of ['urgent','routine','followup']){
  if(weeklyUrgencyMatches(r.slot,category)){
   const totals=demandCategoryTotals[category];totals.set(key,totals.get(key)+r.count);
  }
 }
}
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
 const totals=demandCategoryTotals[category];
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
 Plotly.react(target,traces,{autosize:true,width:chartWidth,margin:{l:58,r:12,t:32,b:65},paper_bgcolor:'white',plot_bgcolor:'white',shapes:shadeSeasons?seasonBackgroundShapes(selectedWeeks):[],legend:{orientation:'h',y:1.17,x:0},xaxis:continuousMode?{type:'date',title:monthlyMode?'Booking month':'Booking date',tickformat:'%b %Y',gridcolor:'#edf1f5',...(continuousRange?{range:continuousRange}: {})}:axis,yaxis:{title:'Appointments',range:[0,sharedMax],dtick:sharedMax/5,gridcolor:'#edf1f5'},hovermode:'closest',annotations:!hasAppointments&&category==='followup'?[{text:'No follow-up appointments found. In Manage Categories, assign slot types to Follow up and click Save.',xref:'paper',yref:'paper',x:.5,y:.55,showarrow:false,align:'center',font:{size:13,color:'#286cb8'}}]:[]},{responsive:true,displaylogo:false,displayModeBar:false});
}
let maximum=0;
let followupMaximum=0;

for(const category of ['urgent','routine','followup']){

 const totals=demandCategoryTotals[category];
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
const STORAGE_KEY='appointment-dashboard-categories-v1';let categoryNames=['Urgent','Routine','Follow up'],assignments={};try{const saved=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');categoryNames=Array.isArray(saved.categories)&&saved.categories.length?saved.categories.filter(x=>typeof x==='string'&&x.trim()&&x!=='Uncategorised'):['Urgent','Routine','Follow up'];assignments=saved.assignments&&typeof saved.assignments==='object'?saved.assignments:{};}catch(e){}let clinicianCategoryNames=[],clinicianAssignments={};try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.clinicianCategories))clinicianCategoryNames=validCategoryNames(local.clinicianCategories);if(local.clinicianAssignments)clinicianAssignments=validAssignments(local.clinicianAssignments,clinicianCategoryNames)}catch(e){}let consultCategoryNames=[],consultAssignments={};try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.consultCategories))consultCategoryNames=validCategoryNames(local.consultCategories);if(local.consultAssignments)consultAssignments=validAssignments(local.consultAssignments,consultCategoryNames)}catch(e){}try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.groupUrgencyCategories))groupUrgencyNames=validCategoryNames(local.groupUrgencyCategories);if(local.groupUrgencyAssignments)groupUrgencyAssignments=validAssignments(local.groupUrgencyAssignments,groupUrgencyNames);if(local.urgencyGroupColours&&typeof local.urgencyGroupColours==='object')urgencyGroupColours=local.urgencyGroupColours;if(local.urgencyTypeColours&&typeof local.urgencyTypeColours==='object')urgencyTypeColours=local.urgencyTypeColours}catch(e){}
try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(Array.isArray(local.clinicianGroupCategories))clinicianGroupNames=validCategoryNames(local.clinicianGroupCategories);if(local.clinicianGroupAssignments)clinicianGroupAssignments=validAssignments(local.clinicianGroupAssignments,clinicianGroupNames)}catch(e){}
try{const local=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(local.peopleLocum&&typeof local.peopleLocum==='object')peopleLocum=local.peopleLocum;if(local.peopleClinicianType&&typeof local.peopleClinicianType==='object')peopleClinicianType=local.peopleClinicianType;if(local.clinicianTypeColours&&typeof local.clinicianTypeColours==='object')clinicianTypeColours=local.clinicianTypeColours;if(local.urgencyTypeColours&&typeof local.urgencyTypeColours==='object')urgencyTypeColours=local.urgencyTypeColours;if(local.clinicianGroupColours&&typeof local.clinicianGroupColours==='object')clinicianGroupColours=local.clinicianGroupColours}catch(e){}
function ensureFollowUpCategory(){if(!categoryNames.some(c=>c.toLowerCase().replace(/[\s-]+/g,'')==='followup'))categoryNames.push('Follow up')}ensureFollowUpCategory();function guessCategory(t){return /RED|YELLOW|AMBER|URGENT|SAME DAY|ACUTE/i.test(t)?'Urgent':/GREEN|FOLLOW UP|ROUTINE/i.test(t)?'Routine':null}function saveCategories(sync=true){lastConsultFilterNames=[];try{localStorage.setItem(STORAGE_KEY,JSON.stringify({categories:categoryNames,assignments,clinicianCategories:clinicianCategoryNames,clinicianAssignments,consultCategories:consultCategoryNames,consultAssignments,groupUrgencyCategories:groupUrgencyNames,groupUrgencyAssignments,urgencyGroupColours,clinicianGroupCategories:clinicianGroupNames,clinicianGroupAssignments,peopleLocum,peopleClinicianType,clinicianGroupColours,clinicianTypeColours,urgencyTypeColours}));}catch(e){alert('Could not save settings in this browser. Please try again.')}if(sync)markCategoryDirty()}function allCategories(){return [...categoryNames,'Uncategorised']}function rebuildCategoryFilter(){const old=new Set([...slots.selectedOptions].map(o=>o.value));slots.replaceChildren();for(const c of allCategories()){const o=document.createElement('option');o.value=c;o.textContent=c;o.selected=old.has(c);slots.append(o)}}function renderEditor(){renderDropEditor('urgency');updateAvailabilityUrgencyFilter()}
function renderGroupUrgencyEditor(){renderDropEditor('groupUrgency');updateAvailabilityUrgencyFilter()}document.getElementById('addCategory').onclick=()=>{const el=document.getElementById('newCategory');const c=el.value.trim();if(!c)return;if(c.toLowerCase()==='uncategorised'||categoryNames.some(x=>x.toLowerCase()===c.toLowerCase())){alert('This category name already exists or is reserved.');return}categoryNames.push(c);el.value='';saveCategories();rebuildCategoryFilter();renderEditor();renderGroupUrgencyEditor()};document.getElementById('newCategory').addEventListener('keydown',e=>{if(e.key==='Enter')document.getElementById('addCategory').click()});document.getElementById('slotSearch').addEventListener('input',renderEditor);function showPage(which){
 if(!isOwner&&(which==='categories'||which==='people'))return;
 for(const [page,nav,name] of [['chartPage','navChart','chart'],['categoryPage','navCategories','categories'],['availabilityPage','navAvailability','availability'],['routineAvailabilityPage','navRoutineAvailability','routineAvailability'],['bookingTimesPage','navBookingTimes','bookingTimes'],['signpostingPage','navSignposting','signposting'],['peoplePage','navPeople','people']]){
  $(page).classList.toggle('hidden',which!==name);$(nav).classList.toggle('active',which===name);
 }
 if(which==='chart')setTimeout(()=>{refresh();scheduleChartResize()},0);
 else if(which==='availability')setTimeout(renderAvailabilityChart,0);
 else if(which==='routineAvailability')setTimeout(renderRoutineAvailabilityChart,0);
 else if(which==='bookingTimes')setTimeout(renderBookingTimesChart,0);
 else if(which==='signposting')setTimeout(renderSignpostingChart,0);
 else if(which==='people'){renderPeopleGroups();renderPeopleGroupFilter();renderPeople();}
 else {renderSlotGroupBoxes();renderSlotTable();renderEditor();renderGroupUrgencyEditor();renderClinicianEditor();renderClinicianGroupEditor();renderConsultEditor()}
}document.getElementById('navChart').onclick=()=>showPage('chart');document.getElementById('navCategories').onclick=()=>showPage('categories');document.getElementById('navAvailability').onclick=()=>showPage('availability');document.getElementById('navRoutineAvailability').onclick=()=>showPage('routineAvailability');document.getElementById('navBookingTimes').onclick=()=>showPage('bookingTimes');document.getElementById('navSignposting').onclick=()=>showPage('signposting');document.getElementById('navPeople').onclick=()=>showPage('people');$('slotTableSearch').addEventListener('input',renderSlotTable);$('slotGroupFilter').addEventListener('change',renderSlotTable);$('slotUnassignedOnly').addEventListener('change',()=>{$('slotGroupFilter').disabled=$('slotUnassignedOnly').checked;renderSlotTable()});$('peopleSearch').addEventListener('input',renderPeople);$('peopleGroupFilter').addEventListener('change',renderPeople);$('peopleUnassignedOnly').addEventListener('change',()=>{const only=$('peopleUnassignedOnly').checked;$('peopleGroupFilter').disabled=only;renderPeople()});
function renderClinicianEditor(){renderDropEditor('clinician');updateClinicianChartFilter()}
function renderClinicianGroupEditor(){renderDropEditor('clinicianGroup');if(!$('peoplePage')?.classList.contains('hidden'))renderPeopleGroups()}
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
rebuildCategoryFilter();renderEditor();renderGroupUrgencyEditor();renderClinicianEditor();renderClinicianGroupEditor();renderConsultEditor();renderSlotGroupBoxes();
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
const urgentClinicianDisabled=new Set();
const urgentConsultDisabled=new Set();
// Eligibility must be independent of currently selected switches: otherwise
// disabling one dimension can cascade and lock every other dimension off.
function availabilityRecordsInRange(page){
 const urgent=page==='urgent';
 const from=urgent?-urgentWeeksBack:routineStartWeek;
 const to=urgent?urgentWeeksForward:routineEndWeek;
 const start=new Date(thisWeek+'T12:00:00Z');start.setUTCDate(start.getUTCDate()+from*7);
 const finish=new Date(thisWeek+'T12:00:00Z');finish.setUTCDate(finish.getUTCDate()+(to+1)*7);
 const first=start.toISOString().slice(0,10),last=finish.toISOString().slice(0,10);
 return availabilityRecords.filter(r=>{
  if(!(r.count>0)||r.date<first||r.date>=last)return false;
  const category=String(assignments[r.slot]||'').trim();
  if(String(groupUrgencyAssignments[category]||'').trim().toLowerCase()!==page)return false;
  if(urgent){const day=new Date(r.date+'T12:00:00Z').getUTCDay();if(day===0||day===6)return false;}
  return true;
 });
}
function availabilityOptionHasData(page,kind,name){
 return availabilityRecordsInRange(page).some(r=>{
  const category=String(assignments[r.slot]||'').trim();
  const clinician=clinicianGroupAssignments[effectiveClinicianType(r.clinician,r.person)];
  const consult=consultAssignments[r.slot];
  return kind==='category'?category===name:kind==='clinician'?clinician===name:consult===name;
 });
}
function availabilityHasVisibleRecords(page){
 return availabilityRecordsInRange(page).some(r=>{
  const category=String(assignments[r.slot]||'').trim();
  const clinician=clinicianGroupAssignments[effectiveClinicianType(r.clinician,r.person)];
  const consult=consultAssignments[r.slot];
  return page==='urgent'?
   !urgentCategoryDisabled.has(category)&&urgentCategoryNames().includes(category)&&
   clinicianGroupNames.includes(clinician)&&!urgentClinicianDisabled.has(clinician)&&
   consultCategoryNames.includes(consult)&&!urgentConsultDisabled.has(consult):
   clinicianGroupNames.includes(clinician)&&!routineDisabledClinicianGroups.has(clinician)&&
   consultCategoryNames.includes(consult)&&!routineDisabledConsultModes.has(consult);
 });
}
function changeAvailabilityFilter(page,disabled,name,checked,checkbox){
 const previouslyOff=disabled.has(name);
 if(checked)disabled.delete(name);else disabled.add(name);
 // Reject any change that would remove every appointment from the chart.
 // If source data itself has no matches, do not block filter changes.
 if(availabilityRecordsInRange(page).length&&!availabilityHasVisibleRecords(page)){
  if(previouslyOff)disabled.add(name);else disabled.delete(name);
  checkbox.checked=!previouslyOff;
  return;
 }
 (page==='urgent'?renderAvailabilityChart:renderRoutineAvailabilityChart)();
}
function refreshAvailabilityOptionStates(page){
 const configs=page==='urgent'?
  [['urgentCategoryToggles','category',urgentCategoryDisabled],['urgentClinicianToggles','clinician',urgentClinicianDisabled],['urgentConsultToggles','consult',urgentConsultDisabled]]:
  [['routineClinicianFilters','clinician',routineDisabledClinicianGroups],['routineConsultFilters','consult',routineDisabledConsultModes]];
 for(const [id,kind,disabled] of configs){
  const root=$(id);if(!root)continue;
  for(const input of root.querySelectorAll('input[data-filter-name]')){
   const name=input.dataset.filterName;
   const possible=availabilityOptionHasData(page,kind,name);
   // Restore options disabled only by an earlier buggy eligibility pass.
   if(!possible)disabled.add(name);
   input.disabled=!possible;
   input.checked=possible&&!disabled.has(name);
   input.closest('.filter-row')?.classList.toggle('availability-option-unavailable',!possible);
   input.title=possible?'':'No source appointments for this option in the selected date range';
  }
 }
 // Recover from stale all-off state caused by earlier versions of this logic.
 if(availabilityRecordsInRange(page).length&&!availabilityHasVisibleRecords(page)){
  for(const [id,kind,disabled] of configs){
   const root=$(id);if(!root)continue;
   for(const input of root.querySelectorAll('input[data-filter-name]')){
    if(!input.disabled){disabled.delete(input.dataset.filterName);input.checked=true;}
   }
  }
 }
}
const urgentDisabledClinicianTypes=new Set(),routineDisabledClinicianTypes=new Set();
function selectedClinicianTypes(page){
 const disabled=page==='urgent'?urgentClinicianDisabled:routineDisabledClinicianGroups;
 return clinicianCategoryNames.filter(type=>{
  const group=clinicianGroupAssignments[type];
  return group&&clinicianGroupNames.includes(group)&&!disabled.has(group);
 });
}
function renderClinicianTypeFilters(page){
 const root=$(page+'ClinicianTypeFilters');if(!root)return;
 const disabled=page==='urgent'?urgentDisabledClinicianTypes:routineDisabledClinicianTypes;
 const names=selectedClinicianTypes(page);
 const existing=[...root.querySelectorAll('input[data-filter-name]')].map(input=>input.dataset.filterName);
 if(existing.length!==names.length||names.some((name,i)=>name!==existing[i])){
  root.replaceChildren();
  if(!names.length){const note=document.createElement('span');note.className='hint';note.textContent='No clinician types in selected groups';root.append(note);}
  for(const name of names){
   const row=document.createElement('label');row.className='filter-row';
   const sw=document.createElement('span');sw.className='switch';
   const input=document.createElement('input');input.type='checkbox';input.dataset.filterName=name;input.setAttribute('role','switch');input.checked=!disabled.has(name);
   const track=document.createElement('span');track.className='switch-track';
   const label=document.createElement('span');label.textContent=name;
   sw.append(input,track);row.append(sw,label);root.append(row);
   input.addEventListener('change',()=>{input.checked?disabled.delete(name):disabled.add(name);(page==='urgent'?renderAvailabilityChart:renderRoutineAvailabilityChart)();});
  }
 }
 updateCollapsedFilterSummaries();
}
function updateCollapsedFilterSummaries(){
 document.querySelectorAll('.collapsed-selection-summary[data-summary-for]').forEach(summary=>{
  const root=$(summary.dataset.summaryFor);if(!root)return;
  const selected=[...root.querySelectorAll('input[data-filter-name]')].filter(input=>input.checked&&!input.disabled).map(input=>input.dataset.filterName);
  summary.textContent=selected.length?selected.join(', '):'None selected';
  const expanded=summary.parentElement.querySelector('.collapse-filter-button')?.getAttribute('aria-expanded')==='true';
  summary.hidden=expanded;root.hidden=!expanded;
 });
}
document.querySelectorAll('.collapse-filter-button').forEach(button=>{
 button.addEventListener('click',()=>{
  const expanded=button.getAttribute('aria-expanded')==='true';
  button.setAttribute('aria-expanded',String(!expanded));
  button.querySelector('.collapse-sign').textContent=expanded?'+':'−';
  updateCollapsedFilterSummaries();
 });
});
function simpleAvailabilityHover(heading,labels,values,total){
 return '<b>'+heading+'</b><br>'+labels.map((label,i)=>label+': '+Number(values[i]||0).toLocaleString('en-GB')).join('<br>')+'<br><b>Total: '+total.toLocaleString('en-GB')+'</b>';
}
function renderUrgentSideFilters(){
 const build=(id,names,disabled,colour)=>{
  const root=$(id);if(!root)return;
  const previous=[...root.querySelectorAll('input[data-filter-name]')].map(el=>el.dataset.filterName);
  if(previous.length===names.length&&names.every((n,i)=>n===previous[i])){root.querySelectorAll('input[data-filter-name]').forEach(input=>input.nextElementSibling?.style.setProperty('--switch-on',id==='urgentClinicianToggles'?clinicianGroupColour(input.dataset.filterName):'#a1c8e8'));return;}
  root.replaceChildren();
  for(const name of names){
   const row=document.createElement('label');row.className='filter-row';
   const sw=document.createElement('span');sw.className='switch';
   const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.dataset.filterName=name;input.checked=!disabled.has(name);
   input.setAttribute('aria-label','Include '+name);
   const track=document.createElement('span');track.className='switch-track';track.style.setProperty('--switch-on',id==='urgentClinicianToggles'?clinicianGroupColour(name):colour);
   const caption=document.createElement('span');caption.textContent=name;
   sw.append(input,track);row.append(sw,caption);root.append(row);
   input.addEventListener('change',()=>{
    changeAvailabilityFilter('urgent',disabled,name,input.checked,input);
   });
  }
 };
 build('urgentClinicianToggles',clinicianGroupNames,urgentClinicianDisabled,'#a6d6b2');
 build('urgentConsultToggles',consultCategoryNames,urgentConsultDisabled,'#a1c8e8');
}
function urgentClinicianMatches(clinician,person){
 const detailed=effectiveClinicianType(clinician,person);const group=clinicianGroupAssignments[detailed];
 return clinicianGroupNames.includes(group)&&!urgentClinicianDisabled.has(group)&&!urgentDisabledClinicianTypes.has(detailed);
}
function urgentConsultMatches(slot){
 const mode=consultAssignments[slot];
 return consultCategoryNames.includes(mode)&&!urgentConsultDisabled.has(mode);
}
const urgentCategoryDisabled=new Set();
let urgentAvailabilityGranularity='day';
function setUrgentAvailabilityGranularity(mode){
 urgentAvailabilityGranularity=mode==='week'?'week':'day';
 const switchInput=$('urgentDaysWeeksSwitch');if(switchInput)switchInput.checked=urgentAvailabilityGranularity==='week';
 renderAvailabilityChart();
}
$('urgentDaysWeeksSwitch')?.addEventListener('change',event=>setUrgentAvailabilityGranularity(event.target.checked?'week':'day'));

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
  if(rail){rail.style.setProperty('--range-start',String((start+3)/7));rail.style.setProperty('--range-span',String((end-start)/7))}
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
 if(current.length===names.length&&current.every((x,i)=>x===names[i])){root.querySelectorAll('input[data-filter-name]').forEach(input=>input.nextElementSibling?.style.setProperty('--switch-on',urgentColour(input.dataset.filterName)));return;}
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
  input.addEventListener('change',()=>{
   changeAvailabilityFilter('urgent',urgentCategoryDisabled,name,input.checked,input);
  });
 }
}
function urgentColour(name){return urgencyTypeColour(name)}
function urgentGradientStops(hex){
 const rgb=[1,3,5].map(i=>parseInt(hex.slice(i,i+2),16));
 const mix=(other,t)=>'#'+rgb.map((v,i)=>Math.round(v*(1-t)+other[i]*t).toString(16).padStart(2,'0')).join('');
 return [['0%',mix([255,255,255],0.37)],['17%',mix([255,255,255],0.12)],['54%',hex],['100%',mix([0,0,0],0.12)]];
}
// SVG gradient is a visual finish only. It does not alter Plotly data, bar
// positions, totals, tooltip targets, Today marker or historical shading.
// Shared flat visual treatment; saved type colours remain the source of truth.
function applyFlatAvailabilityBars(target){
 // Plotly SVG paths need an explicit fill: removing their fill renders them black.
 // Read each bar trace's actual marker colour, including custom colours from Slot settings.
 const barTraces=(target.data||[]).filter(trace=>trace.type==='bar');
 const rendered=target.querySelectorAll('.barlayer .trace.bars');
 rendered.forEach((traceElement,i)=>{
  const marker=barTraces[i]?.marker||{};
  const baseColour=typeof marker.color==='string'?marker.color:'#c2cdd5';
  traceElement.querySelectorAll('.point path').forEach((path,j)=>{
   let height=0;
   try{height=path.getBBox().height}catch(_){}
   if(height<0.6){
    path.style.setProperty('fill','none','important');
    path.style.setProperty('stroke','none','important');
    return;
   }
   const colour=Array.isArray(marker.color)?marker.color[j]||baseColour:baseColour;
   path.style.setProperty('fill',colour,'important');
   path.style.setProperty('fill-opacity','1','important');
   path.style.setProperty('stroke','rgba(255,255,255,0.85)','important');
   path.style.setProperty('stroke-width','0.65px','important');
   path.style.setProperty('stroke-linejoin','round','important');
  });
 });
 target.querySelectorAll('.urgent-whole-bar-outlines').forEach(el=>el.remove());
}

// Independent display-only capacity grouping for both availability charts.
let urgentGroupWithheld=false;
let routineGroupWithheld=false;
for(const [id,isUrgent] of [['availabilityPageGroupWithheld',true],['routineAvailabilityPageGroupWithheld',false]]){
 const input=$(id);
 if(input)input.addEventListener('change',()=>{
  if(isUrgent){urgentGroupWithheld=input.checked;renderAvailabilityChart()}
  else{routineGroupWithheld=input.checked;renderRoutineAvailabilityChart()}
 });
}
function renderAvailabilityChart(){
 const target=$('availabilityChart');
 if($('availabilityPage').classList.contains('hidden')||!window.Plotly||!target)return;
 renderUrgentCategoryToggles();
 renderUrgentSideFilters();
 refreshAvailabilityOptionStates('urgent');
 renderClinicianTypeFilters('urgent');
 updateCollapsedFilterSummaries();
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
  if(!names.includes(category)||!urgentClinicianMatches(r.clinician,r.person)||!urgentConsultMatches(r.slot))continue;
  const status=String(r.availability||'').trim().toLowerCase();
  const label=status==='booked'?category:status==='blocked'?(urgentGroupWithheld?'Held Back':'Blocked'):status==='embargoed'?(urgentGroupWithheld?'Held Back':'Embargoed'):status==='available'?'Available':null;
  if(!label)continue;
  const key=JSON.stringify([r.date,label]);counts.set(key,(counts.get(key)||0)+r.count);
 }
 const weekly=urgentAvailabilityGranularity==='week';
 const buckets=weekly?Array.from({length:weekCount},(_,i)=>dates[i*7]):dates;
 const bucketFor=date=>weekly?mondayOf(date):date;
 if(weekly){
  const aggregated=new Map();
  for(const [key,count] of counts){const [date,label]=JSON.parse(key);const k=JSON.stringify([bucketFor(date),label]);aggregated.set(k,(aggregated.get(k)||0)+count);}
  counts.clear();for(const [key,count] of aggregated)counts.set(key,count);
 }
 const labels=[...names,...(urgentGroupWithheld?['Held Back']:['Blocked','Embargoed']),'Available'];
 const colours=Object.fromEntries(names.map(n=>[n,urgentColour(n)]));
 Object.assign(colours,{Blocked:'#435569',Embargoed:'#586674','Held Back':'#48535f',Available:'#ffffff'});
 const totals=buckets.map(date=>labels.reduce((a,n)=>a+(counts.get(JSON.stringify([date,n]))||0),0));
 const max=Math.ceil(Math.max(1,...totals)*1.2/10)*10;
 const daySummary=buckets.map(date=>{
  const weekday=weekly?'Week commencing '+new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric',timeZone:'UTC'}):new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'short',year:'numeric',timeZone:'UTC'});
  return simpleAvailabilityHover(weekday,labels,labels.map(n=>counts.get(JSON.stringify([date,n]))||0),totals[buckets.indexOf(date)]);
 });
 const traces=labels.map(name=>({type:'bar',name,showlegend:false,x:buckets,y:buckets.map(date=>(weekly||weekdays.has(date))?counts.get(JSON.stringify([date,name]))||0:null),marker:{color:colours[name],line:{width:0}},customdata:daySummary,hovertemplate:'%{customdata}<extra></extra>'}));
 for(const name of labels)traces.push({type:'scatter',mode:'markers',name,x:[null],y:[null],marker:{symbol:'circle',size:14,color:colours[name],line:{color:name==='Available'?'#aeb8c3':colours[name],width:name==='Available'?1.5:0}},hoverinfo:'skip',showlegend:true});
 const weekLabels=Array.from({length:weekCount},(_,i)=>{const offset=i-urgentWeeksBack;return offset===0?'This Week':offset===-1?'Last Week':offset<0?Math.abs(offset)+' weeks ago':offset===1?'Next Week':offset+' weeks'});
 const annotations=weekLabels.map((text,i)=>({x:weekly?buckets[i]:dates[i*7+2],xref:'x',y:1.045,yref:'paper',text,showarrow:false,font:{size:13,color:'#344b60'},xanchor:'center',yanchor:'bottom'}));
 const shapes=Array.from({length:weekCount-1},(_,i)=>({type:'line',xref:'x',yref:'paper',x0:weekly?i+0.5:i*7+5.5,x1:weekly?i+0.5:i*7+5.5,y0:0,y1:1,line:{color:'#bdcbd6',width:1,dash:'dot'}}));
 if(dateSet.has(todayLondon)){
  const x=weekly?urgentWeeksBack-0.5+((new Date(todayLondon+'T12:00:00Z').getUTCDay()+6)%7)/7:dates.indexOf(todayLondon);
  shapes.unshift({type:'rect',xref:'x',yref:'paper',x0:-.5,x1:x,y0:0,y1:1,fillcolor:'rgba(97,112,130,0.12)',line:{width:0},layer:'above'});
  shapes.push({type:'line',xref:'x',yref:'paper',x0:x,x1:x,y0:0,y1:1,line:{color:'#257cb6',width:2.5,dash:'dot'},layer:'above'});
  annotations.push({x,xref:'x',y:1.005,yref:'paper',text:'Today',showarrow:false,font:{size:17,color:'#206ca6'},xanchor:'left',yanchor:'top',xshift:7,yshift:-8});
 }
 buckets.forEach((date,i)=>{if((weekly||weekdays.has(date))&&totals[i])annotations.push({x:date,xref:'x',y:totals[i],yref:'y',text:String(totals[i]),showarrow:false,font:{size:weekly?13:10,color:'#425b70'},yshift:15,yanchor:'bottom'})});
 const note=$('availabilityNote');if(note)note.textContent=names.length?'':'Select an Urgent detailed urgency category to display appointments.';
 Plotly.react(target,traces,{
  autosize:true,height:555,barmode:'stack',barcornerradius:3,bargap:weekly?0.11:0.13,margin:{l:62,r:18,t:85,b:100},paper_bgcolor:'white',plot_bgcolor:'#e1eaf1',
  legend:{orientation:'h',x:0.5,y:-0.23,xanchor:'center',yanchor:'top',font:{size:13,color:'#42576a'},bgcolor:'rgba(255,255,255,0)'},
  xaxis:{type:'category',categoryorder:'array',categoryarray:buckets,tickmode:'array',tickvals:buckets,ticktext:buckets.map(date=>weekly?'':weekdays.has(date)?new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'short',timeZone:'UTC'}):''),tickangle:-90,tickfont:{size:12,color:'#444444'},showgrid:false,range:[-.5,buckets.length-.5],fixedrange:true},
  yaxis:{title:{text:'Appointment slots',font:{size:12,color:'#708397'}},tickfont:{size:12,color:'#52677b'},range:[0,max],gridcolor:'#e0e8ef',gridwidth:1,zeroline:false,showline:false},annotations,shapes,hovermode:'closest',hoverlabel:{bgcolor:'#ffffff',bordercolor:'#dce5ed',font:{family:'Arial',size:12,color:'#182b3b'},align:'left',namelength:0}
 },{responsive:true,displaylogo:false,displayModeBar:false}).then(()=>applyFlatAvailabilityBars(target));
}
window.addEventListener('resize',()=>{if(!$('availabilityPage').classList.contains('hidden'))renderAvailabilityChart()});
setupUrgentRangeControls();

// Independent filters for Routine Appointment Availability.
const routineDisabledTypes=new Set();
function routineDetailedTypes(){
 return [...new Set([...Object.keys(groupUrgencyAssignments),...Object.values(assignments)])]
  .filter(name=>name&&String(groupUrgencyAssignments[name]||'').trim().toLowerCase()==='routine')
  .sort((a,b)=>a.localeCompare(b));
}
function renderRoutineCategoryToggles(){
 const root=$('routineCategoryToggles');if(!root)return;
 const names=routineDetailedTypes();
 const old=[...root.querySelectorAll('input[data-filter-name]')].map(el=>el.dataset.filterName);
 if(old.length===names.length&&names.every((name,i)=>name===old[i])){
  root.querySelectorAll('input[data-filter-name]').forEach(input=>input.nextElementSibling?.style.setProperty('--switch-on',urgencyTypeColour(input.dataset.filterName)));
  return;
 }
 root.replaceChildren();
 if(!names.length){root.textContent='Assign detailed types to the Routine urgency group in Slot settings.';return}
 for(const name of names){
  const row=document.createElement('label');row.className='filter-row';
  const sw=document.createElement('span');sw.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.dataset.filterName=name;
  input.checked=!routineDisabledTypes.has(name);input.setAttribute('aria-label','Include '+name);
  const track=document.createElement('span');track.className='switch-track';
  track.style.setProperty('--switch-on',urgencyTypeColour(name));
  const caption=document.createElement('span');caption.textContent=name;
  sw.append(input,track);row.append(sw,caption);root.append(row);
  input.addEventListener('change',()=>{input.checked?routineDisabledTypes.delete(name):routineDisabledTypes.add(name);renderRoutineAvailabilityChart()});
 }
}
const routineDisabledClinicianGroups=new Set();
const routineDisabledConsultModes=new Set();
function renderRoutineAvailabilityFilters(){
 const build=(id,names,disabled,empty)=>{
  const root=$(id);if(!root)return;
  const previous=[...root.querySelectorAll('input')].map(el=>el.getAttribute('data-filter-name'));
  if(previous.length===names.length&&names.every((name,i)=>name===previous[i])){if(id==='routineClinicianFilters')root.querySelectorAll('input[data-filter-name]').forEach(input=>input.nextElementSibling?.style.setProperty('--switch-on',clinicianGroupColour(input.dataset.filterName)));return;}
  root.replaceChildren();
  if(!names.length){const note=document.createElement('span');note.className='hint';note.textContent=empty;root.append(note);return}
  for(const name of names){
   const row=document.createElement('label');row.className='filter-row';
   const switchEl=document.createElement('span');switchEl.className='switch';
   const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');input.setAttribute('data-filter-name',name);
   input.checked=!disabled.has(name);input.setAttribute('aria-label','Include '+name);
   const track=document.createElement('span');track.className='switch-track';if(id==='routineClinicianFilters')track.style.setProperty('--switch-on',clinicianGroupColour(name));
   const title=document.createElement('span');title.textContent=name;
   switchEl.append(input,track);row.append(switchEl,title);root.append(row);
   input.addEventListener('change',()=>{
    changeAvailabilityFilter('routine',disabled,name,input.checked,input);
   });
  }
 };
 build('routineClinicianFilters',clinicianGroupNames,routineDisabledClinicianGroups,'Create clinician groups in Manage categories.');
 build('routineConsultFilters',consultCategoryNames,routineDisabledConsultModes,'Create consult modes in Manage categories.');
}
function routineClinicianMatches(clinician,person){
 const detailed=effectiveClinicianType(clinician,person);
 const group=clinicianGroupAssignments[detailed];
 return clinicianGroupNames.includes(group)&&!routineDisabledClinicianGroups.has(group)&&!routineDisabledClinicianTypes.has(detailed);
}
function routineConsultMatches(slot){
 const mode=consultAssignments[slot];
 return consultCategoryNames.includes(mode)&&!routineDisabledConsultModes.has(mode);
}
let routineStartWeek=-1,routineEndWeek=6;
let routineAvailabilityGranularity='week';
function setRoutineAvailabilityGranularity(mode){
 routineAvailabilityGranularity=mode==='day'?'day':'week';
 const switchInput=$('routineDaysWeeksSwitch');if(switchInput)switchInput.checked=routineAvailabilityGranularity==='week';
 renderRoutineAvailabilityChart();
}
$('routineDaysWeeksSwitch')?.addEventListener('change',event=>setRoutineAvailabilityGranularity(event.target.checked?'week':'day'));

function setupRoutineRangeControls(){
 const start=$('routineWeeksStart'),end=$('routineWeeksEnd');
 if(!start||!end)return;
 const label=n=>n===0?'This week':n===-1?'Last week':n<0?Math.abs(n)+' weeks ago':n===1?'Next week':n+' weeks ahead';
 const update=changed=>{
  let a=Number(start.value),b=Number(end.value);
  a=Math.max(-2,Math.min(0,a));
  b=Math.max(1,Math.min(7,b));
  start.value=a;end.value=b;
  routineStartWeek=a;routineEndWeek=b;
  $('routineWeeksStartValue').textContent=label(a);
  $('routineWeeksEndValue').textContent=label(b);
  const left=(a+2)/9*100,right=(b+2)/9*100;
  const active=$('routineRangeActive');active.style.left='calc(12px + (100% - 24px)*'+(left/100)+')';active.style.width='calc((100% - 24px)*'+((right-left)/100)+')';
  renderRoutineAvailabilityChart();
 };
 start.addEventListener('input',()=>update('start'));
 end.addEventListener('input',()=>update('end'));
 update('start');
}
// Visual-only finish for the Routine stacked bars. The original Plotly data,
// hover targets, shading and Today geometry are left untouched.
function renderRoutineAvailabilityChart(){
 const target=$('routineAvailabilityChart');
 if(!target||$('routineAvailabilityPage').classList.contains('hidden')||!window.Plotly)return;
 renderRoutineAvailabilityFilters();
 renderRoutineCategoryToggles();
 refreshAvailabilityOptionStates('routine');
 renderClinicianTypeFilters('routine');
 updateCollapsedFilterSummaries();
 const isDaily=routineAvailabilityGranularity==='day';
 const start=new Date(thisWeek+'T12:00:00Z');start.setUTCDate(start.getUTCDate()+7*routineStartWeek);
 const weeks=[];
 for(let i=0;i<=routineEndWeek-routineStartWeek;i++){const d=new Date(start);d.setUTCDate(start.getUTCDate()+7*i);weeks.push(d.toISOString().slice(0,10));}
 const dates=[];
 if(isDaily){
  for(const w of weeks)for(let d=0;d<7;d++){const dt=new Date(w+'T12:00:00Z');dt.setUTCDate(dt.getUTCDate()+d);dates.push(dt.toISOString().slice(0,10));}
 }else dates.push(...weeks);
 const byBucket=new Map();
 let matched=0;
 const bucketSet=new Set(dates);
 for(const r of availabilityRecords){
  const detailed=String(assignments[r.slot]||'').trim();
  const group=String(groupUrgencyAssignments[detailed]||'').trim();
  if(group.toLowerCase()!=='routine'||routineDisabledTypes.has(detailed)||!routineClinicianMatches(r.clinician,r.person)||!routineConsultMatches(r.slot))continue;
  const bucket=isDaily?r.date:mondayOf(r.date);
  if(!bucketSet.has(bucket))continue;
  const status=String(r.availability||'').trim().toLowerCase();
  const label=({booked:'Booked',blocked:routineGroupWithheld?'Held Back':'Blocked',embargoed:routineGroupWithheld?'Held Back':'Embargoed',available:'Available'})[status];
  if(!label)continue;
  matched+=r.count;
  const key=JSON.stringify([bucket,label]);byBucket.set(key,(byBucket.get(key)||0)+r.count);
  if(label==='Booked'){const typeKey=JSON.stringify([bucket,'Booked',detailed]);byBucket.set(typeKey,(byBucket.get(typeKey)||0)+r.count);}
 }
 const labels=['Booked',...(routineGroupWithheld?['Held Back']:['Blocked','Embargoed']),'Available'];
 const colours={'Booked':'#5fc18a','Blocked':'#435569','Embargoed':'#586674','Held Back':'#48535f','Available':'#ffffff'};
 const dateText=d=>new Date(d+'T12:00:00Z').toLocaleDateString('en-GB',{day:'numeric',month:'short',timeZone:'UTC'});
 const summaries=dates.map(date=>{
  const heading=isDaily?new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'short',timeZone:'UTC'}):'Week commencing '+dateText(date);
  return simpleAvailabilityHover(heading,[...routineDetailedTypes().filter(name=>!routineDisabledTypes.has(name)),...labels.filter(name=>name!=='Booked')],[...routineDetailedTypes().filter(name=>!routineDisabledTypes.has(name)).map(name=>byBucket.get(JSON.stringify([date,'Booked',name]))||0),...labels.filter(name=>name!=='Booked').map(name=>byBucket.get(JSON.stringify([date,name]))||0)],labels.reduce((sum,label)=>sum+(byBucket.get(JSON.stringify([date,label]))||0),0));
 });
 const plotBackground='#e1eaf1';
 const routineTypes=routineDetailedTypes().filter(name=>!routineDisabledTypes.has(name));
 const traces=[
  ...routineTypes.map(name=>({type:'bar',name,showlegend:true,x:dates,y:dates.map(date=>byBucket.get(JSON.stringify([date,'Booked',name]))||0),marker:{color:urgencyTypeColour(name),line:{width:0}},customdata:summaries,hovertemplate:'%{customdata}<extra></extra>'})),
  ...labels.filter(label=>label!=='Booked').map(label=>({type:'bar',name:label,showlegend:false,x:dates,y:dates.map(date=>byBucket.get(JSON.stringify([date,label]))||0),marker:{color:colours[label],line:{width:0}},customdata:summaries,hovertemplate:'%{customdata}<extra></extra>'}))
 ];
 for(const label of labels.filter(label=>label!=='Booked'))traces.push({type:'scatter',mode:'markers',name:label,x:[null],y:[null],marker:{symbol:'circle',size:14,color:colours[label],line:{color:label==='Available'?'#a7b8c6':colours[label],width:label==='Available'?1.5:0}},showlegend:true,hoverinfo:'skip'});
 const totals=dates.map(date=>labels.reduce((sum,label)=>sum+(byBucket.get(JSON.stringify([date,label]))||0),0));
 const max=Math.max(1,...totals);
 const annotations=totals.flatMap((total,i)=>total?[{xref:'x',x:i,y:total,text:total.toLocaleString('en-GB'),showarrow:false,yshift:15,font:{size:isDaily?10:13,color:'#425b70'}}]:[]);
 const weekTitles=weeks.map((_,i)=>{const offset=routineStartWeek+i;return offset===0?'This Week':offset===-1?'Last Week':offset<0?Math.abs(offset)+' weeks ago':offset===1?'Next Week':offset+' weeks'});
 for(let i=0;i<weeks.length;i++)annotations.push({xref:'x',yref:'paper',x:isDaily?i*7+3:i,y:1.105,text:weekTitles[i],showarrow:false,font:{family:'Arial',size:13,color:'#344b60'}});
 const dayIndex=(new Date(todayLondon+'T12:00:00Z').getUTCDay()+6)%7;
 const weekIndex=weeks.indexOf(thisWeek);
 const todayX=weekIndex<0?null:isDaily?dates.indexOf(todayLondon):weekIndex-0.5+(Math.min(dayIndex,4)*0.2+0.05);
 const shapes=todayX===null?[]:[
  {type:'rect',xref:'x',yref:'paper',x0:-0.5,x1:todayX,y0:0,y1:1,fillcolor:'rgba(97,112,130,0.12)',line:{width:0},layer:'above'},
  {type:'line',xref:'x',yref:'paper',x0:todayX,x1:todayX,y0:0,y1:1,line:{color:'#257cb6',width:2.5,dash:'dot'},layer:'above'}
 ];
 for(let i=0;i<weeks.length-1;i++){const boundary=isDaily?(i+1)*7-0.5:i+0.5;shapes.push({type:'line',xref:'x',yref:'paper',x0:boundary,x1:boundary,y0:0,y1:1,line:{color:'#bdcbd6',width:1,dash:'dot'},layer:'above'});}
 if(todayX!==null)annotations.push({xref:'x',yref:'paper',x:todayX,y:0.96,text:'Today',showarrow:false,xanchor:'left',yanchor:'top',xshift:7,font:{size:17,color:'#206ca6'}});
 const note=$('routineAvailabilityNote');
 if(note)note.textContent=matched?'':"No matching Routine urgency-group appointments found. Check Manage categories → Urgency Group.";
 Plotly.react(target,traces,{
  height:555,autosize:true,barmode:'stack',barcornerradius:3,bargap:isDaily?0.13:0.11,
  margin:{l:62,r:18,t:85,b:100},paper_bgcolor:'#fff',plot_bgcolor:plotBackground,
  legend:{orientation:'h',x:0.5,y:-0.20,xanchor:'center',yanchor:'top',font:{size:13,color:'#42576a'},traceorder:'normal'},
  xaxis:{type:'category',categoryorder:'array',categoryarray:dates,range:[-0.5,dates.length-0.5],tickmode:'array',tickvals:dates,ticktext:isDaily?dates.map(date=>new Date(date+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'short',timeZone:'UTC'})):dates.map(()=>''),tickangle:-90,tickfont:{size:10},showticklabels:isDaily,showgrid:false,showline:true,linecolor:'#a9b9c7',linewidth:1.5,zeroline:false,fixedrange:true},
  yaxis:{title:{text:'Appointment slots',font:{size:12,color:'#708397'}},tickfont:{size:12,color:'#52677b'},range:[0,Math.ceil(max*1.18/10)*10],gridcolor:'#e0e8ef',gridwidth:1,zeroline:false,showline:false},
  annotations,shapes,hovermode:'closest',hoverlabel:{bgcolor:'#ffffff',bordercolor:'#dce5ed',font:{family:'Arial',size:12,color:'#182b3b'},align:'left',namelength:0}
 },{responsive:true,displaylogo:false,displayModeBar:false}).then(()=>applyFlatAvailabilityBars(target));
}
window.addEventListener('resize',()=>{if(!$('routineAvailabilityPage').classList.contains('hidden'))renderRoutineAvailabilityChart()});
setupRoutineRangeControls();


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
 if(current.length===names.length&&current.every((name,i)=>name===names[i])){root.querySelectorAll('input[data-filter-name]').forEach(input=>{const colour=id==='bookingTimesUrgencyFilters'?urgencyGroupColour(input.dataset.filterName):id==='bookingTimesClinicianFilters'?clinicianGroupColour(input.dataset.filterName):'#596675';input.nextElementSibling?.style.setProperty('--switch-on',colour)});return;}
 root.replaceChildren();
 if(!names.length){const span=document.createElement('span');span.className='hint';span.textContent='Set up categories in Manage categories.';root.append(span);return}
 for(const name of names){
  const label=document.createElement('label');label.className='filter-row';
  const switchWrap=document.createElement('span');switchWrap.className='switch';
  const input=document.createElement('input');input.type='checkbox';input.setAttribute('role','switch');
  input.dataset.filterName=name;input.checked=!disabled.has(name);input.setAttribute('aria-label','Include '+name);
  const track=document.createElement('span');track.className='switch-track';
  if(id==='bookingTimesUrgencyFilters')track.style.setProperty('--switch-on',urgencyGroupColour(name));else if(id==='bookingTimesClinicianFilters')track.style.setProperty('--switch-on',clinicianGroupColour(name));else track.style.setProperty('--switch-on','#596675');
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
  const detailedClinician=effectiveClinicianType(r.clinician,r.person);
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
 },{responsive:true,displaylogo:false,displayModeBar:false});
 const note=$('bookingTimesNote');
 if(note)note.textContent=count?count.toLocaleString('en-GB')+' booked appointments across the loaded CSV files and selected categories.':'No bookings match these filters. Check category assignments in Manage categories.';
}
window.addEventListener('resize',()=>{if(!$('bookingTimesPage').classList.contains('hidden'))renderBookingTimesChart()});
