const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),core=require('../data-correctness.js');
const app=fs.readFileSync(__dirname+'/../app.js','utf8');
const index=m=>Number(m.slice(0,4))*12+Number(m.slice(5,7))-1;
function fixture(prefix='demandRange',store={},user='one',timestamp='2026-10-27T12:30:00Z'){
 const nodes={},months=['2026-01','2026-05','2026-10','2026-12'];let account=user;
 for(const name of ['Start','End','FromLabel','ToLabel','Fill'])nodes[prefix+name]={value:name==='End'?'1':'0',dataset:{},style:{},parentElement:{clientWidth:400,getBoundingClientRect:()=>({left:0})},clientWidth:400,getBoundingClientRect:()=>({left:0}),handlers:{},addEventListener(type,fn){this.handlers[type]=fn}};
 const ctx={GPCorrectness:core,capacityCSVLatestModified:timestamp===null?null:Date.parse(timestamp),readChartPreferences:()=>store[account]||{},writeChartPreference:(k,v)=>{(store[account]??={})[k]=v},$:id=>nodes[id],availabilityRecords:months.map(m=>({date:m+'-01'})),waitingTimeRecords:[],waitingRecordMonths:()=>months,waitingTrendAvailableMonths:months,
 demandMonthBounds:'',demandMonthStart:null,demandMonthEnd:null,waitingAppointmentMonthBounds:'',waitingAppointmentMonthStart:null,waitingAppointmentMonthEnd:null,waitingTrendMonthBounds:'',waitingTrendMonthStart:null,waitingTrendMonthEnd:null,scheduleDemandRefresh(){}};
 vm.createContext(ctx);
 vm.runInContext(app.slice(app.indexOf('function monthRangeSelection('),app.indexOf('function renderDemandUrgencyToggles(')),ctx);
 vm.runInContext(app.slice(app.indexOf('function setupWaitingAppointmentSlider('),app.indexOf('function waitingMatches(')),ctx);
 const setup=prefix==='demandRange'?ctx.setupDemandMonthSlider:prefix==='waitingRange'?ctx.setupWaitingAppointmentSlider:ctx.setupWaitingTrendSlider;
 ctx.renderWaitingTimesChart=setup;ctx.renderWaitingTrendChart=setup;
 return {ctx,nodes,store,setup,setAccount:a=>{account=a},move(handle,month){nodes[prefix+handle].value=String(index(month));nodes[prefix+handle].handlers.input()}};
}
for(const prefix of ['demandRange','waitingRange','waitingTrendRange']){
 test(prefix+': restores both handles after reload and before first page visit',()=>{
  const f=fixture(prefix);f.setup();f.move('Start','2026-05');f.move('End','2026-09');
  const reloaded=fixture(prefix,f.store);reloaded.setup();
  assert.equal(reloaded.nodes[prefix+'Start'].value,String(index('2026-05')));assert.equal(reloaded.nodes[prefix+'End'].value,String(index('2026-09')));
  assert.equal(reloaded.nodes[prefix+'ToLabel'].textContent,'End of Sept 2026');
 });
 test(prefix+': independent accounts and narrower refreshed data keep saved selections intact',()=>{
  const f=fixture(prefix);f.setup();f.move('Start','2026-05');f.move('End','2026-09');f.setAccount('two');f.setup();
  assert.equal(f.nodes[prefix+'Start'].value,String(index('2026-01')));assert.equal(f.nodes[prefix+'End'].value,String(index('2026-10')));
  f.setAccount('one');f.ctx.capacityCSVLatestModified=Date.parse('2026-06-20T12:00Z');f.setup();assert.equal(f.nodes[prefix+'End'].value,String(index('2026-06')));
  f.ctx.capacityCSVLatestModified=Date.parse('2026-10-27T12:30Z');f.setup();assert.equal(f.nodes[prefix+'End'].value,String(index('2026-09')));
 });
 test(prefix+': stale refresh date replaces today; partial day never says End',()=>{
  const f=fixture(prefix,{},'one','2026-09-30T12:00:00Z');f.setup();
  assert.equal(f.nodes[prefix+'End'].max,String(index('2026-09')));assert.equal(f.nodes[prefix+'ToLabel'].textContent,'30 Sept 2026');
  f.ctx.capacityCSVLatestModified=Date.parse('2026-10-27T23:30:00Z');f.setup();assert.equal(f.nodes[prefix+'ToLabel'].textContent,'27 Oct 2026');
 });
 test(prefix+': crossing handles saves a valid pair and legacy individual values migrate',()=>{
  const page=prefix==='demandRange'?'chartPage':prefix==='waitingRange'?'waitingTimesPage':'waitingOverTimePage';
  const saved={one:{['range:'+page+':'+prefix+'Start']:String(index('2026-05')),['range:'+page+':'+prefix+'End']:String(index('2026-09'))}};
  const f=fixture(prefix,saved);f.setup();assert.equal(f.nodes[prefix+'Start'].value,String(index('2026-05')));
  f.move('Start','2026-10');assert.equal(f.nodes[prefix+'End'].value,String(index('2026-10')));
  assert.equal(f.store.one['monthRange:'+prefix].end,'2026-10');f.move('End','2026-01');assert.equal(f.nodes[prefix+'Start'].value,String(index('2026-01')));
 });
}
test('London summer midnight and unknown timestamps do not claim a complete refresh day',()=>{
 const f=fixture('demandRange',{},'one','2026-08-31T23:30:00Z');f.setup();assert.equal(f.nodes.demandRangeToLabel.textContent,'1 Sept 2026');
 f.ctx.capacityCSVLatestModified=null;f.setup();assert.equal(f.nodes.demandRangeToLabel.textContent,'Dec 2026');
});
test('delayed dataset initialization retains both saved month handles',()=>{
 const f=fixture();f.setup();f.move('Start','2026-05');f.move('End','2026-09');
 const reloaded=fixture('demandRange',f.store);const records=reloaded.ctx.availabilityRecords;reloaded.ctx.availabilityRecords=[];reloaded.setup();reloaded.ctx.availabilityRecords=records;reloaded.setup();
 assert.equal(reloaded.nodes.demandRangeStart.value,String(index('2026-05')));assert.equal(reloaded.nodes.demandRangeEnd.value,String(index('2026-09')));
});
function controlsFixture(){
 const store={one:{}},nodes=[];let account='one';
 const ctx={Map,Event:class{constructor(type){this.type=type}},window:{gpChartPreferences:{}},document:{querySelectorAll:()=>nodes},chartPreferenceAccount:()=>account,readChartPreferences:()=>store[account]||{},chartControlKey:e=>e.id,chartControlValue:e=>e.type==='checkbox'?e.checked:e.value};
 vm.createContext(ctx);vm.runInContext(app.slice(app.indexOf('let chartPreferenceRestoring='),app.indexOf('function rememberChartControl(')),ctx);
 function add(id,type,value){const el={id,type,value:type==='checkbox'?'on':value,checked:value,tagName:'INPUT',dataset:{bound:'1'},dispatchEvent:e=>{el.handler?.(e)}};nodes.push(el);return el}
 return {ctx,store,add,setAccount:v=>{account=v}};
}
test('other UI controls restore per user and reset to defaults for a new user',()=>{
 const f=controlsFixture(),toggle=f.add('toggle','checkbox',false),slider=f.add('slider','range','2');f.store.one={toggle:true,slider:'6'};f.ctx.restoreChartCheckboxes();assert.equal(toggle.checked,true);assert.equal(slider.value,'6');
 f.setAccount('two');f.ctx.restoreChartCheckboxes();assert.equal(toggle.checked,false);assert.equal(slider.value,'2');
 f.setAccount('one');f.ctx.restoreChartCheckboxes();assert.equal(toggle.checked,true);assert.equal(slider.value,'6');
});
test('paired generic sliders both reach their saved values even when the first handler redraws',()=>{
 const f=controlsFixture(),start=f.add('start','range','0'),end=f.add('end','range','7');f.store.one={start:'2',end:'5'};
 start.handler=()=>{end.value='7'};let final;end.handler=()=>{final=[start.value,end.value]};
 f.ctx.restoreChartCheckboxes();assert.deepEqual(final,['2','5']);
});

test('preference storage is account-scoped and signed-out writes are ignored',()=>{
 const storage=new Map(),ctx={activeAccount:{homeAccountId:'one'},localStorage:{getItem:k=>storage.get(k),setItem:(k,v)=>storage.set(k,v)},console};vm.createContext(ctx);
 vm.runInContext(app.slice(app.indexOf('function chartPreferenceAccount('),app.indexOf('window.gpChartPreferences=')),ctx);
 ctx.writeChartPreference('test',1);ctx.activeAccount={homeAccountId:'two'};assert.equal(ctx.readChartPreferences().test,undefined);ctx.writeChartPreference('test',2);
 ctx.activeAccount={homeAccountId:'one'};assert.equal(ctx.readChartPreferences().test,1);ctx.activeAccount=null;ctx.writeChartPreference('test',3);assert.equal(storage.size,2);
});
test('Day View date and filters restore independently when accounts change',()=>{
 let account='one';const dates={value:''},saved={one:{'dayView2:selectedDate':'2026-10-21','dayView2:disabledUrgencies':['Routine']},two:{}};
 const ctx={chartPreferenceAccount:()=>account,readChartPreferences:()=>saved[account],$:()=>dates,writeChartPreference:(k,v)=>saved[account][k]=v,closeDayView2Calendar(){},renderDayView2(){},Set};vm.createContext(ctx);
 vm.runInContext(app.slice(app.indexOf('const dayView2DisabledUrgencies='),app.indexOf('let dayView2SegmentObserver=')),ctx);
 ctx.syncDayView2Preferences();assert.equal(dates.value,'2026-10-21');ctx.chooseDayView2Date('2026-10-22');assert.equal(saved.one['dayView2:selectedDate'],'2026-10-22');
 account='two';ctx.syncDayView2Preferences();assert.equal(dates.value,'');ctx.saveDayView2Preferences();assert.equal(saved.two['dayView2:disabledUrgencies'].length,0);
 account='one';ctx.syncDayView2Preferences();assert.equal(dates.value,'2026-10-22');
});
test('Routine Capacity custom target selection restores per account and rejects invalid targets',()=>{
 const source=fs.readFileSync(__dirname+'/../routine-capacity.js','utf8');let account='one';const saved={one:{'routineCapacity:target':5},two:{},three:{'routineCapacity:target':99}};
 const ctx={window:{gpChartPreferences:{account:()=>account,read:()=>saved[account]}},Set};vm.createContext(ctx);
 vm.runInContext(source.slice(source.indexOf('const ui='),source.indexOf('function isRoutine('))+'\nfunction target(){return ui.target}',ctx);
 ctx.syncRoutineCapacityPreferences();assert.equal(ctx.target(),5);account='two';ctx.syncRoutineCapacityPreferences();assert.equal(ctx.target(),2);account='three';ctx.syncRoutineCapacityPreferences();assert.equal(ctx.target(),2);account='one';ctx.syncRoutineCapacityPreferences();assert.equal(ctx.target(),5);
});

test('settings searches and filters save as UI choices while practice edits stay excluded',()=>{
 const ctx={};vm.createContext(ctx);vm.runInContext(app.slice(app.indexOf('function chartControlKey('),app.indexOf('function chartControlValue(')),ctx);
 const control=(id,page,type='text')=>({id,type,tagName:'INPUT',matches:()=>true,closest:()=>({id:page})});
 assert.ok(ctx.chartControlKey(control('slotTableSearch','categoryPage')));
 assert.ok(ctx.chartControlKey(control('peopleGroupFilter','peoplePage')));
 assert.equal(ctx.chartControlKey(control('newClinicianCategory','categoryPage')),null);
 assert.equal(ctx.chartControlKey(control('practiceName','practiceSettingsPage')),null);
 assert.equal(ctx.chartControlKey(control('personalName','personalSettingsPage')),null);
 assert.equal(ctx.chartControlKey(control('demandRangeStart','chartPage','range')),null);
});
test('settings workspace tabs and sorting restore saved choices or account defaults',()=>{
 let saved={'settings:slotWorkspace':'mapping','settings:categoryTab':'tabClinician','settings:slotSort':'count-desc'};const clicks=[];
 const buttons=(field,values)=>values.map(value=>({dataset:{[field]:value},click:()=>clicks.push(value)}));
 const workspace=buttons('slotWorkspace',['inclusion','mapping']),tabs=buttons('categoryTab',['tabUrgency','tabClinician']);let sort;
 const ctx={readChartPreferences:()=>saved,document:{querySelectorAll:q=>q.includes('slot-workspace')?workspace:tabs},setSlotMappingSort:v=>{sort=v}};vm.createContext(ctx);
 vm.runInContext(app.slice(app.indexOf('function restorePersonalSettingsWorkspace('),app.indexOf("document.addEventListener('click',event=>{",app.indexOf('function restorePersonalSettingsWorkspace('))),ctx);
 ctx.restorePersonalSettingsWorkspace();assert.deepEqual(clicks,['mapping','tabClinician']);assert.equal(sort,'count-desc');
 saved={};clicks.length=0;ctx.restorePersonalSettingsWorkspace();assert.deepEqual(clicks,['inclusion','tabUrgency']);assert.equal(sort,'name');
});
