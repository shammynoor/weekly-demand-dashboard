const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const core=require('../data-correctness.js');
const app=fs.readFileSync(__dirname+'/../app.js','utf8');
test('London calendar handles summer midnight and year rollover',()=>{
 assert.equal(core.dateContext(new Date('2026-06-30T23:30:00Z')).today,'2026-07-01');
 assert.equal(core.dateContext(new Date('2026-12-31T23:30:00Z')).today,'2026-12-31');
 assert.deepEqual(core.dateContext(new Date('2027-01-01T00:01:00Z')),{today:'2027-01-01',week:'2026-12-28',month:'2027-01'});
});
test('coverage excludes current, partial and missing weeks',()=>{
 assert.equal(core.completeWeeks([{start:'2026-08-31',end:'2026-10-23',asOf:'2026-10-28'}],'2026-10-28',52).length,8);
 const gaps=core.completeWeeks([{start:'2026-08-31',end:'2026-09-04'},{start:'2026-09-14',end:'2026-09-18'}],'2026-10-28',52);
 assert.deepEqual(gaps,['2026-08-31','2026-09-14']);
 assert.deepEqual(core.completeWeeks([{start:'2026-10-20',end:'2026-12-31',asOf:'2026-10-28'}],'2026-10-28',52),[]);
 assert.deepEqual(core.completeWeeks([{start:'2026-10-19',end:'2026-12-31',asOf:'2026-10-23'}],'2026-10-28',52),[]);
});
test('coverage uses per-file unfiltered slot dates and snapshot',()=>{
 const ranges=core.coverageFromEntries([['a',{kind:'slot',version:JSON.stringify(['a','report','2026-10-24T12:00:00Z',1]),maps:[new Map(),new Map([[JSON.stringify(['2026-10-19','phone']),1],[JSON.stringify(['2026-10-23','other']),1]])]}]]);
 assert.deepEqual(ranges,[{start:'2026-10-19',end:'2026-10-23',asOf:'2026-10-24'}]);
});
test('111 and migrated roles persist through renames and explicit choices',()=>{
 let roles=core.roles({},['111 referral','Routine'],['GP']);
 assert.equal(roles.appointmentTypes['111 referral'].kind,'urgent');
 const id=roles.appointmentTypes.Routine.id;
 core.renameRole(roles,'appointmentTypes','Routine','Long term care');
 core.renameRole(roles,'clinicianGroups','GP','Medical team');
 roles=core.roles(roles,['111 referral','Long term care'],['Medical team']);
 assert.equal(roles.appointmentTypes['Long term care'].id,id);
 assert.equal(roles.appointmentTypes['Long term care'].kind,'routine');
 assert.equal(roles.clinicianGroups['Medical team'].doctor,true);
 roles.appointmentTypes['111 referral'].kind='other';
 assert.equal(core.roles(roles,['111 referral'],[]).appointmentTypes['111 referral'].kind,'other');
});
function modelFixture(){
 const original=fs.readFileSync(__dirname+'/../routine-capacity.js','utf8');
 const source=original.slice(0,original.lastIndexOf('})();'))+'globalThis.audit={modelData,capacityFor,ui};})();';
 class FixedDate extends Date{constructor(...a){super(...(a.length?a:['2026-10-28T12:00:00Z']))}}
 const ctx={Date:FixedDate,Intl,console,GPCorrectness:core,window:{},document:{getElementById:id=>id==='routineCapHistory'?{value:'52'}:null,addEventListener(){}},clinicianGroupNames:['Medical team'],isDoctorPlanningGroup:()=>true,appointmentCalculationRole:()=> 'routine',consultAssignments:{F2F:'Face to Face',TEL:'Phone'},consultCategoryNames:['Face to Face','Phone'],suggestConsultMode:()=>'',excludedSlotTypes:new Set(),capacityUrgencyCategory:()=> 'Long term care',effectiveClinicianType:()=> 'Salaried GP',clinicianGroupAssignments:{'Salaried GP':'Medical team'},isDashboardToggleOption:()=>true,prioritySort:(_,a)=>a,records:[],availabilityRecords:[],waitingTrendRecords:[],slotReportCoverage:[{start:'2026-08-31',end:'2026-10-23',asOf:'2026-10-28'}]};
 for(let w=0;w<8;w++)for(let day=0;day<5;day++)for(const slot of ['F2F','TEL']){
 if(slot==='TEL'&&day===4)continue;
 const date=core.shift('2026-08-31',w*7+day),r={slot,date,bookedDay:date,appointmentDate:date,availability:'Booked',count:10,leadMinutes:40320};
 ctx.records.push(r);ctx.availabilityRecords.push(r);ctx.waitingTrendRecords.push(r);
 }
 vm.createContext(ctx);vm.runInContext(source,ctx);return ctx;
}
test('routine phone service without Fridays retains eight complete covered weeks',()=>{
 const c=modelFixture();assert.equal(c.audit.modelData().count,8);
 c.audit.ui.disabledModes.add('Face to Face');
 const model=c.audit.modelData();assert.equal(model.count,8);assert.equal(model.demand,40);assert.equal(model.capacity,40);assert.equal(model.wait,4);
});
test('covered zero weeks stay in denominator; missing report weeks do not',()=>{
 const c=modelFixture();c.records=c.records.filter(r=>r.date>'2026-09-04');c.availabilityRecords=c.availabilityRecords.filter(r=>r.date>'2026-09-04');
 assert.equal(c.audit.modelData().demand,90*7/8);
 c.slotReportCoverage=[{start:'2026-08-31',end:'2026-09-04'},{start:'2026-09-14',end:'2026-10-23'}];
 assert.match(c.audit.modelData().error,/eight complete/);
});
test('invalid leads excluded and capacity floor preserved',()=>{
 const c=modelFixture();c.waitingTrendRecords.push({...c.waitingTrendRecords[0],leadMinutes:Infinity,count:100000});
 assert.equal(c.audit.modelData().wait,4);assert.equal(c.audit.capacityFor({capacity:300,demand:200,wait:1},6,12),300);
});
test('every active chart gets an awaited refresh',async()=>{
 const renderer=app.slice(app.indexOf('async function renderActiveAnalyticsPage()'),app.indexOf('function refreshOnDateRollover()'));
 const ids=['chartPage','waitingTimesPage','waitingOverTimePage','availabilityPage','routineAvailabilityPage','dayView2Page','bookingTimesPage','signpostingPage','dailyForecastPage','weeklyForecastPage','routineCapacityPage'];
 const names=['refresh','renderWaitingTimesChart','renderWaitingTrendChart','renderAvailabilityChart','renderRoutineAvailabilityChart','renderDayView2','renderBookingTimesChart','renderSignpostingChart'];
 let visible='',calls=0;const draw=async()=>{await Promise.resolve();calls++};
 const ctx={updateDashboardDateContext(){},refreshPrimaryApptTypeBoxes(){},$:id=>({classList:{contains:()=>id!==visible}}),window:{refreshForecastCharts:draw,renderRoutineCapacity:draw}};
 names.forEach(n=>ctx[n]=draw);vm.createContext(ctx);vm.runInContext(renderer,ctx);
 for(const id of ids){visible=id;calls=0;await ctx.renderActiveAnalyticsPage();assert.equal(calls,1,id)}
});
test('autosave only clears edits included in the saved serial',()=>{
 const block=app.slice(app.indexOf('  categoryDirty=cloudChangeSerial!==saveSerial;'),app.indexOf('  return true;',app.indexOf('  categoryDirty=cloudChangeSerial!==saveSerial;')));
 for(const newer of [false,true]){const ctx={categoryDirty:true,cloudChangeSerial:newer?2:1,saveSerial:1,cloudNotice(){}};vm.createContext(ctx);vm.runInContext(block,ctx);assert.equal(ctx.categoryDirty,newer)}
});
test('refresh stops before remote settings when save is in flight or fails',async()=>{
 const source=app.slice(app.indexOf('async function refreshDashboardData('));
 for(const scenario of ['busy','failed','newer','success']){
 let loaded=0,feedback='';
 const ctx={settingsRefreshing:false,busy:false,activeAccount:{},selectedFolder:{id:'a'},cloudFolder:null,cloudBusy:scenario==='busy',categoryDirty:true,
 saveCloudCategories:async()=>{if(scenario==='success')ctx.categoryDirty=false;return scenario!=='failed'},
 loadCloudCategories:async()=>{loaded++},loadFolder:async()=>{loaded++},setButtons(){},showRefreshFeedback:s=>{feedback=s}};
 vm.createContext(ctx);vm.runInContext(source,ctx);await ctx.refreshDashboardData(true);
 assert.equal(loaded,scenario==='success'?2:0);assert.equal(ctx.settingsRefreshing,false);
 if(scenario!=='success')assert.match(feedback,/Refresh failed/);
 }
});
test('all production and inline scripts compile',()=>{
 for(const name of ['app.js','forecast.js','routine-capacity.js','data-correctness.js'])new vm.Script(fs.readFileSync(__dirname+'/../'+name,'utf8'),{filename:name});
 const html=fs.readFileSync(__dirname+'/../index.html','utf8');
 for(const m of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi))if(m[1].trim())new vm.Script(m[1]);
 assert.ok(html.indexOf('data-correctness.js')<html.indexOf('src="app.js'));
});

test('settings refresh rejects concurrent edits before applying remote data and resumes autosave',async()=>{
 const begin=app.indexOf('async function loadCloudCategories('),end=app.indexOf('let categoryDirty=',begin);
 const source=app.slice(begin,end);let resolveRead,rescheduled=0;
 const ctx={activeAccount:{},cloudChangeSerial:1,cloudReady:true,categoryDirty:false,cloudBusy:false,
 cloudNotice(){},cloudFile:()=>new Promise(r=>{resolveRead=r}),markCategoryDirty:()=>rescheduled++,validateCloudCategories:()=>assert.fail('must reject before validation'),analyticsRoles:{sentinel:true}};
 vm.createContext(ctx);vm.runInContext(source,ctx);const pending=ctx.loadCloudCategories({strict:true,protectEdits:true});
 ctx.cloudChangeSerial=2;ctx.categoryDirty=true;resolveRead({});
 await assert.rejects(pending,/Settings changed during refresh/);
 assert.equal(ctx.cloudReady,true);assert.equal(rescheduled,1);assert.deepEqual(ctx.analyticsRoles,{sentinel:true});
});
test('settings refresh propagates read failures without losing established readiness',async()=>{
 const begin=app.indexOf('async function loadCloudCategories('),end=app.indexOf('let categoryDirty=',begin);
 const ctx={activeAccount:{},cloudChangeSerial:1,cloudReady:true,categoryDirty:false,cloudBusy:false,cloudNotice(){},cloudFile:async()=>{throw Error('offline')}};
 vm.createContext(ctx);vm.runInContext(app.slice(begin,end),ctx);
 await assert.rejects(ctx.loadCloudCategories({strict:true,protectEdits:true}),/offline/);assert.equal(ctx.cloudReady,true);
});
test('refresh entry points and unchanged CSV branch use the protected redraw path',()=>{
 assert.ok(app.includes("$('refreshData').onclick=()=>refreshDashboardData(true)"));
 assert.ok(app.includes('selectedFolder&&!busy)refreshDashboardData(false)'));
 assert.ok(app.includes('fingerprint===lastFingerprint){await renderActiveAnalyticsPage();'));
 assert.ok(app.includes('slotReportCoverage=GPCorrectness.coverageFromEntries(pendingCache)'));
});
