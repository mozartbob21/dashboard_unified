// Offline browser contract tests. All source rows and server responses are synthetic.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),test=require('node:test');
const root=process.env.COLLECTIVE_UI_ROOT || path.join(__dirname,'..');
const source=fs.readFileSync(path.join(root,'static/collective.js'),'utf8');
function harness(realExcel=false,admin=false){
 const elements=new Map(),values=new Map(),sessions=new Map(),pending=[],downloads=[],windowEvents={};
 class Element{
  constructor(){this.value='';this.dataset={};this.hidden=true;this.disabled=false;this.checked=false;this.open=false;this.innerHTML='';this.textContent='';this.children=new Map();this.classList={toggle(){},contains(){return false;}};this.listeners={};}
  querySelector(k){if(!this.children.has(k))this.children.set(k,new Element());return this.children.get(k);}querySelectorAll(){return [];}addEventListener(k,fn){this.listeners[k]=fn;}setAttribute(){}focus(){}click(){}remove(){}appendChild(){}showModal(){this.open=true;}close(){this.open=false;}
 }
 const document={querySelectorAll:()=>[],addEventListener(){},createElement:()=>new Element(),body:new Element(),querySelector(k){if(!elements.has(k))elements.set(k,new Element());return elements.get(k);},getElementById(k){return this.querySelector('#'+k);}};
 document.querySelector('.collective-dashboard').dataset={owner:'alice',canManage:admin?'true':'false'};
 document.getElementById('periodSel').value='d7';
 let XLSX={utils:{book_new:()=>[],aoa_to_sheet:data=>({data}),book_append_sheet(wb,sheet,name){wb.push({sheet,name});}},writeFile:wb=>downloads.push(wb)};
 if(realExcel){const lib=require(process.env.COLLECTIVE_XLSX_PATH || path.join(__dirname,'../static/vendor/xlsx.full.min.js'));XLSX={...lib,writeFile(wb){downloads.push(lib.read(lib.write(wb,{type:'buffer',bookType:'xlsx'}),{type:'buffer'}));}};}
 const storage=map=>({getItem:k=>map.get(k) ?? null,setItem:(k,v)=>map.set(k,v),removeItem:k=>map.delete(k)});
 const context={document,localStorage:storage(values),sessionStorage:storage(sessions),fetch:(...args)=>new Promise((resolve,reject)=>pending.push({args,resolve,reject})),AbortController,clearTimeout(){},setInterval(){},setTimeout(fn,ms){if(ms<45000)queueMicrotask(fn);return 1;},Blob,URL:{createObjectURL:()=> 'blob:synthetic',revokeObjectURL(){}},window:{XLSX,addEventListener(k,fn){windowEvents[k]=fn;}},XLSX,console,alert(){},crypto:{randomUUID:()=> 'synthetic-idempotency-key'}};
 vm.createContext(context);vm.runInContext(source.replace(/\}\)\(\);\s*$/,'globalThis.testApi={classify,keyOf,store,load,state,baseFiltered,exportAll,exportPptx,refresh,openCard,captureDraft,saveReport,changeField,useServerReport,requestCloseCard,assigneeSummary,legacyThemeCandidates,migrateLegacyThemes,addAssignee,changeAssigneeActive,syncSavedSnapshot,getRows:()=>ROWS,getPeriod:()=>loadedPeriod,getDraft:()=>current&&DRAFTS.get(current._key)};})();'),context);
 return{api:context.testApi,document,pending,values,sessions,downloads,windowEvents};
}
const head=['ОМСУ','Аннотация/краткое содержание','Подписи (фактическое число подписей)','Дата поступления обращения','Повтор','Статус','Номер обращения ЕЦУР, МСЭД'];
const rows=[['Первый','Нет воды','101','01.01.2026','Да','В работе','a'],['Второй','Капитальный ремонт кровли','100','01.01.2026','','В работе','b'],['Первый','Освещение двора','90','01.01.2026','','Не учитывается','c']];
const wf=(extra={})=>({version:0,assignee_id:null,assignee_name:'',assignee_archived:false,report:'',theme:null,can_assign:true,...extra});
const extra=()=>({keys:['c_a','c_b','c_c'],workflows:{c_a:wf(),c_b:wf(),c_c:wf({can_assign:false})},assignees:[{id:1,name:'Первый отдел',archived:false,version:1},{id:2,name:'Архивный отдел',archived:true,version:2}],since:'2026-01-01'});
const data=(extras={})=>({head,rows,meta:'Synthetic',updated:'19.10.2026',start:'2026-01-01',end:'2026-01-02',...extra(),...extras});
const reply=(data,code=200)=>({ok:code<400,status:code,json:async()=>data});
const flush=()=>new Promise(resolve=>setImmediate(resolve));
function load(h,x=extra()){h.api.load(head,rows,'Synthetic',x);}

test('classification remains stable, server corrections override automatic themes and old local corrections require explicit migration',()=>{
 const h=harness(),a=h.api;
 assert.equal(a.classify({'Аннотация/краткое содержание':'Капитальный ремонт водопровода'}),'kr');
 assert.equal(a.classify({'Аннотация/краткое содержание':'Нет газа'}),'ii');assert.equal(a.classify({'Аннотация/краткое содержание':'Покос газона'}),'pr');
 a.store.set(a.keyOf({'Номер обращения ЕЦУР, МСЭД':'a'}),'pr');load(h);assert.equal(a.getRows()[0]._theme,'ii');assert.equal(a.legacyThemeCandidates().length,1);
 const x=extra();x.workflows.c_a.theme='kr';load(h,x);assert.equal(a.getRows()[0]._theme,'kr');assert.equal(a.legacyThemeCandidates().length,0);
 assert.notEqual(a.keyOf({'Аннотация/краткое содержание':'Без номера А'}),a.keyOf({'Аннотация/краткое содержание':'Без номера Б'}));
});
test('stable keys stay aligned after empty portal rows',()=>{const h=harness();h.api.load(head,[[],rows[0]],'Synthetic',{keys:['empty','actual'],workflows:{actual:wf({report:'saved'})}});assert.equal(h.api.getRows()[0]._key,'actual');assert.equal(h.api.getRows()[0]._workflow.report,'saved');});
test('assignment and unassigned filters obey server eligibility and archived assignees remain visible',()=>{
 const h=harness(),x=extra();x.workflows.c_b=wf({assignee_id:2,assignee_name:'Архивный отдел',assignee_archived:true});load(h,x);
 h.api.state.unassigned=true;assert.equal(h.api.baseFiltered().length,1);h.api.state.unassigned=false;h.api.state.assignee='2';assert.equal(h.api.baseFiltered()[0]._key,'c_b');
 h.api.openCard(1);assert.match(h.document.getElementById('dlgAssignee').innerHTML,/value="2" disabled/);h.api.openCard(0);assert.doesNotMatch(h.document.getElementById('dlgAssignee').innerHTML,/value="2"/);
 const summary=h.api.assigneeSummary(h.api.getRows());assert.equal(summary.length,2);assert.equal(summary.find(o=>o.key==='__unassigned').count,1);
});
test('responsible and manual theme persist through API with optimistic version and no local write',async()=>{
 const h=harness();load(h);h.api.openCard(0);const p=h.api.changeField('assignee_id','1');const req=h.pending.at(-1);assert.equal(req.args[0],'/mingkh/collective/api/appeals/c_a');assert.deepEqual(JSON.parse(req.args[1].body),{version:0,assignee_id:1});req.resolve(reply({workflow:wf({version:1,assignee_id:1,assignee_name:'Первый отдел'})}));await p;assert.equal(h.api.getRows()[0]._workflow.assignee_id,1);
 const t=h.api.changeField('theme','pr');h.pending.at(-1).resolve(reply({workflow:wf({version:2,assignee_id:1,assignee_name:'Первый отдел',theme:'pr'})}));await t;assert.equal(h.api.getRows()[0]._theme,'pr');assert.equal(h.values.size,0);
});
test('manual report is safely escaped, saved once on double click, and stays editable during request',async()=>{
 const h=harness();load(h);h.api.openCard(0);h.document.getElementById('dlgReport').value='<img src=x onerror=alert(1)> Первый текст';h.api.captureDraft();
 const p=h.api.saveReport(),count=h.pending.length;h.api.saveReport();assert.equal(h.pending.length,count);
 h.document.getElementById('dlgReport').value='Продолжаю редактировать';h.api.captureDraft();h.pending.at(-1).resolve(reply({workflow:wf({version:1,report:'<img src=x onerror=alert(1)> Первый текст'})}));await p;
 assert.equal(h.api.getDraft().text,'Продолжаю редактировать');assert.equal(h.api.getDraft().dirty,true);assert.equal(h.api.getRows()[0]._workflow.report,'<img src=x onerror=alert(1)> Первый текст');
 assert.doesNotMatch(h.document.getElementById('list').innerHTML,/<img/);
});
test('network failure preserves report and retry recognizes a previously accepted save',async()=>{
 const h=harness();load(h);h.api.openCard(0);h.document.getElementById('dlgReport').value='Нужный доклад';h.api.captureDraft();const p=h.api.saveReport();h.pending.at(-1).reject(new Error('offline'));await p;assert.equal(h.api.getDraft().text,'Нужный доклад');assert.equal(h.api.getDraft().dirty,true);
 const retry=h.api.saveReport();h.pending.at(-1).resolve(reply({error:'conflict',workflow:wf({version:1,report:'Нужный доклад'})},409));await retry;assert.equal(h.api.getDraft().dirty,false);assert.match(h.document.getElementById('dlgReportStatus').textContent,/уже сохранён/);
});
test('concurrent report conflict preserves draft, presents server text safely and explicit retry uses current version',async()=>{
 const h=harness();load(h);h.api.openCard(0);h.document.getElementById('dlgReport').value='Мой доклад';h.api.captureDraft();const p=h.api.saveReport();h.pending.at(-1).resolve(reply({error:'conflict',workflow:wf({version:2,report:'<script>server</script>'})},409));await p;
 assert.equal(h.api.getDraft().text,'Мой доклад');assert.equal(h.api.getDraft().conflict,true);assert.equal(h.document.getElementById('conflictServerReport').textContent,'<script>server</script>');assert.equal(h.document.getElementById('saveReport').disabled,true);
 const overwrite=h.api.saveReport(true);assert.equal(JSON.parse(h.pending.at(-1).args[1].body).version,2);h.pending.at(-1).resolve(reply({workflow:wf({version:3,report:'Мой доклад'})}));await overwrite;assert.equal(h.api.getDraft().dirty,false);
});
test('closing dirty card preserves draft, Escape is guarded, reload warns and missing keys cannot be saved',()=>{
 const h=harness();load(h);h.api.openCard(0);h.document.getElementById('dlgReport').value='Черновик';h.api.captureDraft();h.api.requestCloseCard();assert.equal(h.document.getElementById('dlg').open,true);assert.equal(h.document.getElementById('closeDraftWarning').hidden,false);
 let warned=false;h.windowEvents.beforeunload({preventDefault(){warned=true;}});assert.equal(warned,true);h.document.getElementById('dlg').close();h.api.openCard(0);assert.equal(h.document.getElementById('dlgReport').value,'Черновик');
 h.api.load(head,rows,'Synthetic');h.api.openCard(1);assert.equal(h.document.getElementById('dlgReport').disabled,true);
});
test('empty report is a valid explicit save and long report is blocked without losing text',async()=>{
 const h=harness(),x=extra();x.workflows.c_a.report='Старый';load(h,x);h.api.openCard(0);h.document.getElementById('dlgReport').value='';h.api.captureDraft();const p=h.api.saveReport();assert.equal(JSON.parse(h.pending.at(-1).args[1].body).report,'');h.pending.at(-1).resolve(reply({workflow:wf({version:1,report:''})}));await p;
 h.document.getElementById('dlgReport').value='x'.repeat(20001);h.api.captureDraft();const count=h.pending.length;await h.api.saveReport();assert.equal(h.pending.length,count);assert.equal(h.api.getDraft().text.length,20001);
});
test('directory is only shown to managers, names are safe text and failed adds retain entered name',async()=>{
 const regular=harness();load(regular);assert.equal(regular.document.getElementById('manageAssignees').hidden,true);
 const h=harness(false,true);load(h);assert.equal(h.document.getElementById('manageAssignees').hidden,false);h.document.getElementById('newAssigneeName').value='<b>Отдел</b>';
 const p=h.api.addAssignee({preventDefault(){}});h.pending.at(-1).resolve(reply({assignee:{id:3,name:'<b>Отдел</b>',archived:false,version:1}}));await p;assert.match(h.document.getElementById('activeAssignees').innerHTML,/&lt;b&gt;Отдел&lt;\/b&gt;/);
 h.document.getElementById('newAssigneeName').value='Остаться';const fail=h.api.addAssignee({preventDefault(){}});h.pending.at(-1).reject(new Error('offline'));await fail;assert.equal(h.document.getElementById('newAssigneeName').value,'Остаться');
});
test('initial saved snapshot loads without starting a portal refresh; exports use loaded period and authoritative keys',async()=>{
 const h=harness();assert.match(h.pending[0].args[0],/\/data\?/);h.pending[0].resolve(reply(data()));await flush();assert.equal(h.pending.length,1);assert.deepEqual(Array.from(h.api.getPeriod()),['2026-01-01','2026-01-02']);
 const p=h.api.exportPptx(),request=h.pending.at(-1),body=JSON.parse(request.args[1].body);assert.deepEqual(body.keys,['c_a','c_b','c_c']);assert.equal(body.rows[0]._key,'c_a');request.resolve({ok:true,headers:{get:()=>''},blob:async()=>new Blob()});await p;
});
test('missing snapshot refreshes asynchronously, retries status without resubmitting and loads persisted result',async()=>{
 const h=harness();h.pending[0].resolve(reply({error:'missing'},404));await flush();assert.equal(h.pending[1].args[0],'/mingkh/collective/api/refresh');assert.equal(h.pending[1].args[1].headers['Idempotency-Key'],'synthetic-idempotency-key');h.pending[1].resolve(reply({job_id:'abc',status_url:'/mingkh/collective/api/jobs/abc'},202));await flush();h.pending[2].reject(new Error('offline'));await flush();h.pending[3].resolve(reply({status:'running'}));await flush();h.pending[4].resolve(reply({status:'success'}));await flush();h.pending[5].resolve(reply(data()));await flush();assert.equal(h.api.getRows().length,3);assert.equal(h.pending.filter(r=>r.args[0].endsWith('/refresh')).length,1);assert.equal(h.sessions.size,0);
});
test('external job status URL is rejected and failed refresh retains previously loaded data',async()=>{
 const h=harness();h.pending[0].resolve(reply(data()));await flush();const p=h.api.refresh(true);h.pending.at(-1).resolve(reply({job_id:'evil',status_url:'https://example.invalid/jobs/a'},202));await p;assert.match(h.document.getElementById('err').textContent,/ранее загруженные/);assert.equal(h.api.getRows().length,3);assert.equal(h.pending.length,2);
});
test('request races cannot replace new snapshot or period with stale response',async()=>{
 const h=harness();h.pending[0].resolve(reply(data()));await flush();const first=h.api.refresh(false),second=h.api.refresh(false);h.pending[2].resolve(reply(data({rows:rows.slice(0,1),start:'2026-02-01',end:'2026-02-02'})));await second;h.pending[1].resolve(reply(data()));await first;assert.equal(h.api.getRows().length,1);assert.equal(h.api.getPeriod()[0],'2026-02-01');
});
test('Excel filters and aggregates include saved responsible/report while unsaved report stays out',()=>{
 const h=harness(),x=extra();x.workflows.c_a=wf({assignee_id:1,assignee_name:'Отдел',report:'Сохранено',theme:'pr'});load(h,x);h.api.state.big=true;h.api.openCard(0);h.document.getElementById('dlgReport').value='Не сохранено';h.api.captureDraft();h.api.exportAll();const wb=h.downloads[0];assert.equal(wb[0].sheet.data.length,2);assert.deepEqual(Array.from(wb[0].sheet.data[1].slice(0,3)),['Прочее','Отдел','Сохранено']);assert.equal(wb[2].name,'По ответственным');assert.equal(wb[2].sheet.data[1][1],1);
});
test('real XLSX cells treat formula-like portal text and reports as text',()=>{
 const h=harness(true),x=extra();x.workflows.c_a.report='=HYPERLINK("https://example.invalid")';h.api.load(head,rows,'Synthetic',x);h.api.state.big=true;h.api.exportAll();const wb=h.downloads[0],sheet=wb.Sheets['Обращения'];assert.equal(wb.SheetNames.length,6);assert.equal(sheet.C2.t,'s');assert.equal(sheet.C2.f,undefined);assert.equal(sheet.F2.v,101);
});
test('repeat whitespace is normalized in filters, KPI and badges',()=>{const h=harness(),r=rows.map(x=>x.slice());r[0][4]=' Да ';r[1][4]='\tда\u00a0';h.api.load(head,r,'Synthetic',extra());assert.match(h.document.getElementById('kpis').innerHTML,/Повторных<\/div><div class="v">2<\/div>/);h.api.state.rep=true;assert.equal(h.api.baseFiltered().length,2);});
test('saved-state polling adopts colleagues changes and preserves active editors and unsaved drafts',async()=>{
 const h=harness();h.pending[0].resolve(reply(data()));await flush();const p=h.api.syncSavedSnapshot(),x=extra();x.workflows.c_a=wf({version:1,report:'Доклад коллеги',assignee_id:1,assignee_name:'Отдел'});h.pending.at(-1).resolve(reply(data(x)));await p;assert.equal(h.api.getRows()[0]._workflow.report,'Доклад коллеги');
 h.api.openCard(0);const count=h.pending.length;await h.api.syncSavedSnapshot();assert.equal(h.pending.length,count);h.document.getElementById('dlgReport').value='Мой черновик';h.api.captureDraft();h.document.getElementById('dlg').close();await h.api.syncSavedSnapshot();assert.equal(h.pending.length,count);assert.equal(h.api.getDraft().text,'Мой черновик');
});
test('poll response arriving after card opens cannot replace its fields',async()=>{
 const h=harness();h.pending[0].resolve(reply(data()));await flush();const p=h.api.syncSavedSnapshot();h.api.openCard(0);const x=extra();x.workflows.c_a=wf({version:3,report:'Concurrent'});h.pending.at(-1).resolve(reply(data(x)));await p;assert.equal(h.document.getElementById('dlgReport').value,'');assert.equal(h.api.getRows()[0]._workflow.version,0);
});
test('local themes migrate only on explicit action with server version',async()=>{
 const h=harness();h.api.store.set(h.api.keyOf({'Номер обращения ЕЦУР, МСЭД':'a'}),'pr');load(h);const count=h.pending.length;assert.equal(count,1);const p=h.api.migrateLegacyThemes();assert.equal(h.pending.length,count+1);assert.equal(JSON.parse(h.pending.at(-1).args[1].body).theme,'pr');h.pending.at(-1).resolve(reply({workflow:wf({version:1,theme:'pr'})}));await p;assert.equal(h.api.getRows()[0]._theme,'pr');assert.equal(h.api.legacyThemeCandidates().length,0);
});
test('lost refresh submission repeats the idempotency key and does not abandon the pending job',async()=>{
 const h=harness();h.pending[0].resolve(reply(data()));await flush();const p=h.api.refresh(true);h.pending[1].reject(new Error('offline'));await flush();assert.equal(h.pending[1].args[1].headers['Idempotency-Key'],h.pending[2].args[1].headers['Idempotency-Key']);h.pending[2].resolve(reply({job_id:'a',status_url:'/mingkh/collective/api/jobs/a'},202));await flush();h.pending[3].resolve(reply({status:'success'}));await flush();h.pending[4].resolve(reply(data()));await p;assert.equal(h.sessions.size,0);
});
test('canonical server report clears dirty state after trimming while in-flight new edits survive',async()=>{
 const h=harness();load(h);h.api.openCard(0);h.document.getElementById('dlgReport').value='  Доклад с переносом\n';h.api.captureDraft();const p=h.api.saveReport();h.pending.at(-1).resolve(reply({workflow:wf({version:1,report:'Доклад с переносом'})}));await p;assert.equal(h.api.getDraft().dirty,false);assert.equal(h.document.getElementById('dlgReport').value,'Доклад с переносом');
 h.document.getElementById('dlgReport').value=' Повторная отправка\n';h.api.captureDraft();const retry=h.api.saveReport();h.pending.at(-1).resolve(reply({error:'conflict',workflow:wf({version:2,report:'Повторная отправка'})},409));await retry;assert.equal(h.api.getDraft().dirty,false);assert.equal(h.document.getElementById('dlgReport').value,'Повторная отправка');
});

test('saved report timestamp remains visible after reopening, invalid date is not rendered',()=>{const h=harness(),x=extra();x.workflows.c_a=wf({report:'Доклад',report_at:'2026-10-19T15:20:00+00:00'});load(h,x);h.api.openCard(0);assert.match(h.document.getElementById('dlgReportSavedAt').textContent,/19.10.2026/);assert.match(h.document.getElementById('dlgReportSavedAt').textContent,/18:20 МСК/);x.workflows.c_a.report_at='invalid';load(h,x);h.api.openCard(0);assert.equal(h.document.getElementById('dlgReportSavedAt').textContent,'');});
test('clean closed draft adopts colleague report after snapshot polling and reopens without dirty state',async()=>{const h=harness();h.pending[0].resolve(reply(data()));await flush();h.api.openCard(0);h.document.getElementById('dlg').close();const p=h.api.syncSavedSnapshot(),x=extra();x.workflows.c_a=wf({version:3,report:'Новый доклад коллеги'});h.pending.at(-1).resolve(reply(data(x)));await p;h.api.openCard(0);assert.equal(h.document.getElementById('dlgReport').value,'Новый доклад коллеги');assert.equal(h.api.getDraft().dirty,false);assert.equal(h.document.getElementById('saveReport').disabled,true);});

test('save from card opened before snapshot replacement updates current list rows with same key',async()=>{const h=harness();load(h);h.api.openCard(0);h.document.getElementById('dlgReport').value='Сохранённый после обновления';h.api.captureDraft();h.api.load(head,[...rows,rows[0]],'Synthetic',{...extra(),keys:['c_a','c_b','c_c','c_a']});const p=h.api.saveReport();h.pending.at(-1).resolve(reply({workflow:wf({version:1,report:'Сохранённый после обновления'})}));await p;assert.equal(h.api.getRows()[0]._workflow.report,'Сохранённый после обновления');assert.equal(h.api.getRows()[3]._workflow.report,'Сохранённый после обновления');assert.equal(h.api.getDraft().dirty,false);});
