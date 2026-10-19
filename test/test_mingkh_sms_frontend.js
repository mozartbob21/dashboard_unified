'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const ROOT=process.env.SMS_UI_ROOT||path.join(__dirname,'..');
const {createClient,escape,csvCell,safeLink,endpoint,requestKey}=require(path.join(ROOT,'static/mingkh-sms/common.js'));
const response=(data,status=200)=>({status,ok:status>=200&&status<300,json:async()=>data});
function storage(){const rows=new Map();return {getItem:key=>rows.get(key),setItem:(key,value)=>rows.set(key,value),removeItem:key=>rows.delete(key),rows};}
const options={method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({date:'2026-10-14'})};
const ticket={job_id:'test123',status_url:'/mingkh/sms/api/jobs/test123'};
test('only the same-origin SMS API can receive requests and external links cannot execute script',()=>{
 assert.equal(endpoint('/api/options'),'/mingkh/sms/api/options');
 for(const value of ['https://other.test/api/options','//other.test/api/options','/api/../users','/api/\\evil'])assert.throws(()=>endpoint(value));
 for(const value of ['javascript:alert(1)','data:text/html,a','https://user:secret@host.test/'])assert.equal(safeLink(value),'#');
 assert.equal(safeLink('https://rm.mosreg.ru/issues/12'),'https://rm.mosreg.ru/issues/12');
 assert.equal(escape('<a x="\'">&'),'&lt;a x=&quot;&#39;&quot;&gt;&amp;');
});
test('queued work returns original result after more than a minute and submits only once',async()=>{
 let polls=0,submits=0,waited=0;const memory=storage();
 const client=createClient({storage:memory,sleep:async ms=>{waited+=ms;},fetch:async(url,opts)=>{
  assert.equal(opts.credentials,'same-origin');assert.equal(opts.cache,'no-store');
  if(opts.method==='POST'){submits++;assert.ok(opts.headers['Idempotency-Key']);return response(ticket,202);}
  return response(++polls<=40?{status:polls===1?'queued':'running'}:{status:'success',result:{text:'Реальный результат',count:0}});
 }});
 const result=await(await client.fetch('/api/build',options)).json();
 assert.deepEqual(result,{text:'Реальный результат',count:0});assert.equal(submits,1);assert.ok(waited>60000);assert.equal(memory.rows.size,0);
});
test('lost submission retries with one idempotency key rather than starting duplicate work',async()=>{
 const keys=[];let i=0;
 const client=createClient({storage:storage(),uuid:()=> 'request-one',sleep:async()=>{},fetch:async(url,opts)=>{
  if(opts.method==='POST'){keys.push(opts.headers['Idempotency-Key']);if(!i++)throw new TypeError('offline');return response(ticket,202);}
  return response({status:'success',result:{text:'Готово'}});
 }});
 assert.equal((await(await client.fetch('/api/build',options)).json()).text,'Готово');assert.deepEqual(keys,['request-one','request-one']);
});
test('polling reconnects after network failure without another POST',async()=>{
 let submits=0,gets=0;
 const client=createClient({storage:storage(),sleep:async()=>{},fetch:async(url,opts)=>{
  if(opts.method==='POST'){submits++;return response(ticket,202);}if(++gets<3)throw new TypeError('offline');return response({status:'success',result:{text:'Данные'}});
 }});
 assert.equal((await(await client.fetch('/api/build',options)).json()).text,'Данные');assert.equal(submits,1);assert.equal(gets,3);
});
test('persistent ticket resumes the previous job after repeated network failures',async()=>{
 let submits=0;const memory=storage();
 const first=createClient({storage:memory,sleep:async()=>{},fetch:async(url,opts)=>{if(opts.method==='POST'){submits++;return response(ticket,202);}throw new TypeError('offline');}});
 await assert.rejects(first.fetch('/api/build',options),/Связь/);assert.equal(memory.rows.size,1);
 const next=createClient({storage:memory,sleep:async()=>{},fetch:async(url,opts)=>{assert.notEqual(opts.method,'POST');return response({status:'success',result:{text:'Восстановлено'}});}});
 assert.equal((await(await next.fetch('/api/build',options)).json()).text,'Восстановлено');assert.equal(submits,1);assert.equal(memory.rows.size,0);
});
test('terminal error clears job ticket and keeps a readable server error',async()=>{
 const memory=storage();const client=createClient({storage:memory,fetch:async(url,opts)=>opts.method==='POST'?response(ticket,202):response({status:'error',error:'Настройте доступ к источнику в админ-панели.'})});
 await assert.rejects(client.fetch('/api/build',options),/Настройте доступ/);assert.equal(memory.rows.size,0);
});
test('forbidden and stale saved tickets are removed, so another account can start its own job',async()=>{
 for(const status of [401,403,404]){const memory=storage();const client=createClient({storage:memory,fetch:async(url,opts)=>opts.method==='POST'?response(ticket,202):response({detail:'Недоступно'},status)});await assert.rejects(client.fetch('/api/build',options));assert.equal(memory.rows.size,0);}
});
test('HTTP errors and non-JSON errors fail initialization instead of creating fake data',async()=>{
 const denied=createClient({fetch:async()=>response({detail:'Недостаточно прав'},403)});
 await assert.rejects(denied.fetch('/api/options'),/Недостаточно прав/);
 const invalid=createClient({fetch:async()=>({ok:true,status:200,json:async()=>{throw new SyntaxError('html');}})});
 await assert.rejects(invalid.fetch('/api/options'),/неполный ответ/);
});
test('malicious status URLs cannot turn polling into an external request',async()=>{
 let count=0;const client=createClient({storage:storage(),fetch:async()=>{count++;return response({...ticket,status_url:'https://evil.test/'},202);}});
 await assert.rejects(client.fetch('/api/build',options),/адрес/);assert.equal(count,1);
});
test('all page scripts parse and templates preserve five pages and exports without CDN',()=>{
 for(const name of ['common','overview','summary','resonance','system','solutions','rez_core','rez_addr'])new vm.Script(fs.readFileSync(path.join(ROOT,'static/mingkh-sms',name+'.js'),'utf8'));
 for(const name of ['overview','summary','resonance','system','solutions'])assert.match(fs.readFileSync(path.join(ROOT,'templates','mingkh-sms-'+name+'.html'),'utf8'),/mingkh-sms-base.html/);
 const base=fs.readFileSync(path.join(ROOT,'templates/mingkh-sms-base.html'),'utf8');
 assert.match(base,/can_manage_import and not embed/);assert.match(base,/\/static\/vendor\/xlsx.full.min.js/);assert.doesNotMatch(base,/https?:\/\//);
 const embed=fs.readFileSync(path.join(ROOT,'templates/mingkh-sms-embed.html'),'utf8');assert.doesNotMatch(embed,/release-notes|module-shell|survey/);
 const rez=fs.readFileSync(path.join(ROOT,'static/mingkh-sms/resonance.js'),'utf8');assert.match(rez,/rmRows/);assert.match(rez,/registryRows/);assert.match(rez,/writeFile/);
 const sys=fs.readFileSync(path.join(ROOT,'static/mingkh-sms/system.js'),'utf8');assert.match(sys,/toCP1251/);assert.match(sys,/R.files/);
});
test('overview awaits actual business promises and cleans up frames',()=>{
 const overview=fs.readFileSync(path.join(ROOT,'static/mingkh-sms/overview.js'),'utf8');assert.match(overview,/await frame.contentWindow.SMS.ready/);assert.match(overview,/await frame.contentWindow.SMS.run/);assert.match(overview,/finally\{frame\?\.remove/);assert.match(overview,/\?embed=1/);assert.doesNotMatch(overview,/600 &&|\.click\(\)/);
});

test('CSV exports neutralize spreadsheet formulas while keeping numeric values and ordinary text',()=>{
 for(const value of ['=HYPERLINK("https://evil.test")','+CMD','-CMD','@SUM(A1:A2)','  =1+1','\tplain','\rformula','\nformula'])assert.equal(csvCell(value),"'"+value);
 assert.equal(csvCell(0),0);assert.equal(csvCell(-4),-4);assert.equal(csvCell('Город, улица; дом'),'Город, улица; дом');
 for(const name of ['resonance','system'])assert.match(fs.readFileSync(path.join(ROOT,'static/mingkh-sms',name+'.js'),'utf8'),/map\(SMS.csvCell\)\.map\(REZ.csvField\)/);
});
function browserFixture(handler){
 const node=()=>({disabled:false,value:'',className:'',textContent:'',attributes:{},setAttribute(k,v){this.attributes[k]=v;},removeAttribute(k){delete this.attributes[k];},dispatchEvent(){},classList:{contains(){return false;}}});
 const nodes={go:node(),out:node(),msg:node(),from:node(),date:node()};
 nodes.msg.classList.contains=name=>nodes.msg.className.split(' ').includes(name);
 nodes.go.onclick=()=>handler(nodes);
 const document={getElementById:id=>nodes[id]||null,querySelector:()=>null,querySelectorAll:sel=>sel.includes('.wrap')?[nodes.go,nodes.from,nodes.date]:[],createElement:()=>node()};
 const window={document,fetch:async()=>response({}),sessionStorage:storage(),crypto:{randomUUID:()=> 'uuid'}};
 const context=vm.createContext({window,URL,AbortController,Event,setTimeout,clearTimeout});
 vm.runInContext(fs.readFileSync(path.join(ROOT,'static/mingkh-sms/common.js'),'utf8'),context);
 return {SMS:window.SMS,nodes};
}
test('embedded runner awaits initialization and uses the selected year for cumulative summary',async()=>{
 let ready;const start=new Promise(resolve=>ready=resolve);
 const fixture=browserFixture(async n=>{assert.equal(n.from.value,'2025-01-01');n.out.value='Итоги 2025';n.msg.textContent='Готово';});
 fixture.SMS.initialize(async()=>{await start;fixture.nodes.from.value='2026-01-01';});fixture.SMS.bind();
 const result=fixture.SMS.run({date:'2025-09-30'});assert.equal(fixture.nodes.go.disabled,true);ready();
 assert.equal((await result).text,'Итоги 2025');assert.equal(fixture.nodes.date.value,'2025-09-30');
});
test('double activation runs one handler, and failed run cannot return the previous output',async()=>{
 let calls=0,release;const gate=new Promise(resolve=>release=resolve);
 const fixture=browserFixture(async n=>{calls++;n.go.disabled=true;await gate;n.go.disabled=false;n.msg.className='status err';n.msg.textContent='Нет доступа к источнику';});
 fixture.SMS.initialize(async()=>{});fixture.SMS.bind();await fixture.SMS.ready;
 fixture.nodes.out.value='Старые данные';const first=fixture.SMS.run({date:'2026-10-14'});await Promise.resolve();await fixture.nodes.go.onclick();release();
 await assert.rejects(first,/Нет доступа/);assert.equal(calls,1);assert.equal(fixture.nodes.out.value,'');
});

test('options warning is surfaced and system default date respects available registries',async()=>{
 const warnings=[];const client=createClient({warning:value=>warnings.push(value),fetch:async()=>response({date:'2026-10-14',system_date:'2026-10-19',warning:'Проверьте доступ к источнику'})});
 await client.fetch('/api/options');assert.deepEqual(warnings,['Проверьте доступ к источнику']);
 const system=fs.readFileSync(path.join(ROOT,'static/mingkh-sms/system.js'),'utf8');assert.match(system,/o.system_date\|\|o.date/);
 const overview=fs.readFileSync(path.join(ROOT,'static/mingkh-sms/overview.js'),'utf8');assert.match(overview,/b.id==='b3'\?\(OPT.system_date\|\|OPT.week.to\)/);
});
