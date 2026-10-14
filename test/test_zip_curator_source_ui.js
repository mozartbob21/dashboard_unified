'use strict';
const assert=require('node:assert/strict'),test=require('node:test'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const html=fs.readFileSync(process.env.ZIP_CURATOR_TEMPLATE||path.join(__dirname,'..','templates','zip_curator_original.html'),'utf8');
const dateCode=html.slice(html.indexOf('function sourceDateText('),html.indexOf('\nconst fmt='));
const scanCode=html.slice(html.indexOf('let driveSyncRunning=false;'),html.indexOf('// сравнение дат вида'));
class Element{
  constructor(tag='div'){this.tagName=tag;this.children=[];this.dataset={};this.attributes={};this._text='';this.disabled=false;this.hidden=false;}
  set textContent(value){this._text=String(value);this.children=[];}
  get textContent(){return this._text+this.children.map(item=>item.textContent).join(' ');}
  set innerHTML(_){throw new Error('Source diagnostics must be safe text');}
  append(...items){this.children.push(...items);}
  replaceChildren(...items){this._text='';this.children=items;}
  setAttribute(name,value){this.attributes[name]=String(value);}
  removeAttribute(name){delete this.attributes[name];}
}
const response=(value,status=200)=>({ok:status>=200&&status<300,status,json:async()=>value});
function fixture(){
  const nodes=Object.fromEntries(['sourceBadge','cfgWarn','btnDrive','driveSyncResult'].map(id=>['#'+id,new Element()]));
  nodes['#btnDrive'].textContent='⟳ Обновить данные';
  const calls=[],states=[];
  let source={provider:'combined',label:'Яндекс Диск + Облако МО',configured:true,sources:[{provider:'yandex',label:'Яндекс Диск',configured:true},{provider:'mosreg',label:'Облако МО',configured:true}]};
  let next=response({ok:true,files:2,downloaded:1,duplicates:1,state:{pending:[],clean:{}}});
  const context=vm.createContext({$:selector=>nodes[selector],document:{createElement:tag=>new Element(tag)},
    fetch:async(url,options)=>{calls.push([url,options]);return url.endsWith('/source')?response(source):await next;},
    toast:()=>{},applyServerState:state=>states.push(state),
  });
  vm.runInContext(dateCode+'\n'+scanCode,context);
  return {context,nodes,calls,states,setSource:value=>{source=value;},setResponse:value=>{next=value;},panel:nodes['#driveSyncResult'],button:nodes['#btnDrive']};
}
function format(f,record){f.context.record=record;return vm.runInContext('sourceDateText(record)',f.context);}

test('source dates preserve creation/modification semantics and reject invalid or unrelated dates',()=>{
  const f=fixture();
  assert.match(format(f,{source_created_at:'2025-02-03T08:15:00Z',source_modified_at:'2026-10-13T10:00:00Z',loaded_at:'2026-10-14T11:00:00Z'}),/^Создан в источнике:.*2025/);
  assert.match(format(f,{source_modified_at:'2026-10-13T10:00:00Z'}),/^Изменён в источнике:/);
  assert.equal(format(f,{uploaded:1700000000000,loaded_at:'2026-10-14T11:00:00Z'}),'Дата источника неизвестна');
  assert.equal(format(f,{source_created_at:'invalid'}),'Дата источника неизвестна');
  assert.equal(format(f,{source_created_at:'2026-02-30T12:00:00Z'}),'Дата источника неизвестна');
  assert.doesNotMatch(html,/YANDEX_FOLDER_A_URL/);
  assert.match(html,/<th>Дата файла в источнике<\/th>/);
});

test('both configured sources remain visible and the primary update sends no provider override',async()=>{
  const f=fixture();await f.context.loadSource();
  assert.equal(f.nodes['#sourceBadge'].textContent,'Источники: Яндекс Диск + Облако МО');
  await f.context.fetchFromDrive();
  const [url,options]=f.calls.find(([url])=>url.endsWith('/scan'));
  assert.equal(url,'/zip_curator/api/scan');assert.equal(options.method,'POST');assert.deepEqual(JSON.parse(options.body),{});
  assert.equal(f.button.disabled,false);assert.equal(f.panel.dataset.kind,'success');
  assert.match(f.panel.textContent,/Повторов исключено: 1/);assert.equal(f.states.length,1);
});

test('Yandex stays available when cloud has not been configured',async()=>{
  const f=fixture();f.setSource({provider:'combined',configured:true,sources:[{provider:'yandex',configured:true},{provider:'mosreg',configured:false}]});
  await f.context.loadSource();
  assert.equal(f.nodes['#sourceBadge'].textContent,'Источник: Яндекс Диск');assert.equal(f.nodes['#cfgWarn'].textContent,'');
  await f.context.fetchFromDrive();assert.equal(f.states.length,1);
});

test('a failed cloud source does not hide successful Yandex data and has its own safe diagnostic',async()=>{
  const f=fixture();f.setResponse(response({ok:true,partial:true,files:4,downloaded:2,duplicates:1,failed:0,
    sources:[{provider:'yandex',ok:true,files:4},{provider:'mosreg',ok:false,error:'Не удалось открыть папку.'}],
    state:{pending:[{rso:'Получено из Яндекса'}],clean:{}}}));
  await f.context.fetchFromDrive();
  assert.equal(f.panel.dataset.kind,'partial');assert.equal(f.states.length,1);assert.equal(f.states[0].pending[0].rso,'Получено из Яндекса');
  assert.match(f.panel.textContent,/Яндекс Диск.*Проверен.*Файлов: 4/);
  assert.match(f.panel.textContent,/Облако МО.*Не обновлён.*Не удалось открыть папку/);
  assert.match(f.panel.textContent,/Повторов исключено: 1/);
  const list=f.panel.children.find(item=>item.attributes.role==='list');assert.ok(list);assert.equal(list.children.length,2);
  assert.equal(f.calls.filter(([url])=>url.endsWith('/scan')).length,1,'no separate silent retry is made');
});

test('any failed source marks a partial outcome even without a top-level partial flag',async()=>{
  const f=fixture();f.setResponse(response({ok:true,files:1,sources:[{provider:'yandex',ok:false,error:'Недоступен'},{provider:'mosreg',ok:true,files:1}],state:{pending:[],clean:{}}}));
  await f.context.fetchFromDrive();assert.equal(f.panel.dataset.kind,'partial');assert.equal(f.states.length,1);
});

test('a complete failure reports both sources, sanitizes transport secrets and preserves existing data',async()=>{
  const f=fixture();f.setResponse(response({ok:false,error:'Оба источника временно недоступны.',sources:[
    {provider:'yandex',ok:false,error:'https://downloader.invalid/?token=secret'},
    {provider:'mosreg',ok:false,error:'<img src=x onerror=alert(1)>'},
  ]},502));
  await f.context.fetchFromDrive();assert.equal(f.states.length,0);assert.equal(f.panel.dataset.kind,'error');assert.equal(f.button.disabled,false);
  assert.match(f.panel.textContent,/Яндекс Диск/);assert.match(f.panel.textContent,/Облако МО/);
  assert.doesNotMatch(f.panel.textContent,/token|secret|downloader/);
  assert.match(f.panel.textContent,/<img src=x onerror=alert\(1\)>/,'HTML-like diagnostic remains inert text');
});

test('metadata failure does not prevent updating or expose signed links',async()=>{
  const f=fixture();f.setSource({provider:'combined',configured:true,sources:[{provider:'yandex',configured:true}],warning:'https://cloud.invalid/?token=secret'});
  await f.context.loadSource();assert.doesNotMatch(f.nodes['#cfgWarn'].textContent,/secret|https:/);assert.equal(f.button.disabled,false);
  f.setSource({configured:false});await f.context.loadSource();assert.match(f.nodes['#cfgWarn'].textContent,/вручную/);
  await f.context.fetchFromDrive();assert.equal(f.states.length,1);
});

test('double clicks start one combined update and a later retry is safe after losing the response',async()=>{
  const f=fixture();let reject;f.setResponse(new Promise((_,no)=>{reject=no;}));
  const pending=f.context.fetchFromDrive();assert.equal(f.button.disabled,true);await f.context.fetchFromDrive();assert.equal(f.calls.length,1);
  reject(new Error('Network disconnected'));await pending;assert.equal(f.button.disabled,false);assert.equal(f.states.length,0);
  f.setResponse(response({ok:true,files:1,state:{pending:[],clean:{}}}));await f.context.fetchFromDrive();
  assert.equal(f.calls.length,2);assert.equal(f.states.length,1);assert.equal(f.panel.dataset.kind,'success');
});
