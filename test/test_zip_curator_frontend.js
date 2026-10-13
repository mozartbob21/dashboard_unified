// Run with: node --test test/test_zip_curator_frontend.js
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(process.env.ZIP_CURATOR_TEMPLATE || path.join(__dirname, '../templates/zip_curator_original.html'), 'utf8');
const refreshCode = source.slice(source.indexOf('let driveSyncRunning=false;'), source.indexOf('// сравнение дат'));

class Element {
  constructor(tag='div') {this.tagName=tag;this.children=[];this.dataset={};this.attributes={};this.hidden=false;this.disabled=false;this._text='';}
  set textContent(value) {this._text=String(value);this.children=[];}
  get textContent() {return this._text+this.children.map(child=>child.textContent).join(' ');}
  set innerHTML(value) {throw new Error('Refresh diagnostics must not be interpreted as HTML');}
  append(...items) {this.children.push(...items);}
  replaceChildren(...items) {this._text='';this.children=items;}
  setAttribute(name,value) {this.attributes[name]=String(value);}
  removeAttribute(name) {delete this.attributes[name];}
}
function deferred() {let resolve,reject;const promise=new Promise((yes,no)=>{resolve=yes;reject=no;});return {promise,resolve,reject};}
function fixture() {
  const button=new Element('button');button.textContent='⟳ Обновить данные (из папки A)';
  const panel=new Element();panel.hidden=true;
  const requests=[],states=[],notices=[];
  const nodes={'#btnDrive':button,'#driveSyncResult':panel};
  const context=vm.createContext({YANDEX_FOLDER_A_URL:'https://disk.yandex.ru/d/test',
    $:selector=>nodes[selector],document:{createElement:tag=>new Element(tag)},
    applyServerState:state=>states.push(state),toast:text=>notices.push(text),
    fetch:(url,options)=>{const pending=deferred();requests.push({url,options,...pending});return pending.promise;}});
  vm.runInContext(refreshCode,context);
  return {context,button,panel,requests,states,notices};
}
const state={pending:[{rso:'Новый реестр'}],clean:{saved:{items:[{qty:5}]}}};
const response=(body,status=200)=>({ok:status>=200&&status<300,status,json:async()=>body});
const result={ok:true,partial:false,files:5,downloaded:2,reused:3,removed:1,added:2,skipped:3,failed:0,errors:[],state};

 test('the curator embedded scripts compile',()=>{
  for(const match of source.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))new vm.Script(match[1]);
});

test('double clicks share one refresh and the button stays disabled until the response is applied',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  assert.equal(f.button.disabled,true);assert.equal(f.panel.dataset.kind,'running');
  await f.context.fetchFromDrive();assert.equal(f.requests.length,1);
  assert.equal(f.requests[0].url,'/zip_curator/api/scan');
  assert.equal(f.requests[0].options.method,'POST');
  assert.equal(f.requests[0].options.signal,undefined,'a browser timeout must not abandon work still running on the server');
  f.requests[0].resolve(response(result));await pending;
  assert.equal(f.states.length,1);assert.equal(f.states[0],state);
  assert.equal(f.button.disabled,false);assert.equal(f.button.textContent,'⟳ Обновить данные (из папки A)');
  assert.equal(f.button.attributes['aria-busy'],undefined);
  assert.equal(f.panel.dataset.kind,'success');assert.equal(f.panel.hidden,false);
  assert.match(f.panel.textContent,/Скачано: 2/);assert.match(f.panel.textContent,/Без изменений: 3/);
});

test('a partial refresh still applies successful files and identifies retained old versions',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  const name='<img src=x onerror=alert(1)>'+('длинное_название_'.repeat(30))+'.xlsx';
  f.requests[0].resolve(response({...result,partial:true,failed:1,errors:[{name,message:'Яндекс Диск временно недоступен.',retained:true}]}));
  await pending;
  assert.equal(f.states.length,1);assert.equal(f.panel.dataset.kind,'partial');
  assert.match(f.panel.textContent,/Обновление завершено частично/);
  assert.ok(f.panel.textContent.includes(name),'the complete filename remains available as text');
  assert.match(f.panel.textContent,/предыдущая локальная версия/);
  assert.match(f.panel.textContent,/не является новым обновлением/);
  const details=f.panel.children.find(child=>child.tagName==='details');
  assert.ok(details);assert.equal(details.children[0].tagName,'summary');
  assert.equal(f.button.disabled,false);
});

test('all files failing never appears as a successful update',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  f.requests[0].resolve(response({...result,files:2,downloaded:0,reused:0,added:0,failed:2,partial:true,errors:[{name:'1.xlsx',message:'Ошибка загрузки',retained:false},{name:'2.xlsx',message:'Ошибка загрузки',retained:true}]}));
  await pending;
  assert.equal(f.panel.dataset.kind,'partial');assert.match(f.panel.textContent,/Не удалось обновить файлы/);
  assert.equal(f.notices.includes('Папка А обновлена'),false);
  assert.match(f.panel.textContent,/Новая версия не загружена/);
  assert.equal(f.states.length,1,'the usable state returned by the server is still applied');
});

test('legacy signed download URLs are hidden in both the durable panel and toast',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  f.requests[0].resolve(response({ok:false,error:'500 Server Error: https://downloader.disk.yandex.ru/disk/secret?token=secret'},502));
  await pending;
  assert.equal(f.states.length,0);assert.equal(f.panel.dataset.kind,'error');
  assert.doesNotMatch(f.panel.textContent,/https:|secret|downloader/);
  assert.doesNotMatch(f.notices.join(' '),/https:|secret|downloader/);
  assert.equal(f.button.disabled,false);
});

test('a lost response preserves the queue, restores the button and permits a later retry',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  f.requests[0].reject(new Error('Network error https://example.invalid/private?signature=secret'));
  await pending;
  assert.equal(f.states.length,0);assert.equal(f.panel.dataset.kind,'error');
  assert.match(f.panel.textContent,/Результат обновления пока не подтверждён/);
  assert.doesNotMatch(f.panel.textContent,/signature|private/);assert.equal(f.button.disabled,false);
  const retry=f.context.fetchFromDrive();assert.equal(f.requests.length,2);
  f.requests[1].resolve(response(result));await retry;
  assert.equal(f.panel.dataset.kind,'success');assert.doesNotMatch(f.panel.textContent,/прервалась/);
});

test('refreshes without a complete state cannot clear the current queue',async()=>{
  for(const incomplete of [undefined,{}, {pending:[]}]){
    const f=fixture();const pending=f.context.fetchFromDrive();
    f.requests[0].resolve(response({...result,state:incomplete}));await pending;
    assert.equal(f.states.length,0);
  }
});

test('a server-side conflict gives a durable explanation and does not replace data',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  f.requests[0].resolve(response({ok:false,error:'Обновление папки А уже выполняется. Повторите проверку позже.'},409));await pending;
  assert.match(f.panel.textContent,/уже выполняется/);assert.equal(f.states.length,0);assert.equal(f.button.disabled,false);
});

test('a redirected login page or malformed JSON cannot masquerade as an empty successful update',async()=>{
  const f=fixture();const pending=f.context.fetchFromDrive();
  f.requests[0].resolve({ok:true,status:200,json:async()=>{throw new Error('Unexpected HTML login page');}});await pending;
  assert.equal(f.panel.dataset.kind,'error');assert.equal(f.states.length,0);assert.equal(f.button.disabled,false);
});
