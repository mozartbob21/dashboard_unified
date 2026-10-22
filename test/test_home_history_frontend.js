const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const home = fs.readFileSync(require.resolve('../templates/home.html'),'utf8');
const script = [...home.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)].map(x=>x[1]).find(x=>x.includes('const listEl = document.getElementById("historyList")'));

async function render(entries) {
  const list = {innerHTML:''}; const calls=[];
  const context = {document:{hidden:false,getElementById:id=>id==='historyList'?list:null},
    fetch:async url=>{ calls.push(url); return {ok:true,json:async()=>entries}; },
    setInterval:()=>0, console:{warn:()=>{}}, Date};
  vm.runInNewContext(script,context);
  await new Promise(resolve=>setImmediate(resolve));
  return {html:list.innerHTML,calls};
}

test('history uses current Russian server labels and icons for every job, including the water map',async()=>{
  const data=[
    {module_id:'mingkh-water-map',module_label:'МИНЖКХ — карта жалоб',module_icon:'🗺️'},
    {module_id:'water_dashboard',module_label:'Качество водоснабжения',module_icon:'🌊'},
    {module_id:'camera_prescriptions',module_label:'Предписания по камерам',module_icon:'📄'},
    {module_id:'edds',module_label:'ЕДДС — контроль заявок',module_icon:'🚨'}
  ].map(entry=>({...entry,user:'Автозапуск',status:'success',started_at:new Date().toISOString()}));
  const {html,calls}=await render(data);
  for(const entry of data){assert.ok(html.includes(entry.module_label));assert.ok(html.includes(entry.module_icon));assert.ok(!html.includes(entry.module_id));}
  assert.deepEqual(calls,['/api/history/recent']);
});

test('unknown history entries have Russian fallback and unsafe server text is escaped',async()=>{
  const {html}=await render([
    {module_id:'untranslated-job',status:'untranslated-status'},
    {module_id:'x',module_label:'<img src=x onerror=alert(1)>',module_icon:'<svg/onload=alert(1)>',user:'<script>evil</script>',status:'" onclick="evil()'}
  ]);
  assert.match(html,/Модуль платформы/);
  assert.doesNotMatch(html,/untranslated-|onclick=|<img|<svg|<script>/);
  assert.match(html,/&lt;img/);assert.match(html,/&lt;svg/);assert.match(html,/&lt;script/);
});

test('the home header has no notification bell or polling and retains profile release history',()=>{
  assert.doesNotMatch(home,/notificationsBell|notificationsDropdown|api\/notifications|neurona-notifications-read/);
  assert.match(home,/data-release-notes-open/);assert.match(home,/Версия 3/);
  assert.match(home,/id="profileMenuButton"/);assert.match(home,/id="themeToggle"/);
});


const tick = () => new Promise(resolve => setImmediate(resolve));
function historyWidget({storage = new Map(), request} = {}) {
  const list = {innerHTML: ''};
  const body = {hidden: false};
  const classes = new Set();
  const widget = {classList: {toggle(name, enabled) { enabled ? classes.add(name) : classes.delete(name); }}};
  const label = {textContent: ''};
  const attrs = new Map();
  const events = {};
  const button = {
    setAttribute: (name, value) => attrs.set(name, value),
    querySelector: () => label,
    addEventListener: (event, listener) => { events[event] = listener; },
  };
  const elements = {historyList: list, historyBody: body, historyWidget: widget, historyCollapseBtn: button};
  const requests = [];
  let poll;
  const context = {
    document: {hidden: false, getElementById: id => elements[id] || null},
    localStorage: {getItem: key => storage.get(key), setItem: (key, value) => storage.set(key, value)},
    fetch: async url => {
      requests.push(url);
      return request ? request() : {ok: true, json: async () => []};
    },
    setInterval: listener => { poll = listener; }, console: {warn() {}}, Date,
  };
  vm.runInNewContext(script, context);
  return {list, body, classes, label, attrs, requests, click: () => events.click(), poll: () => poll(), context};
}

test('collapse hides history, updates accessible state and pauses polling; expand refreshes', async () => {
  const ui = historyWidget();
  await tick();
  assert.equal(ui.requests.length, 1);
  assert.equal(ui.attrs.get('aria-expanded'), 'true');
  ui.click();
  assert.equal(ui.body.hidden, true);
  assert.equal(ui.classes.has('is-collapsed'), true);
  assert.equal(ui.attrs.get('aria-expanded'), 'false');
  assert.equal(ui.attrs.get('aria-label'), 'Развернуть историю запусков');
  assert.equal(ui.label.textContent, 'Развернуть');
  await ui.poll();
  assert.equal(ui.requests.length, 1);
  ui.click();
  await tick();
  assert.equal(ui.requests.length, 2);
  assert.equal(ui.body.hidden, false);
  assert.equal(ui.attrs.get('aria-expanded'), 'true');
  assert.equal(ui.label.textContent, 'Свернуть');
});

test('collapse choice survives page reload and opening is saved too', async () => {
  const storage = new Map();
  const first = historyWidget({storage});
  first.click();
  const second = historyWidget({storage});
  await tick();
  assert.equal(second.body.hidden, true);
  assert.equal(second.requests.length, 0);
  second.click();
  const third = historyWidget({storage});
  await tick();
  assert.equal(third.body.hidden, false);
  assert.equal(third.requests.length, 1);
});

test('blocked browser storage does not disable history or the collapse control', async () => {
  const storage = {get() {throw Error('blocked');}, set() {throw Error('blocked');}};
  const ui = historyWidget({storage});
  await tick();
  ui.click();
  ui.click();
  await tick();
  assert.equal(ui.body.hidden, false);
  assert.equal(ui.requests.length, 2);
});

test('an in-flight response cannot reopen collapsed history or duplicate requests', async () => {
  let finish;
  const pending = new Promise(resolve => { finish = resolve; });
  const ui = historyWidget({request: () => pending});
  ui.click();
  ui.click();
  assert.equal(ui.requests.length, 1);
  ui.click();
  finish({ok: true, json: async () => [{module_label: 'Проверка', status: 'success'}]});
  await tick();
  assert.equal(ui.body.hidden, true);
  assert.equal(ui.attrs.get('aria-expanded'), 'false');
  assert.match(ui.list.innerHTML, /Проверка/);
});
