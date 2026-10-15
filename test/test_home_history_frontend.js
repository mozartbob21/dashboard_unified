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
