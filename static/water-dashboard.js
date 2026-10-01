(() => {
'use strict';
const esc = v => String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = v => v == null ? '—' : typeof v === 'number' ? v.toLocaleString('ru-RU') : String(v);
const stamp = v => v ? String(v).replace('T',' ') : 'ещё не получены';
const defs = [
 ['valves','Замена задвижек (ZULUGIS)','https://datalens.yandex/e2q0obgt7xsex'],
 ['flush','Промывки сетей и РЧВ','https://datalens.yandex/j9dqqujx03qa3'],
 ['tasks','Просроченные задачи ОМСУ','https://datalens.yandex/hdxgldxnx8ui1'],
 ['sys_vs','Системные и резонансные адреса (ВС)','https://datalens.yandex/6k9dbjyurmu0q'],
 ['edo_rso','Переход РСО на ЭДО','https://datalens.yandex/f5wqqij889haz'],
 ['nvos','Плата за НВОС','https://datalens.yandex/62l1qih3msocq?tab=KD'],
 ['meetings','Присутствие на совещаниях','https://datalens.yandex/5f1g88ft7v1up'],
 ['sys_kr','Системные и резонансные адреса (КР)','https://datalens.yandex/6rxd41nckzkep']
];
let snap = JSON.parse(document.getElementById('waterSnapshot').textContent || '{}');
let sortKey='resVS',asc=false,busy=false,observedRunning=false,autoAttempted=false,polling=false,lastRevision=null;
const fields={tasks:[['tasks','Просроченные задачи']],sys_vs:[['sysVS','Системные адреса'],['resVS','Резонансные адреса']],sys_kr:[['sysKR','Системные адреса'],['resKR','Резонансные адреса']],meetings:[['att','Средняя явка, %']]};
function total(key){const rows=(snap.table||[]).filter(r=>r[key]!=null);if(!rows.length)return null;const n=rows.reduce((a,r)=>a+Number(r[key]),0);return key==='att'?Math.round(n/rows.length):n;}
function render(){
 document.getElementById('wdCoverage').textContent=(snap.table||[]).length+' ОМСУ';
 document.querySelector('.wd-snap b').textContent=stamp(snap.updated_at);
 document.getElementById('kpiGrid').innerHTML=defs.map(([id,name,url])=>{
  const d=snap.sources?.[id]||{};
  const metrics=(d.updated_at?fields[id]||[]:[]).map(([key,label])=>({label,value:total(key)})).filter(x=>x.value!=null);
  const items=metrics.length?metrics:(d.widgets||[]);
  const first=items[0];
  const tables=(d.tables||[]).map(t=>'<div class="wd-source-table"><table><thead><tr>'+t.headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+t.rows.map(r=>'<tr>'+r.map(c=>'<td>'+esc(c)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>').join('');
  const rows=items.slice(1,4).map(x=>'<div class="wd-row"><span>'+esc(x.label)+'</span><b>'+esc(fmt(x.value))+'</b></div>').join('');
  const extra=items.slice(4).map(x=>'<div class="wd-row"><span>'+esc(x.label)+'</span><b>'+esc(fmt(x.value))+'</b></div>').join('');
  const stale=d.updated_at && Date.now()-new Date(d.updated_at).getTime()>30*60*1000;
  const label=!d.updated_at?'Нет данных':!d.ok?'Прежние данные':stale?'Требует обновления':'Данные получены';
  return '<article class="wd-card s-'+(d.ok&&!stale?'good':'warn')+'"><div class="wd-head"><span class="wd-src">'+esc(name)+'</span><span class="wd-status">'+label+'</span></div><p class="wd-name">'+esc(first?.label||'Показатели источника')+'</p><div class="wd-num">'+esc(first?fmt(first.value):'—')+'</div>'+(rows?'<div class="wd-rows">'+rows+'</div>':'')+'<p class="wd-source-error">'+esc(d.error||(!d.updated_at?'Нажмите «Обновить снимок» для загрузки данных.':''))+'</p><p class="wd-hint">Получено: '+esc(stamp(d.updated_at))+'</p>'+(tables||extra?'<details><summary>Данные источника</summary>'+extra+tables+'</details>':'')+'<a href="'+url+'" target="_blank" rel="noopener noreferrer">Открыть источник ↗</a></article>';
 }).join('');
 document.getElementById('dynGrid').innerHTML=defs.map(([id,name])=>{const d=snap.sources?.[id]||{};return '<div class="wd-dcard"><p class="lbl">'+esc(name)+'</p><p>'+esc(stamp(d.updated_at))+'</p><p class="note">'+esc(d.error||d.refresh||'Период обновления источника не указан')+'</p></div>';}).join('');
 renderTable();
}
function renderTable(){
 const keys=['resVS','sysVS','tasks','sysKR','resKR','att'];
 const rows=[...(snap.table||[])].sort((a,b)=>{if(a[sortKey]==null)return 1;if(b[sortKey]==null)return -1;return (asc?1:-1)*(sortKey==='name'?String(a.name).localeCompare(b.name,'ru'):a[sortKey]-b[sortKey]);});
 document.querySelector('#omsu tbody').innerHTML=rows.map(r=>'<tr><td>'+esc(r.name)+'</td>'+keys.map(k=>'<td>'+esc(fmt(r[k]))+'</td>').join('')+'</tr>').join('')||'<tr><td colspan="7">Данные ещё не получены из источников.</td></tr>';
 keys.forEach(k=>{document.getElementById('sum'+k[0].toUpperCase()+k.slice(1)).textContent=fmt(total(k));});
 document.querySelectorAll('#omsu th[data-k]').forEach(h=>{h.classList.toggle('sorted',h.dataset.k===sortKey);h.querySelector('.arw').textContent=asc?'▲':'▼';});
}
document.querySelectorAll('#omsu th[data-k]').forEach(h=>h.addEventListener('click',()=>{asc=h.dataset.k===sortKey?!asc:h.dataset.k==='name';sortKey=h.dataset.k;renderTable();}));
const button=document.getElementById('refreshSnap'),status=document.getElementById('wdStatus');
async function startRefresh(){
 if(busy)return;busy=true;button.disabled=true;status.textContent='Запускаю обновление восьми источников…';
 try{const r=await fetch('/water-dashboard/run-check',{method:'POST'});const d=await r.json();if(!r.ok||d.ok===false)throw Error(d.message||d.detail||'Не удалось запустить обновление');observedRunning=true;}
 catch(e){status.textContent=e.message;busy=false;button.disabled=false;}
}
button.addEventListener('click',startRefresh);
async function poll(){
 if(polling)return;polling=true;
 try{
  const r=await fetch('/water-dashboard/run-status',{cache:'no-store'});if(!r.ok)throw Error('Не удалось проверить обновление');const state=await r.json();
  busy=!!state.running;button.disabled=busy;button.textContent=busy?'Обновление…':'⟳ Обновить снимок';
  if(busy){observedRunning=true;status.textContent='Обновление: '+(state.stage||'сбор данных');return;}
  if(lastRevision===null||lastRevision!==state.snapshot_revision){
   const response=await fetch('/water-dashboard/snapshot',{cache:'no-store'});if(!response.ok)throw Error('Не удалось получить снимок');const latest=await response.json();
   if(latest.checked_at!==snap.checked_at){snap=latest;render();}
   lastRevision=state.snapshot_revision??'';
  }
  status.textContent=state.last_error?'Ошибка обновления: '+state.last_error:snap.checked_at?'Последняя проверка: '+stamp(snap.checked_at):'Данные ещё не загружены';
  // One refresh per page visit when the last attempt is older than 30 minutes.
  const elapsed=Date.now()-new Date(snap.checked_at||0).getTime();
  if(!autoAttempted&&!observedRunning){autoAttempted=true;if(!snap.checked_at||elapsed>30*60*1000)await startRefresh();}
 }catch(e){status.textContent=e.message;}finally{polling=false;}
}
render();poll();setInterval(poll,5000);
})();
