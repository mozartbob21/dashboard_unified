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
let sortKey='resVS',asc=false,busy=false,observedRunning=false,autoAttempted=false,polling=false,lastRevision=null,loginRequired=false;
const metricValue = item => item?.value == null ? '—' : fmt(item.value)+(item.unit ? ' '+item.unit : '');
function total(key){const fields={resVS:'res_vs',sysVS:'sys_vs',tasks:'tasks_total',sysKR:'sys_kr',resKR:'res_kr',att:'att_avg'};return snap.metric_schema===1?snap.kpis?.[fields[key]]??null:null;}
function render(){
 document.getElementById('wdCoverage').textContent=(snap.table||[]).length+' ОМСУ';
 document.querySelector('.wd-snap b').textContent=stamp(snap.updated_at);
 document.getElementById('kpiGrid').innerHTML=defs.map(([id,name,url])=>{
  const d=snap.sources?.[id]||{};
  const items=d.metric_schema===1 ? d.metrics||[] : [];
  const first=items[0];
  const tables=(d.tables||[]).map(t=>'<div class="wd-source-table"><table><thead><tr>'+t.headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+t.rows.map(r=>'<tr>'+r.map(c=>'<td>'+esc(c)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>').join('');
  const rows=items.slice(1,4).map(x=>'<div class="wd-row"><span>'+esc(x.label)+'</span><b>'+esc(metricValue(x))+'</b></div>').join('');
  const extra=items.slice(4).map(x=>'<div class="wd-row"><span>'+esc(x.label)+'</span><b>'+esc(metricValue(x))+'</b></div>').join('');
  const stale=d.updated_at && Date.now()-new Date(d.updated_at).getTime()>30*60*1000;
  const label=d.metric_schema!==1?'Требует обновления':!d.updated_at?'Нет данных':!d.ok?'Прежние данные':stale?'Требует обновления':'Данные получены';
  return '<article class="wd-card s-'+(d.metric_schema===1&&d.ok&&!stale?'good':'warn')+'"><div class="wd-head"><span class="wd-src">'+esc(name)+'</span><span class="wd-status">'+label+'</span></div><p class="wd-name">'+esc(first?.label||'Показатели источника')+'</p><div class="wd-num">'+esc(metricValue(first))+'</div>'+(rows?'<div class="wd-rows">'+rows+'</div>':'')+'<p class="wd-source-error">'+esc(d.error||(!d.updated_at?'Нажмите «Обновить снимок» для загрузки данных.':''))+'</p>'+(d.data_date?'<p class="wd-hint">Дата данных источника: '+esc(d.data_date.split('-').reverse().join('.'))+'</p>':'')+'<p class="wd-hint">Получено: '+esc(stamp(d.updated_at))+'</p>'+(tables||extra?'<details><summary>Данные источника</summary>'+extra+tables+'</details>':'')+'<a href="'+url+'" target="_blank" rel="noopener noreferrer">Открыть источник ↗</a></article>';
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
async function readResponse(response){
 if(response.status===401||(response.redirected&&new URL(response.url,location.href).pathname==='/login')){
  loginRequired=true;throw Error('Сессия истекла. Войдите в Нейрону заново, чтобы проверить обновление свода.');
 }
 if(response.status===403)throw Error('Нет доступа к сводному дашборду. Обратитесь к администратору.');
 if(!response.headers.get('content-type')?.includes('application/json'))throw Error('Сервер не вернул состояние обновления. Обновите страницу.');
 return response.json();
}
async function startRefresh(){
 if(busy)return;busy=true;button.disabled=true;status.textContent='Запускаю обновление восьми источников…';
 try{const r=await fetch('/water-dashboard/run-check',{method:'POST'});const d=await readResponse(r);if(!r.ok||d.ok===false)throw Error(d.message||d.detail||'Не удалось запустить обновление');observedRunning=true;}
 catch(e){status.textContent=e.message;busy=false;button.disabled=false;}
}
button.addEventListener('click',startRefresh);
async function poll(){
 if(polling||loginRequired)return;polling=true;
 try{
  const r=await fetch('/water-dashboard/run-status',{cache:'no-store'});const state=await readResponse(r);if(!r.ok)throw Error('Не удалось проверить обновление');
  busy=!!state.running;button.disabled=busy;button.textContent=busy?'Обновление…':'⟳ Обновить снимок';
  if(busy){observedRunning=true;status.textContent='Обновление: '+(state.stage||'сбор данных');return;}
  if(lastRevision===null||lastRevision!==state.snapshot_revision){
   const response=await fetch('/water-dashboard/snapshot',{cache:'no-store'});const latest=await readResponse(response);if(!response.ok)throw Error('Не удалось получить снимок');
   if(latest.checked_at!==snap.checked_at){snap=latest;render();}
   lastRevision=state.snapshot_revision??'';
  }
  status.textContent=state.last_error?'Ошибка обновления: '+state.last_error:snap.checked_at?'Обновлено '+Object.values(snap.sources||{}).filter(s=>s.metric_schema===1&&s.ok).length+' из 8 источников. Последняя проверка: '+stamp(snap.checked_at):'Данные ещё не загружены';
  // One refresh per page visit when the last attempt is older than 30 minutes.
  const elapsed=Date.now()-new Date(snap.checked_at||0).getTime();
  if(!autoAttempted&&!observedRunning){autoAttempted=true;if(snap.metric_schema!==1||!snap.checked_at||elapsed>30*60*1000)await startRefresh();}
 }catch(e){status.textContent=e.message;}finally{polling=false;}
}
render();poll();setInterval(poll,5000);
})();
