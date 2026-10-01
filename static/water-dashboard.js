(() => {
'use strict';
const esc = v => String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = v => v == null ? '—' : typeof v === 'number' ? v.toLocaleString('ru-RU') : String(v);
const stamp = v => v ? String(v).replace('T',' ') : 'ещё не получены';
const defs = [
 ['valves','Замена задвижек (ZULUGIS)','https://datalens.yandex/e2q0obgt7xsex'],
 ['tasks','Задачи РМ МИНЖКХ','https://datalens.yandex/hdxgldxnx8ui1'],
 ['sys_vs','Водоснабжение','https://datalens.yandex/6k9dbjyurmu0q'],
 ['edo_rso','Переход РСО на ЭДО','https://datalens.yandex/f5wqqij889haz'],
 ['nvos','Плата за НВОС','https://datalens.yandex/62l1qih3msocq?tab=KD'],
 ['meetings','Дисциплина совещаний','https://datalens.yandex/5f1g88ft7v1up'],
 ['sys_kr','Капитальный ремонт','https://datalens.yandex/6rxd41nckzkep']
];
let snap = JSON.parse(document.getElementById('waterSnapshot').textContent || '{}');
let sortKey='resVS',asc=false,busy=false,observedRunning=false,autoAttempted=false,polling=false,lastRevision=null,loginRequired=false,submitting=false,refreshError='';
const metricValue = item => item?.value == null ? '—' : fmt(item.value)+(item.unit ? ' '+item.unit : '');
function total(key){const fields={resVS:'res_vs',sysVS:'sys_vs',tasks:'tasks_total',sysKR:'sys_kr',resKR:'res_kr',att:'att_avg'};return snap.metric_schema===1?snap.kpis?.[fields[key]]??null:null;}
const drawer=document.getElementById('sourceDetails');
let openedSource=null,detailTrigger=null;
function sourceState(d){
 const stale=d.updated_at && Date.now()-new Date(d.updated_at).getTime()>30*60*1000;
 return {good:d.metric_schema===1&&d.ok&&!stale,label:!d.updated_at?'Нет данных':!d.ok?'Прежние данные':stale?'Требует обновления':'Данные получены'};
}
const frequency = d => d.refresh || 'Частота не указана источником';
const metricRows = items => items.map(x=>'<div class="wd-row"><span>'+esc(x.label)+'</span><b>'+esc(metricValue(x))+'</b></div>').join('');
function render(){
 document.getElementById('wdCoverage').textContent=(snap.table||[]).length+' ОМСУ';
 document.querySelector('.wd-snap b').textContent=stamp(snap.updated_at);
 document.getElementById('kpiGrid').innerHTML=defs.map(([id,name,url])=>{
  const d=snap.sources?.[id]||{},state=sourceState(d);
  const items=d.metric_schema===1 ? d.metrics||[] : [], first=items[0];
  return '<article class="wd-card s-'+(state.good?'good':'warn')+'" data-source="'+id+'">'+
   '<div class="wd-head"><h3 class="wd-src">'+esc(name)+'</h3><button class="wd-refresh-card" type="button" data-refresh="'+id+'" aria-label="Обновить: '+esc(name)+'" title="Обновить только этот источник" '+(busy?'disabled':'')+'>↻</button></div>'+
   '<span class="wd-status '+(state.good?'good':'warn')+'">'+state.label+'</span><p class="wd-name">'+esc(first?.label||'Показатели источника')+'</p><div class="wd-num">'+esc(metricValue(first))+'</div>'+
   '<div class="wd-rows">'+metricRows(items.slice(1,4))+'</div><div class="wd-card-meta">'+
   (d.error?'<p class="wd-card-note">Последняя проверка не завершена. Подробности — в данных источника.</p>':'')+
   (d.data_date?'<p class="wd-hint">Данные на '+esc(d.data_date.split('-').reverse().join('.'))+'</p>':'')+
   '<p class="wd-hint">Получено: '+esc(stamp(d.updated_at))+'</p></div>'+
   '<div class="wd-card-actions"><button class="wd-details-button" type="button" data-details="'+id+'" aria-haspopup="dialog" aria-controls="sourceDetails">Данные источника <span aria-hidden="true">↗</span></button><a class="wd-source-link" aria-label="Открыть источник: '+esc(name)+'" href="'+url+'" target="_blank" rel="noopener noreferrer">Источник ↗</a></div>'+
   '<p class="wd-frequency" title="Частота по странице первоисточника. Проверено: '+esc(stamp(d.refresh_checked_at))+'">'+esc(frequency(d))+'</p></article>';
 }).join('');
 document.getElementById('dynGrid').innerHTML=defs.map(([id,name])=>{const d=snap.sources?.[id]||{};return '<div class="wd-dcard"><p class="lbl">'+esc(name)+'</p><p>'+esc(stamp(d.updated_at))+'</p><p class="note">'+esc(d.error||frequency(d))+'</p></div>';}).join('');
 renderTable();
 if(drawer.open&&openedSource)renderDetails(openedSource);
}
function renderDetails(id){
 const def=defs.find(x=>x[0]===id);if(!def)return;
 const [,name,url]=def,d=snap.sources?.[id]||{},detail=d.details||{};
 document.getElementById('sourceDetailsTitle').textContent=name;
 let html='<div class="wd-detail-meta"><span>'+esc(sourceState(d).label)+'</span><span>Получено: '+esc(stamp(d.updated_at))+'</span>'+(d.data_date?'<span>Данные на '+esc(d.data_date.split('-').reverse().join('.'))+'</span>':'')+'</div>';
 if(d.error)html+='<p class="wd-detail-note">'+esc(d.error)+'</p>';
 if(detail.basis)html+='<p class="wd-detail-basis">'+esc(detail.basis)+'</p>';
 if(detail.coverage?.note)html+='<p class="wd-detail-note">'+esc(detail.coverage.note)+'</p>';
 const groups=detail.groups||[];
 for(const group of groups){
  html+='<section class="wd-ranking '+(group.id==='best'?'best':'worst')+'"><h3>'+esc(group.title)+'</h3><ol class="wd-rank-list">'+(group.items||[]).map((row,index)=>
   '<li class="wd-rank-item"><span class="wd-rank-order">'+String(index+1).padStart(2,'0')+'</span><span class="wd-rank-name">'+esc(row.name)+'</span><strong class="wd-rank-value">'+esc(metricValue(row))+'</strong>'+((row.secondary||[]).length?'<span class="wd-rank-secondary">'+row.secondary.map(x=>esc(x.label)+': '+esc(metricValue(x))).join(' · ')+'</span>':'')+'</li>'
  ).join('')+'</ol></section>';
 }
 const metrics=detail.metrics||[];
 if(metrics.length)html+='<div class="wd-mini-metrics">'+metricRows(metrics)+'</div>';
 if(!groups.some(g=>g.items?.length)&&!metrics.length)html+='<p class="wd-detail-note">Источник пока не передал достаточных данных для детализации. Обновите карточку или откройте первоисточник.</p>';
 html+='<p class="wd-hint">'+esc(frequency(d))+(d.refresh_checked_at?' · проверено '+esc(stamp(d.refresh_checked_at)):'')+'</p>';
 document.getElementById('sourceDetailsBody').innerHTML=html;
 document.getElementById('sourceDetailsFooter').innerHTML='<a class="wd-source-link" href="'+url+'" target="_blank" rel="noopener noreferrer">Открыть первоисточник ↗</a><a class="wd-ai-link" href="/aichat?module=water-dashboard&amp;source='+encodeURIComponent(id)+'">Отчёт в ИИ-чате →</a>';
}
document.getElementById('kpiGrid').addEventListener('click',event=>{
 const refresh=event.target.closest('[data-refresh]');
 if(refresh){startRefresh(refresh.dataset.refresh);return;}
 const trigger=event.target.closest('[data-details]');if(!trigger)return;
 detailTrigger=trigger;openedSource=trigger.dataset.details;renderDetails(openedSource);drawer.showModal();
 document.getElementById('sourceDetailsBody').scrollTop=0;
});
document.getElementById('closeSourceDetails').addEventListener('click',()=>drawer.close());
drawer.addEventListener('click',event=>{if(event.target===drawer){const r=drawer.getBoundingClientRect();if(event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom)drawer.close();}});
drawer.addEventListener('close',()=>{
 const target=detailTrigger?.isConnected?detailTrigger:document.querySelector('[data-details="'+openedSource+'"]');
 openedSource=null;target?.focus();
});
function setBusy(value){busy=value;button.disabled=value;document.querySelectorAll('[data-refresh]').forEach(b=>b.disabled=value);}
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
async function startRefresh(sourceId=null){
 if(busy||submitting)return;submitting=true;refreshError='';setBusy(true);
 const definition=defs.find(d=>d[0]===sourceId);
 status.textContent=definition?'Обновление: '+definition[1]+'…':'Запускаю обновление '+defs.length+' источников…';
 try{const r=await fetch(definition?'/water-dashboard/refresh-source/'+sourceId:'/water-dashboard/run-check',{method:'POST'});const d=await readResponse(r);if(!r.ok||d.ok===false)throw Error(d.message||d.detail||'Не удалось запустить обновление');observedRunning=true;}
 catch(e){refreshError=e.message;status.textContent=e.message;setBusy(false);}
 finally{submitting=false;}
}
button.addEventListener('click',()=>startRefresh());
async function poll(){
 if(polling||loginRequired||submitting)return;polling=true;
 try{
  const r=await fetch('/water-dashboard/run-status',{cache:'no-store'});const state=await readResponse(r);if(!r.ok)throw Error('Не удалось проверить обновление');
  if(submitting)return;
  setBusy(!!state.running);button.textContent=busy?'Обновление…':'⟳ Обновить снимок';
  if(busy){observedRunning=true;status.textContent='Обновление: '+(state.stage||'сбор данных');return;}
  if(lastRevision===null||lastRevision!==state.snapshot_revision){
   const response=await fetch('/water-dashboard/snapshot',{cache:'no-store'});const latest=await readResponse(response);if(!response.ok)throw Error('Не удалось получить снимок');
   if(latest.checked_at!==snap.checked_at){snap=latest;render();}
   lastRevision=state.snapshot_revision??'';
  }
  const checked=snap.last_checked_sources||[];
  const lastName=checked.length===1?defs.find(([id])=>id===checked[0])?.[1]:'';
  const result=lastName?'Последняя проверка: '+lastName+' · '+stamp(snap.checked_at):'Обновлено '+defs.filter(([id])=>snap.sources?.[id]?.metric_schema===1&&snap.sources[id].ok).length+' из '+defs.length+' источников. Последняя проверка: '+stamp(snap.checked_at);
  status.textContent=refreshError||(state.last_error?'Ошибка обновления: '+state.last_error:snap.checked_at?result:'Данные ещё не загружены');
  // One refresh per page visit when the last attempt is older than 30 minutes.
  const elapsed=Date.now()-new Date(snap.checked_at||0).getTime();
  if(!autoAttempted&&!observedRunning){autoAttempted=true;if(snap.metric_schema!==1||!snap.checked_at||elapsed>30*60*1000)await startRefresh();}
 }catch(e){status.textContent=e.message;}finally{polling=false;}
}
render();poll();setInterval(poll,5000);
})();
