
const $ = id => document.getElementById(id);
const KEY = 'neurona.mingkh.sms.summary.v1';
let DEPT = '-1';   // исполнитель: задаёт только пресет ФКР
let OPT = null;

const esc = SMS.escape;
function picked(box){ return [...$(box).querySelectorAll('input:checked')].map(i=>i.value); }
function load(){ try{ return JSON.parse(localStorage.getItem(KEY)||'null'); }catch(e){ return null; } }
function save(){
  try{ localStorage.setItem(KEY, JSON.stringify({subs:picked('subs'), srcs:picked('srcs'),
                                                 omsu:$('omsu').value, dept:DEPT})); }catch(e){}
}

/* что выбрано — в свёрнутой шапке блока */
function summary(box, list){
  const sel = picked(box), n = list.length;
  const names = list.filter(x=>sel.includes(x.id)).map(x=>x.title);
  $(box+'Sel').textContent = !sel.length ? '— ничего не выбрано' :
    sel.length===n ? 'все (' + n + ')' :
    names.length<=3 ? names.join(', ') : names.length + ' из ' + n + ': ' + names.slice(0,2).join(', ') + '…';
}
function refresh(){
  summary('subs', OPT.subtopics); summary('srcs', OPT.sources);
  /* подсветка пресета, если выбор с ним совпадает */
  const subs = picked('subs').sort().join(), allSubs = OPT.subtopics.map(x=>x.id).sort().join();
  const fullRest = picked('srcs').length===OPT.sources.length;   // ОМСУ от пресета не зависит
  $('deptBadge').hidden = DEPT==='-1';
  document.querySelectorAll('.preset').forEach(b=>{
    const p = OPT.presets.find(x=>x.id===b.dataset.id);
    const ps = (p.subtopics==='all' ? allSubs : [...p.subtopics].sort().join());
    b.classList.toggle('on', fullRest && subs===ps && DEPT===p.dept);
  });
  save();
}
function setChecks(box, ids){ $(box).querySelectorAll('input').forEach(i=>{ i.checked = ids===true || Array.isArray(ids) && ids.includes(i.value); }); refresh(); }
function checks(box, items, on){
  $(box).innerHTML = items.map(x=>
    `<label><input type="checkbox" value="${esc(x.id)}" ${on.includes(x.id)?'checked':''}> ${esc(x.title)}</label>`).join('');
  $(box).addEventListener('change', refresh);
}

function applyPreset(id){
  const p = OPT.presets.find(x=>x.id===id);
  setChecks('subs', p.subtopics==='all' ? true : p.subtopics);
  setChecks('srcs', true);
  DEPT = p.dept;
  refresh();
}

async function init(){
  OPT = await (await SMS.fetch('/api/options')).json();
  const s = load();
  $('date').value = OPT.date;
  $('from').value = OPT.date.slice(0,4)+'-01-01';
  $('presets').innerHTML = OPT.presets.map(p=>
    `<button class="preset" data-id="${esc(p.id)}"><b>${esc(p.title)}</b><span>${esc(p.hint)}</span></button>`).join('');
  $('presets').querySelectorAll('.preset').forEach(b=>b.onclick=()=>applyPreset(b.dataset.id));
  $('omsu').innerHTML = '<option value="-1">Все ОМСУ</option>' +
    OPT.omsu.map(x=>`<option value="${esc(x.id)}">${esc(x.title)}</option>`).join('');
  $('omsu').onchange = refresh;
  $('deptOff').onclick = ()=>{ DEPT='-1'; refresh(); };
  checks('subs', OPT.subtopics, []);
  checks('srcs', OPT.sources, []);
  if(s){
    DEPT = s.dept || '-1';
    setChecks('subs', s.subs); setChecks('srcs', s.srcs);
    $('omsu').value = s.omsu || '-1';
    refresh();
  } else applyPreset('water');
}

document.querySelectorAll('[data-subs]').forEach(b=>b.onclick=()=>setChecks('subs', b.dataset.subs==='all'? true : []));
document.querySelectorAll('[data-src]').forEach(b=>b.onclick=()=>setChecks('srcs', b.dataset.src==='all'? true : []));

$('go').onclick = async ()=>{
  if(!$('from').value || $('from').value>$('date').value){ $('msg').className='status err'; $('msg').textContent='Проверьте период'; return; }
  const body = {from:$('from').value, date:$('date').value, subtopics:picked('subs'), sources:picked('srcs'),
                omsu:$('omsu').value, dept:DEPT};
  const m=$('msg'); m.className='status';
  if(!body.subtopics.length){ m.className='status err'; m.textContent='Выберите хотя бы одну тематику'; return; }
  if(!body.sources.length){ m.className='status err'; m.textContent='Выберите хотя бы один источник'; return; }
  const all = body.sources.length===OPT.sources.length;
  m.textContent = 'Собираю с портала… ' + (all? '' : '(источников: '+body.sources.length+', может занять минуту-другую)');
  $('go').disabled=true;
  const t0=Date.now();
  try{
    const r = await SMS.fetch('/api/build',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    const j = await r.json();
    if(!r.ok) throw new Error(j.error||r.status);
    $('out').value = j.text;
    m.textContent = 'Готово за '+Math.round((Date.now()-t0)/1000)+' с. Портал обновлён: '+j.portal_updated;
  }catch(e){ m.className='status err'; m.textContent='Ошибка: '+e.message; }
  $('go').disabled=false;
};

$('copy').onclick = async ()=>{
  try{ await navigator.clipboard.writeText($('out').value); }
  catch(e){ $('out').select(); document.execCommand('copy'); }
  $('copy').textContent='Скопировано'; setTimeout(()=>$('copy').textContent='Копировать',1500);
};


/* ---------- ежедневный доклад ---------- */
(()=>{ const t=new Date(); t.setDate(t.getDate()-1);   // по умолчанию — вчера
  $('dFrom').value=$('dDate').value=t.getFullYear()+'-'+String(t.getMonth()+1).padStart(2,'0')+'-'+String(t.getDate()).padStart(2,'0'); })();
$('dGo').onclick = async ()=>{
  const m=$('dMsg'); m.className='status';
  if(!$('dFrom').value || $('dFrom').value>$('dDate').value){ m.className='status err'; m.textContent='Проверьте период'; return; }
  m.textContent='Собираю с портала…'; $('dGo').disabled=true;
  try{
    const r=await SMS.fetch('/api/daily',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({from:$('dFrom').value, date:$('dDate').value})});
    const j=await r.json(); if(!r.ok) throw new Error(j.error||r.status);
    const dm=x=>x.split('-').reverse().join('.');
    $('dOut').value=j.text; const rng=(a,b)=>a===b? dm(b) : dm(a)+'–'+dm(b);
    m.textContent='Готово · сравнение с '+rng(j.prev_from,j.prev)+' · портал обновлён '+j.portal_updated;
    const pc=(c,p)=>p? Math.round((c-p)*100/p)+'%' : (c? 'рост' : '—');
    const tc=j.rows.reduce((s,r)=>s+r.curr,0), tp=j.rows.reduce((s,r)=>s+r.prev,0);
    $('dTbl').innerHTML='<thead><tr><th>Строка доклада</th><th>Подтема портала</th><th>'+rng(j.date_from,j.date)+'</th><th>'+rng(j.prev_from,j.prev)+'</th><th>Разница</th><th>%</th></tr></thead><tbody>'
      + j.rows.map(r=>'<tr><td>'+esc(r.icon)+esc(r.title)+'</td><td>'+esc(r.portal)+'</td><td>'+r.curr+'</td><td>'+r.prev+'</td><td>'+(r.curr-r.prev>0?'+':'')+(r.curr-r.prev)+'</td><td>'+pc(r.curr,r.prev)+'</td></tr>').join('')
      + '<tr><td><b>🔷Поступило / динамика</b></td><td>все подтемы</td><td><b>'+tc+'</b></td><td><b>'+tp+'</b></td><td><b>'+(tc-tp>0?'+':'')+(tc-tp)+'</b></td><td><b>'+pc(tc,tp)+'</b></td></tr></tbody>';
    $('dHow').hidden=false;
  }catch(e){ m.className='status err'; m.textContent='Ошибка: '+e.message; }
  $('dGo').disabled=false;
};
$('dCopy').onclick = async ()=>{ try{ await navigator.clipboard.writeText($('dOut').value); }catch(e){ $('dOut').select(); document.execCommand('copy'); } $('dCopy').textContent='Скопировано'; setTimeout(()=>$('dCopy').textContent='Копировать',1500); };
SMS.initialize(init);

SMS.bind();
