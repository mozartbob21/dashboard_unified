
const $ = id => document.getElementById(id);
const KEY = 'neurona.mingkh.sms.resonance.v1';
const WORK_DOMAINS = ['konus.group', 'mosreg.ru'];   // рабочие учётки — не в топ жалобщиков
const FOCI_MIN = 2;                                   // очаг: жалоб по адресу за период
const FILL_DAYS = 90;                                 // первое заполнение базы найденного
let OPT = null, DEPT = '-1';
let R = null;   // результат поиска: {cases, prevNew, foci, prevFoci, users, period}

const esc = SMS.escape;
const iso = d => d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
const fromIso = s => { const [y,m,d]=s.split('-').map(Number); return new Date(y,m-1,d); };
const addDays = (s,n) => { const d=fromIso(s); d.setDate(d.getDate()+n); return iso(d); };
const ddmm = s => s.slice(8,10)+'.'+s.slice(5,7);
const dmy = s => s.slice(8,10)+'.'+s.slice(5,7)+'.'+s.slice(0,4);
const days = (a,b) => Math.round((fromIso(b)-fromIso(a))/86400000)+1;

/* ---------- фильтры: как на странице блока 1 ---------- */
function picked(box){ return [...$(box).querySelectorAll('input:checked')].map(i=>i.value); }
function save(){ try{ localStorage.setItem(KEY, JSON.stringify({subs:picked('subs'), srcs:picked('srcs'), omsu:$('omsu').value, dept:DEPT, thr:$('thr').value, ctrl:$('ctrl').value, foci:$('withFoci').checked})); }catch(e){} }
function load(){ try{ return JSON.parse(localStorage.getItem(KEY)||'null'); }catch(e){ return null; } }
function summary(box, list){
  const sel=picked(box), n=list.length, names=list.filter(x=>sel.includes(x.id)).map(x=>x.title);
  $(box+'Sel').textContent = !sel.length ? '— ничего не выбрано' : sel.length===n ? 'все ('+n+')' :
    names.length<=3 ? names.join(', ') : names.length+' из '+n+': '+names.slice(0,2).join(', ')+'…';
}
function refresh(){
  summary('subs', OPT.subtopics); summary('srcs', OPT.sources);
  $('deptBadge').hidden = DEPT==='-1';
  const subs=picked('subs').sort().join(), allSubs=OPT.subtopics.map(x=>x.id).sort().join();
  const fullSrc=picked('srcs').length===OPT.sources.length;
  document.querySelectorAll('.preset').forEach(b=>{
    const p=OPT.presets.find(x=>x.id===b.dataset.id);
    const ps=p.subtopics==='all'? allSubs : [...p.subtopics].sort().join();
    b.classList.toggle('on', fullSrc && subs===ps && DEPT===p.dept);
  });
  const f=$('from').value, t=$('to').value;
  if(f && t && f<=t){ const n=days(f,t); $('prevNote').textContent='Сравнение с '+dmy(addDays(f,-n))+' — '+dmy(addDays(f,-1))+' ('+n+' дн.)'; }
  save();
}
function setChecks(box, ids){ $(box).querySelectorAll('input').forEach(i=>{ i.checked = ids===true || Array.isArray(ids) && ids.includes(i.value); }); refresh(); }
function checks(box, items){ $(box).innerHTML=items.map(x=>`<label><input type="checkbox" value="${esc(x.id)}"> ${esc(x.title)}</label>`).join(''); $(box).addEventListener('change', refresh); }
function applyPreset(id){ const p=OPT.presets.find(x=>x.id===id); DEPT=p.dept; setChecks('subs', p.subtopics==='all'? true : p.subtopics); setChecks('srcs', true); }

async function init(){
  OPT = await (await SMS.fetch('/api/options')).json();
  const s = load();
  $('presets').innerHTML = OPT.presets.map(p=>`<button class="preset" data-id="${esc(p.id)}"><b>${esc(p.title)}</b><span>${esc(p.hint)}</span></button>`).join('');
  $('presets').querySelectorAll('.preset').forEach(b=>b.onclick=()=>applyPreset(b.dataset.id));
  $('omsu').innerHTML = '<option value="-1">Все ОМСУ</option>'+OPT.omsu.map(x=>`<option value="${esc(x.id)}">${esc(x.title)}</option>`).join('');
  checks('subs', OPT.subtopics); checks('srcs', OPT.sources);
  $('from').value=OPT.week.from; $('to').value=OPT.week.to;
  ['omsu','thr','from','to'].forEach(id=>$(id).addEventListener('change', refresh));
  $('ctrl').addEventListener('input', save);
  $('withFoci').addEventListener('change', ()=>{ save(); smsText(); });
  $('deptOff').onclick=()=>{ DEPT='-1'; refresh(); };
  $('lastWeek').onclick=()=>{ $('from').value=OPT.week.from; $('to').value=OPT.week.to; refresh(); };
  if(s){
    DEPT=s.dept||'-1'; setChecks('subs', s.subs); setChecks('srcs', s.srcs);
    $('omsu').value=s.omsu||'-1'; $('thr').value=s.thr||'7'; $('ctrl').value=s.ctrl||''; $('withFoci').checked=!!s.foci;
    refresh();
  } else applyPreset('water');
  foundInfo(await getFound());
}
document.querySelectorAll('[data-subs]').forEach(b=>b.onclick=()=>setChecks('subs', b.dataset.subs==='all'? true : []));
document.querySelectorAll('[data-src]').forEach(b=>b.onclick=()=>setChecks('srcs', b.dataset.src==='all'? true : []));

/* ---------- данные и база найденного ---------- */
async function post(url, body){
  const r=await SMS.fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const j=await r.json(); if(!r.ok) throw new Error(j.error||r.status); return j;
}
async function getFound(){ return (await SMS.fetch('/api/rez/found')).json(); }
function foundInfo(base){
  const n=Object.keys(base).length;
  $('foundInfo').textContent = n ? 'В базе найденного: '+n+' адресов' : 'База найденного пуста — при первом поиске заполнится за '+FILL_DAYS+' дней';
}
function filters(){ return {subtopics:picked('subs'), sources:picked('srcs'), omsu:$('omsu').value, dept:DEPT}; }

/* строка портала -> обращение (как строка детализации на старой странице) */
function items(pack, rows){
  const F=pack.fields;
  return rows.map(r=>{ const o={}; F.forEach((f,i)=>o[f]=r[i]); return REZ.makeItem(o); }).filter(Boolean);
}
const dateIso = it => it.date ? iso(it.date) : '';

/* кейсы периода: история адресов = архив + всё, что не позже конца периода */
function score(periodItems, pool, upTo){
  const DB=REZ.makeDB();
  DB.merge(pool.filter(it=>dateIso(it)<=upTo));
  DB.merge(periodItems);
  return REZ.buildCases(periodItems, DB);
}
/* точный адрес (улица, деревня, СНТ, дом) — иначе кейс не учитываем, см. rez_addr.js */
const precise = c => !!c.addr && REZADDR.precise(c.addr, REZ.shortAddr);
const passed = (cs, T) => cs.filter(c=>c.score>=T && precise(c));
const toFound = cs => cs.map(c=>({key:c.key, addr:c.addr, omsu:c.omsu, score:c.score}));

/* первый запуск: база найденного пуста — заполняем неделями за FILL_DAYS дней,
   как будто поиск запускали каждую неделю */
async function fillFound(f, T){
  const to=addDays(f,-1), from=addDays(f,-FILL_DAYS);
  $('msg').textContent='База найденного пуста — заполняю за '+FILL_DAYS+' дней ('+dmy(from)+' — '+dmy(to)+')…';
  const pack=await post('/api/rez/data', Object.assign(filters(), {from, to, prev_from:from, prev_to:to}));
  const cur=items(pack, pack.current), pool=items(pack, pack.history).concat(cur);
  let wf=from, total=0;
  while(wf<=to){
    const wt=addDays(wf,6)<to? addDays(wf,6) : to;
    const wk=cur.filter(it=>dateIso(it)>=wf && dateIso(it)<=wt);
    const cs=passed(score(wk, pool, wt), T);
    total+=(await post('/api/rez/found',{from:wf, to:wt, cases:toFound(cs)})).added;
    wf=addDays(wt,1);
  }
  return total;
}

$('go').onclick = async ()=>{
  const f=$('from').value, t=$('to').value, T=+$('thr').value, m=$('msg');
  m.className='status';
  if(!f || !t || f>t){ m.className='status err'; m.textContent='Проверьте период'; return; }
  if(!picked('subs').length){ m.className='status err'; m.textContent='Выберите хотя бы одну тематику'; return; }
  if(!picked('srcs').length){ m.className='status err'; m.textContent='Выберите хотя бы один источник'; return; }
  const n=days(f,t), pf=addDays(f,-n), pt=addDays(f,-1);
  $('go').disabled=true; const t0=Date.now();
  try{
    let base=await getFound(), filled=0;
    if(!Object.keys(base).length) filled=await fillFound(f, T);
    m.textContent='Беру обращения с портала…';
    const pack=await post('/api/rez/data', Object.assign(filters(), {from:f, to:t, prev_from:pf, prev_to:pt}));
    m.textContent='Считаю баллы…';
    const cur=items(pack, pack.current), prev=items(pack, pack.prev), hist=items(pack, pack.history);
    const pool=hist.concat(prev, cur);
    const cases=score(cur, pool, t), prevCases=score(prev, pool, pt);
    // в память: сначала прошлый период, потом текущий — так адрес, найденный
    // и там и там, для текущего периода окажется повторным
    await post('/api/rez/found',{from:pf, to:pt, cases:toFound(passed(prevCases,T))});
    const up=await post('/api/rez/found',{from:f, to:t, cases:toFound(passed(cases,T))});
    base=await getFound(); foundInfo(base);
    const firstOf = c => (base[c.key]||{}).first_from;
    for(const c of cases){ const ff=firstOf(c); c.isNew = !ff || ff>=f; c.firstFound=ff; c.sel = c.isNew && c.score>=T && precise(c); }
    const prevNew = passed(prevCases,T).filter(c=>{ const ff=firstOf(c); return !ff || ff>=pf; });
    R={cases, prevNew, period:{f,t,pf,pt,n}, T,
       foci:foci(cases, T, f), prevFoci:foci(prevCases, T, pf, base),
       users:topUsers(cur), stat:{rows:pack.current.length, items:cur.length}};
    $('ds').value=$('ds').value||dmy(iso(new Date())); $('de').value=$('de').value||dmy(addDays(iso(new Date()),3));
    $('res').classList.remove('hidden');
    render(); renderFoci(); renderUsers(); smsText();
    m.textContent='Готово за '+Math.round((Date.now()-t0)/1000)+' с'+(filled?' · база найденного заполнена: '+filled+' адресов':'')
      +' · новых в базе: '+up.added+' · история адресов с '+dmy(pack.history_from);
  }catch(e){ m.className='status err'; m.textContent='Ошибка: '+e.message;  }
  $('go').disabled=false;
};

/* ---------- очаги ---------- */
function foci(cases, T, from, base){
  base = base || null;
  return cases.filter(c=>precise(c) && (c.list.length>=FOCI_MIN ||
    (base ? ((base[c.key]||{}).first_from||'9') < from : (c.firstFound && c.firstFound < from))));
}

/* ---------- топ жалобщиков ---------- */
function topUsers(list){
  const work = m => WORK_DOMAINS.some(d=>m.endsWith('@'+d) || m.endsWith('.'+d));
  const by=new Map();
  for(const it of list){
    if(!it.mail || work(it.mail)) continue;
    let u=by.get(it.mail); if(!u){ u={mail:it.mail, n:0, addrs:new Set(), omsu:new Set()}; by.set(it.mail,u); }
    u.n++; if(it.addr) u.addrs.add(REZ.shortAddr(it.addr)); if(it.omsu) u.omsu.add(it.omsu);
  }
  return [...by.values()].sort((a,b)=>b.n-a.n || a.mail.localeCompare(b.mail));
}

/* ---------- таблица кейсов ---------- */
/* обращения кейса или очага, свежие сверху */
function complaints(c){
  return '<div class="detbox">'+c.list.slice().reverse().map(it=>'<div class="detrow"><span class="detmeta">'+REZ.fmtD(it.date)+' · '+esc(it.src)+' · ID '+esc(it.id)+(it.mail?' · '+esc(it.mail):'')+'</span><br>'+esc(it.desc.slice(0,600))+'</div>').join('')+'</div>';
}
function tier(s){ return s>=11?'t3' : s>=9?'t2' : s>=7?'t1':'t0'; }
function render(){
  const {cases,T}=R, showRep=$('showRep').checked;
  const hit=cases.filter(c=>c.score>=T && precise(c)), nNew=hit.filter(c=>c.isNew).length;
  const vague=cases.filter(c=>c.score>=T && !precise(c));
  $('pills').innerHTML='<span class="pill">Обращений за период: '+R.stat.items+'</span>'
    +'<span class="pill">Кейсов (адресов): '+cases.length+'</span>'
    +'<span class="pill">Выше порога '+T+': '+hit.length+'</span>'
    +'<span class="pill new">Новых: '+nNew+'</span>'
    +'<span class="pill rep">Повторных: '+(hit.length-nNew)+'</span>'
    +'<span class="pill" id="pillSel"></span>';
  $('vague').innerHTML = vague.length ? '<details><summary>Без точного адреса — не учитываются: <b>'+vague.length+'</b></summary><div class="twrap"><table>'
    +'<thead><tr><th>Балл</th><th>ОМСУ</th><th>Адрес</th><th>Обращений</th><th>Последнее обращение</th></tr></thead><tbody>'
    +vague.map(c=>'<tr><td>'+c.score+'</td><td>'+esc(c.omsu)+'</td><td>'+esc(c.addr)+'</td><td>'+c.n+'</td><td>'+esc(c.lastDesc.slice(0,160))+'</td></tr>').join('')
    +'</tbody></table></div></details>' : '';
  const shown=hit.filter(c=>showRep || c.isNew);
  let h='<thead><tr><th style="width:28px"></th><th style="width:70px">Балл</th><th>Статус</th><th>ОМСУ</th>'
      +'<th>Адрес (можно править)</th><th style="width:130px">Обращения</th><th>Признаки</th><th>Последнее обращение</th></tr></thead><tbody>';
  for(const c of shown){
    const i=cases.indexOf(c);
    const tags=c.tags.map(k=>'<span class="tag" title="'+esc(c.marks[k]||'')+'">'+REZ.TAG_NAMES[k]+'</span>').join('')||'—';
    const st=c.isNew? '<span class="st new">новый</span>' : '<span class="st rep" title="впервые найден '+(c.firstFound?dmy(c.firstFound):'')+'">повторный</span>';
    h+='<tr class="case'+(c.sel?'':' off')+'" data-i="'+i+'">'
      +'<td><input type="checkbox" class="ck" aria-label="Включить адрес '+esc(c.addr)+' в СМС" data-i="'+i+'"'+(c.sel?' checked':'')+'></td>'
      +'<td><span class="sc '+tier(c.score)+'">'+c.score+'</span><div class="axes">С'+c.sut+' Ж'+c.zhi+' А'+c.adr+'</div></td>'
      +'<td>'+st+'</td><td>'+esc(c.omsu)+'</td>'
      +'<td><textarea class="adr" aria-label="Адрес обращения" rows="'+Math.max(2,Math.ceil(c.addr.length/38))+'" data-i="'+i+'">'+esc(c.addr)+'</textarea></td>'
      +'<td><b>'+c.n+'</b> всего'+(c.nNew<c.n?' (+'+c.nNew+' за период)':'')+'<br>за 90 дн: <b>'+c.nAct+'</b> от <b>'+(c.aAct||'?')+'</b> чел.'
      +(c.days?'<br>тянется '+c.days+' дн':'')+'</td>'
      +'<td>'+tags+'</td>'
      +'<td><div class="txtprev">'+(c.open?'▾ ':'▸ ')+esc(c.lastDesc.slice(0,200))+'</div></td></tr>';
    if(c.open){
      h+='<tr><td colspan="8">'+complaints(c)+'</td></tr>';
    }
  }
  $('tbl').innerHTML=h+'</tbody>';
  $('tblN').textContent='('+shown.length+')';
  $('tbl').querySelectorAll('.ck').forEach(x=>{ x.onclick=e=>e.stopPropagation(); x.onchange=e=>{ cases[+e.target.dataset.i].sel=e.target.checked; e.target.closest('tr').classList.toggle('off',!e.target.checked); updSel(); }; });
  $('tbl').querySelectorAll('.adr').forEach(a=>{ a.onclick=e=>e.stopPropagation(); a.oninput=e=>{ cases[+e.target.dataset.i].addr=e.target.value.trim(); }; });
  $('tbl').querySelectorAll('tr.case').forEach(tr=>{ const c=cases[+tr.dataset.i]; tr.tabIndex=0; tr.setAttribute('aria-expanded',String(!!c.open)); tr.setAttribute('aria-label','Обращения по адресу '+c.addr); tr.onclick=()=>{c.open=!c.open;render();}; tr.onkeydown=event=>{if(event.target===tr && ['Enter',' '].includes(event.key)){event.preventDefault();tr.click();}}; });
  updSel();
}
function updSel(){ $('pillSel').textContent='Отмечено: '+R.cases.filter(c=>c.sel).length; smsText(); }
$('showRep').onchange=()=>R&&render();
$('allOn').onclick=()=>{ R.cases.forEach(c=>{ if(c.isNew && c.score>=R.T && precise(c)) c.sel=true; }); render(); };
$('allOff').onclick=()=>{ R.cases.forEach(c=>c.sel=false); render(); };

function renderFoci(){
  const L=R.foci.slice().sort((a,b)=>b.list.length-a.list.length || b.score-a.score);
  $('fociPills').innerHTML='<span class="pill">Очагов: '+L.length+'</span><span class="pill">в прошлом периоде: '+R.prevFoci.length+'</span>';
  $('fociN').textContent='('+L.length+')';
  $('foci').innerHTML='<thead><tr><th>ОМСУ</th><th>Адрес</th><th>Жалоб за период</th><th>Всего / за 90 дн</th><th>Балл</th><th>Ранее найден резонансным</th><th>Последнее обращение</th></tr></thead><tbody>'
    +L.map((c,i)=>'<tr class="focus" data-i="'+i+'"><td>'+esc(c.omsu)+'</td><td>'+esc(REZ.shortAddr(c.addr))+'</td><td><b>'+c.list.length+'</b></td><td>'+c.n+' / '+c.nAct+'</td><td>'+c.score+'</td><td>'
      +(c.firstFound && c.firstFound<R.period.f ? dmy(c.firstFound) : '—')+'</td><td><div class="txtprev">'+(c.fopen?'▾ ':'▸ ')+esc(c.lastDesc.slice(0,160))+'</div></td></tr>'
      +(c.fopen ? '<tr><td colspan="7">'+complaints(c)+'</td></tr>' : '')).join('')+'</tbody>';
  $('foci').querySelectorAll('tr.focus').forEach(tr=>{const c=L[+tr.dataset.i];tr.tabIndex=0;tr.setAttribute('aria-expanded',String(!!c.fopen));tr.setAttribute('aria-label','Обращения по адресу '+c.addr);tr.onclick=()=>{c.fopen=!c.fopen;renderFoci();};tr.onkeydown=event=>{if(event.target===tr && ['Enter',' '].includes(event.key)){event.preventDefault();tr.click();}};});
  R.fociSorted=L;
}
function renderUsers(){
  const L=R.users.slice(0,30);
  $('usersN').textContent='('+L.length+')';
  $('users').innerHTML='<thead><tr><th>#</th><th>Почта</th><th>Жалоб</th><th>Адресов</th><th>ОМСУ</th></tr></thead><tbody>'
    +L.map((u,i)=>'<tr><td>'+(i+1)+'</td><td>'+esc(u.mail)+'</td><td><b>'+u.n+'</b></td><td>'+u.addrs.size+'</td><td>'+esc([...u.omsu].join(', '))+'</td></tr>').join('')+'</tbody>';
}

/* ---------- текст СМС ---------- */
/* «СОЛНЕЧНОГОРСК», «сергиево-посадский» -> «Солнечногорск», «Сергиево-Посадский» */
function omsuName(s){ s=String(s||'').trim().toLowerCase(); return s ? s.replace(/(^|[\s-])([а-яёa-z])/g,(m,a,b)=>a+b.toUpperCase()) : 'без ОМСУ'; }
/* топ ОМСУ: первые три места и все, у кого столько же, сколько у третьего (не меньше 2) */
function topOmsu(cs){
  const by=new Map(); cs.forEach(c=>by.set(omsuName(c.omsu),(by.get(omsuName(c.omsu))||0)+1));
  const L=[...by.entries()].sort((a,b)=>b[1]-a[1] || a[0].localeCompare(b[0]));
  if(!L.length) return [];
  let cut=L[Math.min(2,L.length-1)][1]; if(L[0][1]>=2) cut=Math.max(cut,2);
  return L.filter(x=>x[1]>=cut).map(([o,n])=>o+' - '+n);
}
const diff = (a,b) => { const d=a-b; return d>0? '(+'+d+')' : d<0? '('+d+')' : '(0)'; };
function smsText(){
  if(!R) return;
  const {f,t,n}=R.period, lbl=(n===7?'за неделю':'за период')+' ('+ddmm(f)+'-'+ddmm(t)+')';
  const sel=R.cases.filter(c=>c.sel && c.isNew);
  const out=['Резонансные обращения '+lbl+': '+sel.length+' '+diff(sel.length, R.prevNew.length)];
  const top=topOmsu(sel); if(top.length) out.push('ТОП ОМСУ:', ...top);
  if($('withFoci').checked){
    out.push('', 'Очаги (2+ жалобы по адресу) '+lbl+': '+R.foci.length+' '+diff(R.foci.length, R.prevFoci.length));
    const tf=topOmsu(R.foci); if(tf.length) out.push('ТОП ОМСУ:', ...tf);
  }
  $('out').value=out.join('\n');
}
$('copy').onclick=async()=>{ try{ await navigator.clipboard.writeText($('out').value); }catch(e){ $('out').select(); document.execCommand('copy'); } $('copy').textContent='Скопировано'; setTimeout(()=>$('copy').textContent='Копировать',1500); };

/* ---------- выгрузки: как кнопки старой страницы ---------- */
function msg2(t,cls){ $('msg2').textContent=t; $('msg2').className='status '+(cls||''); }
$('dlCsv').onclick=()=>{
  const sel=R.cases.filter(c=>c.sel);
  if(!sel.length) return msg2('Не отмечено ни одного кейса','err');
  if(sel.some(c=>!c.addr)) return msg2('У отмеченных кейсов не заполнен адрес','err');
  const ctrl=$('ctrl').value.trim(), ds=$('ds').value.trim(), de=$('de').value.trim();
  if(!/^\d{2}\.\d{2}\.\d{4}$/.test(ds)||!/^\d{2}\.\d{2}\.\d{4}$/.test(de)) return msg2('Даты должны быть в формате ДД.ММ.ГГГГ','err');
  const text=REZ.rmRows(sel,ctrl,ds,de).map(r=>r.map(SMS.csvCell).map(REZ.csvField).join(';')).join('\r\n')+'\r\n';
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([REZ.toCP1251(text)],{type:'text/csv;charset=windows-1251'}));
  a.download='Резонансные_'+ds.replace(/\./g,'-')+'.csv'; a.click(); URL.revokeObjectURL(a.href);
  msg2('CSV для РМ сформирован: кейсов — '+sel.length+' (обращений внутри: '+sel.reduce((s,c)=>s+c.n,0)+') ✓');
};
$('dlFoci').onclick=()=>{
  if(!R || !R.fociSorted || !R.fociSorted.length) return msg2('Очагов нет','err');
  const aoa=[['ОМСУ','Адрес','Жалоб за период','Всего жалоб по адресу','За 90 дней','Балл','Признаки','Ранее найден резонансным','Номера жалоб за период','Последнее обращение']];
  R.fociSorted.forEach(c=>aoa.push([c.omsu, REZ.shortAddr(c.addr), c.list.length, c.n, c.nAct, c.score,
    c.tags.map(k=>REZ.TAG_NAMES[k]).join(', '), c.firstFound && c.firstFound<R.period.f ? dmy(c.firstFound) : '',
    c.list.map(it=>it.id).filter(Boolean).join(', '), c.lastDesc]));
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[{wch:18},{wch:48},{wch:10},{wch:10},{wch:10},{wch:7},{wch:30},{wch:14},{wch:40},{wch:70}];
  ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:aoa.length-1,c:aoa[0].length-1}})};
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Очаги');
  XLSX.writeFile(wb, 'Очаги_'+dmy(R.period.f).replace(/\./g,'-')+'_'+dmy(R.period.t).replace(/\./g,'-')+'.xlsx');
  msg2('Excel очагов сформирован: '+(aoa.length-1)+' адресов ✓');
};
$('dlXls').onclick=()=>{
  const sel=R.cases.filter(c=>c.sel);
  if(!sel.length) return msg2('Не отмечено ни одного кейса','err');
  const aoa=REZ.registryRows(sel), ws=XLSX.utils.aoa_to_sheet(aoa,{cellDates:true});
  for(let r=1;r<aoa.length;r++) for(const col of [1,12]){ const cell=ws[XLSX.utils.encode_cell({r,c:col})]; if(cell && cell.v instanceof Date){ cell.t='d'; cell.z='dd.mm.yyyy'; } }
  ws['!cols']=[{wch:14},{wch:11},{wch:18},{wch:38},{wch:10},{wch:24},{wch:70},{wch:28},{wch:40},{wch:8},{wch:8},{wch:20},{wch:13},{wch:16}];
  ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:aoa.length-1,c:aoa[0].length-1}})};
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Резонансные');
  XLSX.writeFile(wb, 'Резонансные_реестр_'+$('ds').value.trim().replace(/\./g,'-')+'.xlsx');
  msg2('Excel-реестр сформирован: обращений — '+(aoa.length-1)+' (кейсов — '+sel.length+') ✓');
};

SMS.initialize(init);

SMS.bind();
