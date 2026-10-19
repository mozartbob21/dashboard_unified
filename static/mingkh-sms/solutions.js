
const $ = id => document.getElementById(id);
let OPT=null, D=null;   // D: {measures, rm}

const esc = SMS.escape;
const fromIso = s => { const [y,m,d]=s.split('-').map(Number); return new Date(y,m-1,d); };
const days = (a,b) => Math.round((fromIso(b)-fromIso(a))/86400000)+1;
const ddmm = s => s.slice(8,10)+'.'+s.slice(5,7);
const dmy = s => s ? s.slice(8,10)+'.'+s.slice(5,7)+'.'+s.slice(0,4) : '';
/* «КОЛОМНА» -> «Коломна», «СЕРГИЕВО-ПОСАДСКИЙ» -> «Сергиево-Посадский» */
const omsuName = s => { s=String(s||'').trim().toLowerCase(); return s ? s.replace(/(^|[\s-])([а-яёa-z])/g,(m,a,b)=>a+b.toUpperCase()) : 'без ОМСУ'; };
const plural = (n, one, few, many) => { const a=n%10, b=n%100; return n+' '+(a===1&&b!==11? one : a>=2&&a<=4&&(b<12||b>14)? few : many); };
const addrN = n => plural(n,'адрес','адреса','адресов');
const evN = n => plural(n,'мероприятие','мероприятия','мероприятий');
const akey = m => REZ.addrKey(m.subject) || ('#'+m.id);

/* место для системных: части адреса без ОМСУ/округа и без номера дома,
   «село» -> «с.», «улица» -> «ул.»; общие для адресов ОМСУ (не больше трёх) */
const HOUSE=/^(д\.?\s*|дом\s*|к\.?\s*|корп\.?\s*)?\d+[а-яa-z]?([\/\-]\d+[а-я]?)?(\s*(к|корп|корпус|с|стр|строение)\.?\s*\d+[а-я]?)?$/i;
const ABBR=[[/^село\s+/i,'с. '],[/^деревня\s+/i,'д. '],[/^посёлок городского типа\s+|^поселок городского типа\s+/i,'пгт '],[/^посёлок\s+|^поселок\s+/i,'п. '],
            [/^микрорайон\s+/i,'мкр '],[/\s+улица$/i,' ул.'],[/^улица\s+/i,'ул. '],[/\s+проспект$/i,' пр-т'],[/\s+переулок$/i,' пер.'],[/\s+шоссе$/i,' ш.']];
function placeParts(subject, omsu){
  const o=omsu.toLowerCase().replace(/ё/g,'е');
  return REZ.shortAddr(subject).split(',').map(s=>s.trim()).filter(p=>{
    const l=p.toLowerCase().replace(/ё/g,'е');
    return p && !HOUSE.test(p) && !/городской округ|муниципальный округ|^г\.\s*о\.|^\d+-й микрорайон$/.test(l) && l!==o && !l.startsWith(o+' ');
  }).map(p=>ABBR.reduce((s,[re,to])=>s.replace(re,to),p));
}
function place(list, omsu){
  const seen=[];
  for(const m of list) for(const p of placeParts(m.subject, omsu)) if(!seen.includes(p)) seen.push(p);
  return seen.length && seen.length<=3 ? ' ('+seen.join(', ')+')' : '';
}

/* ТОП по адресам: все ОМСУ, как в СМС (за 01–07 и 08–14.10 перечислены все) */
function topOmsu(byOmsu){
  return [...byOmsu.entries()].map(([o,ms])=>[o,new Set(ms.map(akey)).size]).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
}
const group = (list, f) => { const m=new Map(); list.forEach(x=>{ const k=f(x); if(!m.has(k)) m.set(k,[]); m.get(k).push(x); }); return m; };

function smsText(){
  const f=$('from').value, t=$('to').value, n=days(f,t);
  const lbl = n===7 ? 'за неделю' : 'за период '+ddmm(f)+'-'+ddmm(t);
  const L=shown(), out=[];

  /* быстрые */
  const q=L.filter(m=>m.group==='quick');
  out.push('Быстрые решения '+lbl+':');
  if(!q.length) out.push('Промывки - нет');
  else{
    out.push('Промывки - '+addrN(new Set(q.map(akey)).size)+' / '+evN(q.length));
    const top=topOmsu(group(q, m=>omsuName(m.omsu)));
    out.push('ТОП по адресам:', ...top.map(([o,c],i)=>o+': '+c+(i===top.length-1?'.':'')));
    out.push('', 'Разбивка по типам мероприятий:');
    for(const [fld,ms] of [...group(q,m=>m.field).entries()].sort((a,b)=>b[1].length-a[1].length)) out.push(fld+' — '+ms.length);
  }

  /* системные */
  const s=L.filter(m=>m.group==='system');
  out.push('');
  if(!s.length) out.push('Системные решения '+lbl+': нет');
  else{
    out.push('Системные решения '+lbl+': '+addrN(new Set(s.map(akey)).size)+' / '+evN(s.length));
    const rows=[];
    const byO=[...group(s, m=>omsuName(m.omsu)).entries()]
      .sort((a,b)=>new Set(b[1].map(akey)).size-new Set(a[1].map(akey)).size || a[0].localeCompare(b[0]));
    for(const [o,ms] of byO){
      const na=new Set(ms.map(akey)).size, pl=place(ms, o), byF=[...group(ms,m=>m.field).entries()];
      if(na>1){
        rows.push(o+pl+' — '+addrN(na)+':');
        byF.forEach(([fld,x])=>rows.push(fld+' — '+x.length+';'));
      } else rows.push(o+pl+' — '+byF.map(([fld,x])=>x[0].short+(x.length>1?' — '+x.length:'')).join(', ')+';');
    }
    rows[rows.length-1]=rows[rows.length-1].replace(/;$/,'.');
    out.push(...rows);
  }
  $('out').value=out.join('\n');
}

function shown(){ const o=$('omsu').value; return D ? D.measures.filter(m=>!o || m.omsu===o) : []; }

function render(){
  const L=shown(), q=L.filter(m=>m.group==='quick'), s=L.filter(m=>m.group==='system');
  $('pills').innerHTML='<span class="pill">Мероприятий: '+L.length+'</span><span class="pill">Задач: '+new Set(L.map(m=>m.id)).size+'</span>'
    +'<span class="pill">Быстрые: '+addrN(new Set(q.map(akey)).size)+' / '+evN(q.length)+'</span>'
    +'<span class="pill">Системные: '+addrN(new Set(s.map(akey)).size)+' / '+evN(s.length)+'</span>';
  $('tbl').innerHTML='<thead><tr><th>№ задачи</th><th>ОМСУ</th><th>Адрес</th><th>Тип задачи</th><th>Статус</th><th>Решение</th><th>Мероприятие</th><th>Дата</th></tr></thead><tbody>'
    +L.map(m=>'<tr><td><a href="'+esc(SMS.safeLink(m.url))+'" target="_blank" rel="noopener">#'+esc(m.id)+'</a></td><td>'+esc(omsuName(m.omsu))+'</td><td>'+esc(m.subject)+'</td><td>'+esc(m.type)+'</td><td>'+esc(m.status)+'</td>'
      +'<td><span class="grp '+(m.group==='quick'?'quick':'system')+'">'+(m.group==='quick'?'быстрое':'системное')+'</span></td><td>'+esc(m.field)+'</td><td>'+esc(dmy(m.date))+'</td></tr>').join('')+'</tbody>';
  smsText();
}

$('go').onclick=async()=>{
  const f=$('from').value, t=$('to').value, msg=$('msg'); msg.className='status';
  if(!f||!t||f>t){ msg.className='status err'; msg.textContent='Проверьте период'; return; }
  $('go').disabled=true; msg.textContent='Читаю задачи RM…'; const t0=Date.now();
  try{
    const r=await SMS.fetch('/api/rm/measures',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({from:f,to:t})});
    const j=await r.json(); if(!r.ok) throw new Error(j.error||r.status);
    D=j; D.period={f,t};
    const keep=$('omsu').value, om=[...new Set(D.measures.map(m=>m.omsu))].sort();
    $('omsu').innerHTML='<option value="">Все ОМСУ</option>'+om.map(o=>'<option value="'+esc(o)+'">'+esc(omsuName(o))+'</option>').join('');
    $('omsu').value=om.includes(keep)?keep:'';
    $('res').classList.remove('hidden'); render();
    msg.textContent='Готово за '+Math.round((Date.now()-t0)/1000)+' с · данные RM на момент загрузки';
  }catch(e){ msg.className='status err'; msg.textContent='Ошибка: '+e.message; }
  $('go').disabled=false;
};
$('omsu').onchange=()=>D&&render();
['from','to'].forEach(id=>$(id).onchange=()=>{ if(D) $('msg').textContent='Период изменён — нажмите «Загрузить из RM»'; });
$('lastWeek').onclick=()=>{ $('from').value=OPT.week.from; $('to').value=OPT.week.to; };
$('copy').onclick=async()=>{ try{ await navigator.clipboard.writeText($('out').value); }catch(e){ $('out').select(); document.execCommand('copy'); } $('copy').textContent='Скопировано'; setTimeout(()=>$('copy').textContent='Копировать',1500); };

/* выгрузка: строка = мероприятие, № задачи — гиперссылка на RM */
$('xls').onclick=()=>{
  const L=shown();
  const aoa=[['№ задачи','Ссылка','ОМСУ','Адрес','Тип задачи','Статус','Решение','Мероприятие','Дата']];
  L.forEach(m=>aoa.push([m.id, m.url, omsuName(m.omsu), m.subject, m.type, m.status, m.group==='quick'?'быстрое':'системное', m.field, fromIso(m.date)]));
  const ws=XLSX.utils.aoa_to_sheet(aoa,{cellDates:true});
  for(let r=1;r<aoa.length;r++){
    const c=ws[XLSX.utils.encode_cell({r,c:0})]; c.l={Target:SMS.safeLink(aoa[r][1]), Tooltip:'Открыть задачу в RM'};
    const d=ws[XLSX.utils.encode_cell({r,c:8})]; if(d){ d.t='d'; d.z='dd.mm.yyyy'; }
  }
  ws['!cols']=[{wch:10},{wch:38},{wch:20},{wch:60},{wch:16},{wch:22},{wch:11},{wch:40},{wch:11}];
  ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:aoa.length-1,c:aoa[0].length-1}})};
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Мероприятия');
  XLSX.writeFile(wb, 'Решения_RM_'+dmy(D.period.f).replace(/\./g,'-')+'_'+dmy(D.period.t).replace(/\./g,'-')+'.xlsx');
};

SMS.initialize(async()=>{ OPT=await (await SMS.fetch('/api/options')).json(); $('from').value=OPT.week.from; $('to').value=OPT.week.to; });

SMS.bind();
