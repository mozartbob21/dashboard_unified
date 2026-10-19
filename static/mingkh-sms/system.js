
const $ = id => document.getElementById(id);
const KEY='neurona.mingkh.sms.system.v1';
let R=null;
const esc = SMS.escape;
const dmy = d => String(d.getDate()).padStart(2,'0')+'.'+String(d.getMonth()+1).padStart(2,'0')+'.'+d.getFullYear();
function save(){ try{ localStorage.setItem(KEY, JSON.stringify({ctrl:$('ctrl').value})); }catch(e){} }

$('ctrl').oninput=save;

$('go').onclick=async()=>{
  const m=$('msg'); m.className='status';
  const body={date:$('date').value, source:'gsheet'};
  if(!body.date){ m.className='status err'; m.textContent='Укажите дату отсчёта'; return; }
  $('go').disabled=true; const t0=Date.now();
  m.textContent='Беру реестр и жалобы ВС/ВО с портала, пересчитываю… это может занять несколько минут';
  try{
    const r=await SMS.fetch('/api/sys/run',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    const j=await r.json(); if(!r.ok) throw new Error(j.error||r.status);
    R=j; render();
    m.textContent='Готово за '+Math.round((Date.now()-t0)/1000)+' с · от '+j.source.name+' · жалоб ВС/ВО за год: '+j.complaints+' · результат сохранён как отправная точка на '+j.stamp;
    await regInfo();
  }catch(e){ m.className='status err'; m.textContent='Ошибка: '+e.message; }
  $('go').disabled=false;
};

function streetsTable(id, L, cols){
  $(id+'N').textContent = L.length ? L.length+' ул.' : 'нет';
  const link = t => '<a href="'+esc(SMS.safeLink(R.rm+'/issues/'+encodeURIComponent(t)))+'" target="_blank" rel="noopener">#'+esc(t)+'</a>';
  $(id).innerHTML = !L.length ? '<tbody><tr><td class="note">нет</td></tr></tbody>' :
    '<thead><tr><th>ОМСУ</th><th>Улица</th><th>Жалоб</th><th>Последняя жалоба</th><th>'+cols+'</th></tr></thead><tbody>'
    + L.map(s=>'<tr><td>'+esc(s.omsu)+'</td><td>'+esc(s.street)+'</td><td>'+s.n+'</td><td>'+esc(s.last)+'</td><td>'
      + (s.tasks ? s.tasks.map(link).join(', ') : esc(s.nums)) + '</td></tr>').join('') + '</tbody>';
}
function render(){
  const c=R.counts;
  $('out').value=R.sms;
  $('pills').innerHTML='<span class="pill">На контроле: '+c.total+'</span><span class="pill new">Новых: '+c.new+'</span>'
    +'<span class="pill ok">Перешли в «нет жалоб»: '+c.to_net+'</span><span class="pill back">Вернули «системный»: '+c.to_prod+'</span>'
    +'<span class="pill">Остаются «продолжаются»: '+c.stay_prod+'</span><span class="pill">Остаются «нет жалоб»: '+c.stay_net+'</span>';
  $('warn').textContent=(R.warning?'⚠ '+R.warning+' ':'')+(R.unmatched.length?'Не сопоставлено улиц реестра с жалобами: '+R.unmatched.length+' ('+R.unmatched.map(u=>u[1]).slice(0,5).join('; ')+(R.unmatched.length>5?'…':'')+')':'');
  streetsTable('tNew', R.new_streets, 'Номера жалоб');
  streetsTable('tBack', R.returned, 'Задачи RM');
  streetsTable('tNet', R.to_net, 'Задачи RM');
  $('tasks').value=R.tasks_line;
  $('res').classList.remove('hidden');
}

function download(bytes, name){ const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([bytes])); a.download=name; a.click(); URL.revokeObjectURL(a.href); }
document.querySelectorAll('[data-file]').forEach(b=>b.onclick=()=>{
  if(!R) return; const f=R.files[b.dataset.file];
  const bin=atob(f.b64), u=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i);
  download(u, f.name);
});
/* CSV для RM — как «загрузка в RM\системные.csv»: «;», Windows-1251, строка = новая улица */
$('rmCsv').onclick=()=>{
  if(!R) return; const m2=$('msg2'); m2.className='status';
  if(!R.new_streets.length){ m2.textContent='Новых системных улиц нет — файл не нужен'; return; }
  const ctrl=$('ctrl').value.trim(), ds=$('start').value.trim(), de=$('end').value.trim();
  if(!/^\d{2}\.\d{2}\.\d{4}$/.test(ds)||!/^\d{2}\.\d{2}\.\d{4}$/.test(de)){ m2.className='status err'; m2.textContent='Даты — в формате ДД.ММ.ГГГГ'; return; }
  const rows=[['№','улица','ОМСУ','назначено','контролер ','начало ','завершение','Описание','жалобы']];
  R.new_streets.forEach((s,i)=>rows.push([i+1, s.street, s.omsu, String(s.omsu).toUpperCase(), ctrl, ds, de, R.rm_desc, s.nums]));
  download(REZ.toCP1251(rows.map(r=>r.map(SMS.csvCell).map(REZ.csvField).join(';')).join('\r\n')+'\r\n'), 'Системные_новые_'+R.stamp+'.csv');
  m2.textContent='CSV для RM сформирован: улиц — '+R.new_streets.length+' ✓';
};
async function copy(el, btn){ try{ await navigator.clipboard.writeText(el.value); }catch(e){ el.select(); document.execCommand('copy'); } const t=btn.textContent; btn.textContent='Скопировано'; setTimeout(()=>btn.textContent=t,1500); }
$('copy').onclick=()=>copy($('out'),$('copy'));
$('copyTasks').onclick=()=>copy($('tasks'),$('copyTasks'));

/* какой реестр станет «прошлой неделей» для выбранной даты */
let REGS=[];
async function regInfo(){
  REGS=(await (await SMS.fetch('/api/sys/state')).json()).registries;
  const d=$('date').value, prev=REGS.filter(x=>x<d).pop(), last=REGS[REGS.length-1];
  const f=x=>esc(x.slice(8,10)+'.'+x.slice(5,7)+'.'+x.slice(0,4));
  $('regInfo').innerHTML = !REGS.length ? 'Сохранённых реестров нет.' :
    (prev ? 'Пересчёт пойдёт от реестра на <b>'+f(prev)+'</b>.' : 'Раньше этой даты сохранённого реестра нет — выберите дату позже '+f(REGS[0])+'.')
    +' Последний сохранённый — на '+f(last)+'. Каждый пересчёт сохраняет свой результат как отправную точку следующей недели.';
}
$('date').addEventListener('change', ()=>regInfo().catch(SMS.showError));

SMS.initialize(async()=>{
  const o=await (await SMS.fetch('/api/options')).json();
  $('date').value=o.system_date||o.date;
  try{ const s=JSON.parse(localStorage.getItem(KEY)||'null'); if(s) $('ctrl').value=s.ctrl||''; }catch(e){}
  const t=new Date(), e=new Date(); e.setDate(e.getDate()+5);
  $('start').value=dmy(t); $('end').value=dmy(e);
  await regInfo();
});

SMS.bind();

document.addEventListener('sms:imported',()=>regInfo().catch(SMS.showError));
