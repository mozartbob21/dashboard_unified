'use strict';
const $ = id => document.getElementById(id);
const BLOCKS = [
  {id:'b1', title:'Свод по обращениям', page:'/mingkh/sms/summary', fields:['date']},
  {id:'b2', title:'Резонансные', page:'/mingkh/sms/resonance', fields:['from','to']},
  {id:'b3', title:'Системные', page:'/mingkh/sms/system', fields:['date']},
  {id:'b4', title:'Решения РМ', page:'/mingkh/sms/solutions', fields:['from','to']},
];
$('blocks').innerHTML = BLOCKS.map(b=>'<div class="blk"><div class="name"><label><input type="checkbox" id="'+b.id+'" checked> '+b.title+'</label><a href="'+b.page+'" target="_blank" rel="noopener">Настроить раздел ↗</a></div>'
  +'<div class="dates'+(b.fields.length===1?' single':'')+'">'+b.fields.map(f=>'<label for="'+b.id+'_'+f+'">'+(f==='from'?'С':'По')+'<input type="date" id="'+b.id+'_'+f+'" aria-label="'+b.title+': '+(f==='from'?'начало':'конец')+' периода"></label>').join('')+'</div><span class="st" id="'+b.id+'st"></span></div>').join('');
const val=(b,f)=>$(b.id+'_'+f).value;
function st(b,text,cls=''){const node=$(b.id+'st');node.textContent=text;node.className='st '+cls;}
function openPage(b){
  return new Promise((resolve,reject)=>{
    const frame=document.createElement('iframe'); frame.className='sms-runner-frame'; frame.tabIndex=-1; frame.title='Подготовка: '+b.title;
    const timeout=setTimeout(()=>{frame.remove();reject(new Error('Не удалось открыть раздел. Проверьте подключение и повторите сбор.'));},45000);
    frame.onload=async()=>{
      clearTimeout(timeout);
      try{
        if(!frame.contentWindow.SMS?.ready || !frame.contentWindow.SMS?.run) throw new Error('Не удалось открыть раздел. Возможно, нужно повторно войти в Нейрону.');
        await frame.contentWindow.SMS.ready; resolve(frame);
      }catch(error){frame.remove();reject(error);}
    };
    frame.onerror=()=>{clearTimeout(timeout);frame.remove();reject(new Error('Не удалось загрузить раздел.'));};
    frame.src=b.page+'?embed=1'; $('frames').appendChild(frame);
  });
}
async function runBlock(b){
  st(b,'Собираем данные…','run'); let frame;
  try{
    frame=await openPage(b);
    const result=await frame.contentWindow.SMS.run(Object.fromEntries(b.fields.map(f=>[f,val(b,f)])));
    st(b,result.message,'ok');return result.text;
  }finally{frame?.remove();}
}
$('go').onclick=async()=>{
  const message=$('msg');message.className='status';
  const todo=BLOCKS.filter(b=>$(b.id).checked);
  if(!todo.length){message.className='status err';message.textContent='Выберите хотя бы один блок.';return;}
  const bad=todo.find(b=>b.fields.some(f=>!val(b,f)) || (b.fields.length===2 && val(b,'from')>val(b,'to')));
  if(bad){message.className='status err';message.textContent='Проверьте период: '+bad.title;return;}
  BLOCKS.forEach(b=>st(b,$(b.id).checked?'В очереди':''));
  $('go').disabled=true; const start=Date.now(), texts=[],failed=[];
  message.textContent='Готовим разделы по очереди. Сбор данных может занять несколько минут.';
  try{
    for(const block of todo){
      try{texts.push(await runBlock(block));}
      catch(error){st(block,error.message,'err');failed.push(block.title);}
    }
    if(texts.length) $('out').value=texts.join('\n\n');
    message.className='status'+(failed.length?' err':'');
    message.textContent=failed.length ? 'Собрано '+texts.length+' из '+todo.length+' блоков. Не удалось получить: '+failed.join(', ')+'. Повторите сбор нужных разделов.' : 'СМС готова · '+Math.round((Date.now()-start)/1000)+' с';
  }finally{$('go').disabled=false;}
};
$('copy').onclick=async()=>{
  if(!$('out').value.trim())return;
  try{await navigator.clipboard.writeText($('out').value);}catch(_){$('out').select();document.execCommand('copy');}
  $('copy').textContent='Скопировано';setTimeout(()=>$('copy').textContent='Копировать',1500);
};
let OPT=null;
function lastWeek(){if(!OPT)return;BLOCKS.forEach(b=>b.fields.forEach(f=>$(b.id+'_'+f).value=b.id==='b3'?(OPT.system_date||OPT.week.to):f==='from'?OPT.week.from:OPT.week.to));}
$('lastWeek').onclick=lastWeek;
SMS.initialize(async()=>{OPT=await(await SMS.fetch('/api/options')).json();lastWeek();});
SMS.bind();
