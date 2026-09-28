/* Report functions and CSV parser from «ЕДДС-мониторинг (2).zip».
   Executed INSIDE the authenticated ARM tab in the server's Chromium-GOST.
   Neurona adds only bounded, same-origin transport and response validation. */
async ({fromDate, toDate, coordinates, timeout, limit}) => {
  const origin = 'https://zkh-kontur.mosreg.ru';
  if(location.origin !== origin) return {error:'Откройте АРМ ЕДДС в рабочем Chromium-GOST.', status:409};
  const SITE = {path:'/new9/', act:'cds_report_svod', id:'3608', omsu:'0'};
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  const nativeFetch = window.fetch.bind(window);
  let size = 0, failure = null;
  const fail = (message, status=502) => {failure={error:message,status}; throw new Error(message);};
  const fetch = async (target, options={}) => {
    const url = new URL(target, origin + SITE.path);
    if(url.origin !== origin || url.username || url.password)
      return fail('АРМ ЕДДС вернул ссылку на другой адрес. Загрузка остановлена.');
    try {
      const response = await nativeFetch(url.href, {...options, mode:'same-origin',
        credentials:'same-origin', redirect:'follow', cache:'no-store', signal:controller.signal});
      if(response.status===401 || response.status===403)
        return fail('Сессия АРМ ЕДДС истекла или доступ запрещён. Проверьте сохранённый логин и пароль.',403);
      if(!response.ok) return fail('АРМ ЕДДС ответил кодом ' + response.status + '.');
      if(Number(response.headers.get('content-length')) > limit)
        return fail('Отчёт слишком большой. Сократите период.',413);
      const reader=response.body.getReader(); const chunks=[];
      while(true){
        const {value,done}=await reader.read(); if(done)break;
        size+=value.byteLength;
        if(size>limit){await reader.cancel(); return fail('Отчёт слишком большой. Сократите период.',413);}
        chunks.push(value);
      }
      const bytes=new Uint8Array(chunks.reduce((n,c)=>n+c.byteLength,0));let offset=0;
      for(const chunk of chunks){bytes.set(chunk,offset);offset+=chunk.byteLength;}
      let text;
      try{text=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}
      catch(_){text=new TextDecoder('windows-1251').decode(bytes);}
      return {ok:true,status:response.status,text:async()=>text};
    } catch(error) {
      if(failure)throw error;
      return fail(controller.signal.aborted ? 'АРМ ЕДДС не успел сформировать отчёт. Сократите период.' :
        'Chromium-GOST не смог получить отчёт. Проверьте портал в окне браузера на сервере.',controller.signal.aborted?504:502);
    }
  };
function parseCsv(txt, sep){
  const rows = [];
  let row = [], cell = '', q = false, fresh = true;   // fresh — мы в начале поля
  for(let i = 0; i < txt.length; i++){
    const c = txt[i];
    if(q){
      if(c === '"'){
        if(txt[i + 1] === '"'){ cell += '"'; i++; }   // удвоенная кавычка — это одна кавычка
        else q = false;
      }else cell += c;
      continue;
    }
    // Кавычка открывает поле только в его начале, внутри текста она обычный
    // символ. Выгрузка ЕДДС местами экранирует криво («СНТ ""Сырьево""»), и
    // если переключаться на любой кавычке, режим слетает посреди описания —
    // переводы строк начинают рвать запись. На сводке за неделю из-за этого
    // получалось 3658 записей вместо 2282 и 405 инцидентов ВС/ВО вместо 701.
    if(c === '"' && fresh){ q = true; fresh = false; continue; }
    if(c === sep){ row.push(cell); cell = ''; fresh = true; continue; }
    if(c === '\n'){ row.push(cell); rows.push(row); row = []; cell = ''; fresh = true; continue; }
    if(c === '\r') continue;
    cell += c; fresh = false;
  }
  if(cell.length || row.length){ row.push(cell); rows.push(row); }
  return rows;
}

/* Разделитель определяем по результату, а не по частоте символов: берём тот,
   что даёт больше колонок в шапке. */
function csvGrid(txt){
  const a = parseCsv(txt, ';'), b = parseCsv(txt, ',');
  const wa = a.length ? a[0].length : 0, wb = b.length ? b[0].length : 0;
  return wa >= wb ? a : b;
}

async function fetchReportCsv(ot, doo){
  const body = new URLSearchParams();
  body.set('act', SITE.act);
  body.set('id', SITE.id);
  body.set('date_ot', ot);
  body.set('date_do', doo);
  body.set('id_tu_mun_raion', SITE.omsu);
  body.set('saveToCSV', '1');

  const url = SITE.path + '?act=' + encodeURIComponent(SITE.act) + '&id=' + encodeURIComponent(SITE.id);
  let r;
  try{
    r = await fetch(url, {
      method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
      body,
    });
  }catch(_){ throw new Error('Не получилось обратиться к АРМ ЕДДС.'); }
  if(!r.ok) throw new Error('АРМ ЕДДС ответил кодом ' + r.status + ' на запрос CSV.');

  const txt = (await r.text()).replace(/^\uFEFF/, '');
  // вместо файла могла прийти страница: истёкшая сессия или отказ
  if(/^\s*<(!doctype|html)/i.test(txt))
    throw new Error(/act=login/.test(txt)
      ? 'Сессия истекла. Войдите в АРМ ЕДДС заново.'
      : 'Вместо CSV пришла страница.');
  const grid = csvGrid(txt);
  if(!grid.length) throw new Error('CSV пришёл пустым.');
  return grid;
}

const GEO_SITE = { act: 'cds_claim_report', id: '3654' };
const geoDay = iso => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? m[3] + '.' + m[2] + '.' + m[1] : '';
};

async function geoFetchLive(fromIso, toIso){
  const ot = geoDay(fromIso), doo = geoDay(toIso);
  if(!ot || !doo) throw new Error('Не задан период.');

  const body = new URLSearchParams();
  body.set('start', ot);
  body.set('end', doo);
  // чекбоксы формы шлются как «on» — так же, как их отправил бы браузер
  for(const f of ['id_cds_claim', 'lat_', 'lon_']) body.set('fields[' + f + ']', 'on');
  body.set('submit', '1');

  const url = SITE.path + '?act=' + encodeURIComponent(GEO_SITE.act) +
              '&id=' + encodeURIComponent(GEO_SITE.id);
  let r;
  try{
    r = await fetch(url, {
      method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
      body,
    });
  }catch(_){ throw new Error('Не получилось обратиться к АРМ ЕДДС.'); }
  if(!r.ok) throw new Error('АРМ ЕДДС ответил кодом ' + r.status + '.');

  let txt = (await r.text()).replace(/^﻿/, '');

  // Ответом бывает и сам файл, и страница со ссылкой на него — в АРМ отчёт
  // складывается в файл вида EDDS_claim_report_<время>.csv. Разбираем оба случая.
  if(/^\s*<(!doctype|html)/i.test(txt)){
    if(/act=login/.test(txt)) throw new Error('Сессия истекла. Войдите в АРМ ЕДДС заново.');
    const m = /href\s*=\s*["']([^"']*EDDS_claim_report[^"']*\.csv)["']/i.exec(txt) ||
              /href\s*=\s*["']([^"']+\.csv)["']/i.exec(txt);
    if(!m) throw new Error('В ответе нет ни файла, ни ссылки на него.');
    const link = new URL(m[1], location.origin + SITE.path).href;
    const r2 = await fetch(link, { credentials: 'same-origin' });
    if(!r2.ok) throw new Error('Файл отчёта не скачался (код ' + r2.status + ').');
    txt = (await r2.text()).replace(/^﻿/, '');
  }
  const grid = csvGrid(txt);
  if(!grid.length) throw new Error('Отчёт пришёл пустым.');
  return grid;
}

  try {
    const day = iso => iso.slice(8,10)+'.'+iso.slice(5,7)+'.'+iso.slice(2,4);
    const grid = coordinates ? await geoFetchLive(fromDate,toDate) : await fetchReportCsv(day(fromDate),day(toDate));
    const norm = v => String(v ?? '').replace(/\s+/g,' ').trim().toLowerCase();
    let header=-1;
    for(let i=0;i<Math.min(25,grid.length);i++){
      const row=grid[i].map(norm);
      if(!coordinates){if(row.includes('id_cds_claim')){header=i;break;}continue;}
      const lat=row.findIndex(v=>/широт/.test(v)||v==='lat_');
      const lon=row.findIndex(v=>/долгот/.test(v)||v==='lon_');
      const id=row.findIndex(v=>v==='id_cds_claim'||(/номер/.test(v)&&/заявк/.test(v))||/^№\s*заявк|^номер$|^№$/.test(v));
      if(lat>=0&&lon>=0&&id>=0){
        header=i;grid[i][lat]='Широта';grid[i][lon]='Долгота';grid[i][id]='Номер заявки';break;
      }
    }
    if(header<0) return {error:coordinates ? 'В отчёте нет колонок «Номер заявки», «Широта» и «Долгота».' :
      'В отчёте нет колонки id_cds_claim. Проверьте вход в АРМ ЕДДС и права на отчёт.',status:502};
    return {grid:grid.slice(header)};
  } catch(error) {
    return failure || {error:error.message,status:/[Сс]ессия/.test(error.message)?403:502};
  } finally {clearTimeout(timer);controller.abort();}
}
