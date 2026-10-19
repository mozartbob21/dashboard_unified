/* rez_core.js — СОБРАНО АВТОМАТИЧЕСКИ rez/build_core.py, руками не править.
   Логика поиска резонанса дословно из
   ЖКХ\Мой сайт по контролю\Резонанстные жалобы\Вода_жалобы.html */
(function(root){
'use strict';

function toCP1251(str){
  const out=[];
  const MAP={0x0401:0xA8,0x0451:0xB8,0x2116:0xB9,0x00AB:0xAB,0x00BB:0xBB,0x2014:0x97,0x2013:0x96,
             0x2018:0x91,0x2019:0x92,0x201C:0x93,0x201D:0x94,0x2026:0x85,0x00A0:0xA0,0x2022:0x95,0x00B7:0xB7};
  for(const ch of str){
    const c=ch.codePointAt(0);
    if(c<128) out.push(c);
    else if(c>=0x0410&&c<=0x044F) out.push(c-0x0410+0xC0);
    else if(MAP[c]!==undefined) out.push(MAP[c]);
    else out.push(0x3F);
  }
  return new Uint8Array(out);
}
function csvField(v){
  const s=String(v==null?'':v);
  return /[;"\n\r]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s;
}
function downloadCsv1251(rows, filename){
  const text=rows.map(r=>r.map(csvField).join(';')).join('\r\n')+'\r\n';
  const blob=new Blob([toCP1251(text)],{type:'text/csv;charset=windows-1251'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob); a.download=filename; a.click();
  URL.revokeObjectURL(a.href);
}
/* ===================== ПРАВИЛА ОТБОРА =====================
   Три независимые оси. Источник обращения намеренно НЕ участвует
   в баллах: важен факт проблемы и её суть, а не канал поступления.
   ========================================================== */

/* --- ОСЬ 1. СУТЬ: что фактически произошло --- */
const SUT_MARKS = [
  {k:'sanit',  w:3, name:'санитарная угроза',
   re:/черв|личинк|опарыш|фекал|нечистот|сероводород|зловон|тухл[а-яё]*\s*(запах|вод)|канализац[а-яё]{0,4}\s*(течет|течёт|прорв|затоп|в\s+подвал)|стоки\s+(в|на)\s|затоплен[а-яё]*\s+подвал/i},
  {k:'noaqua', w:3, name:'нет воды более суток',
   re:/(без\s+вод[ыу]|нет\s+вод[ыы]|отсутств[а-яё]*\s+вод[ыа]|отключил[а-яё]*\s+вод)[^.!?]{0,60}(\d+\s*(ден|дня|дней|сут|недел|месяц)|недел|месяц|двое|трое|сутк)|(\d+\s*(ден|дня|дней|суток|недел|месяц)|вторые\s+сутки|третьи\s+сутки)[^.!?]{0,40}(без\s+вод|нет\s+вод|отсутств[а-яё]*\s+вод)/i},
  {k:'vulner', w:3, name:'уязвимые жильцы',
   re:/грудн[а-яё]*\s*(ребен|реб[её]н|млад)|младен|новорожд|грудничк|\d\s*-?\s*месячн[а-яё]*\s*(ребен|реб[её]н|малыш)|инвалид|лежач[а-яё]*\s*(больн|бабушк|дедушк)|колясочник|маломобильн|беременн|многодетн/i},
  {k:'unfit',  w:2, name:'вода непригодна',
   re:/непригодн|нельзя\s+пить|невозможно\s+(пить|мыться|купать|стирать|готовить)|не\s+пригодн|опасн[а-яё]*\s+для\s+здоров|отравл|сыпь|аллерг|инфекц/i},
  {k:'social', w:2, name:'соцобъект',
   re:/детск[а-яё]*\s+сад|садик|\bшкол[аеуы]|больниц|поликлиник|\bфап\b|интернат|роддом|амбулатор/i},
  {k:'buy',    w:2, name:'покупают воду',
   re:/покупа[а-яё]*\s+воду|вынужден[а-яё]*\s+покупать|бутилирован|возим\s+воду|привозим\s+воду|за\s+свой\s+сч[её]т\s+воду|носим\s+воду/i},
  {k:'scale',  w:2, name:'массовость',
   re:/вес[ьъ]\s+(дом|поселок|посёлок|микрорайон|подъезд)|вся\s+(улица|деревня)|все\s+жител|\d+\s*(домов|квартир|подъезд)/i},
];

/* --- ОСЬ 2. ЖИТЕЛЬ: кто пишет и как ---
   Внимание: в escal НЕТ голого «суд» — оно ловило «посуду»
   («посуда покрывается налётом» — типовая фраза про ржавую воду). */
const ZHI_MARKS = [
  {k:'escal', w:3, name:'эскалация',
   re:/прокурат|президент|губернатор|роспотребнадзор|следствен|судебн|\bв\s+суд\b|пода(м|ю|ём|ем)\s+иск|исково[ег]|телевид|журналист|(^|[^а-яё])сми([^а-яё]|$)|жилинспекц|гжи\b/i},
  {k:'again', w:2, name:'повторно',
   re:/повторн[а-яё]*\s+обраща|не\s+первый\s+раз|неоднократн|обраща[а-яё]*\s+уже|писал[а-яё]*\s+(уже\s+)?(\d+|много|неоднократ)|в\s+который\s+раз|снова\s+обраща|обращались\s+уже/i},
  {k:'coll',  w:2, name:'коллективное',
   re:/коллективн|все\s+жильцы|жители\s+дома|подпис(ей|и|ями)|от\s+лица\s+жител|общедомов[а-яё]*\s+собран|от\s+имени\s+жител/i},
];

/* --- пороговые шкалы ---
   Активность считается в скользящем окне: за всё время история есть почти
   у любого адреса, и ось перестаёт различать. Окно показывает, что проблема
   живая сейчас, а «хроника» — что она при этом ещё и давняя. */
const ACT_WINDOW = 90;                                           // дней
function ptsAuthors(n){ return n>=5?3 : n>=3?2 : n>=2?1 : 0; }   // разных заявителей за окно
function ptsCount(n){   return n>=10?3 : n>=5?2 : n>=3?1 : 0; }  // обращений за окно
function ptsDays(d,act){ return act<2 ? 0 : (d>180?2 : d>30?1 : 0); } // хроника — только если проблема живая

const AX_CAP = {sut:10, zhi:8, adr:6};
const ALL_MARKS = SUT_MARKS.concat(ZHI_MARKS);          // порядок = позиция бита в маске
const TAG_NAMES = {};
ALL_MARKS.forEach(m=>TAG_NAMES[m.k]=m.name);
TAG_NAMES.many='много обращений'; TAG_NAMES.long='тянется долго';

/* признаки пакуются в битовую маску — так история адреса помнит,
   что месяц назад по этому же адресу писали про червей */
function marksToMask(marks){
  let v=0; ALL_MARKS.forEach((m,i)=>{ if(marks[m.k]!==undefined) v|=(1<<i); }); return v;
}
function maskToKeys(v){
  return ALL_MARKS.filter((m,i)=>v&(1<<i)).map(m=>m.k);
}

/* ===================== ЧЕК-ЛИСТ ДЛЯ ОМСУ =====================
   Базовые пункты идут всегда, остальные подставляются под признаки
   кейса: нет смысла требовать протокол лаборатории там, где жалуются
   на давление, и график подвоза там, где вода есть.
   ============================================================= */
const CHECK_BASE = [
  'Акт обследования адреса: дата, ФИО и должность составившего, что фактически обнаружено.',
  'Причина проблемы конкретно: какой участок сети, какое оборудование, с какого года эксплуатации. Формулировка «износ сетей» без указания участка не принимается.',
  'Быстрое решение: что сделано для прекращения жалоб, в какие даты, чем подтверждается.',
  'Системное решение: мероприятие, стоимость, источник финансирования, срок ввода, включено ли в действующую программу.',
  'Информирование жителей о принятых решениях и сроках: подтверждающие материалы (объявление, сообщение в чате дома, публикация) с датой.',
  'Ответ заявителю: дата, канал, содержание ответа.'
];
const CHECK_COND = [
  {when:c=>c.marks.sanit!==undefined, title:'Санитарная угроза (черви, фекалии, канализация, запах)', items:[
    'Протокол исследования воды аккредитованной лабораторией: точка отбора, дата, показатели в сравнении с ПДК.',
    'Уведомление территориального отдела Роспотребнадзора с датой и исходящим номером.',
    'Чем жители обеспечены питьевой водой на период до устранения.'
  ]},
  {when:c=>c.marks.noaqua!==undefined, title:'Отсутствие воды более суток', items:[
    'График подвоза воды: точки размещения, объём, периодичность.',
    'Срок восстановления подачи и документальное подтверждение факта восстановления.'
  ]},
  {when:c=>c.marks.vulner!==undefined, title:'На адресе уязвимые жильцы (грудные дети, инвалиды, беременные, многодетные)', items:[
    'Адресная работа именно с этими жильцами: доставка воды, персональное информирование. Отчёт «по дому в целом» не принимается.'
  ]},
  {when:c=>c.marks.social!==undefined, title:'Затронут социальный объект (школа, детский сад, больница)', items:[
    'Отдельный акт обследования по объекту.',
    'Уведомление руководителя учреждения с датой.'
  ]},
  /* порог по заявителям — 5, а не 3: на архиве триггер «3 заявителя» давал
     требование очной встречи в 86% кейсов, то есть почти не различал */
  {when:c=>c.marks.coll!==undefined||c.marks.escal!==undefined||c.aAct>=5, title:'Коллективное обращение или эскалация', items:[
    'Очная встреча с жителями: дата, количество участников, протокол, фотоподтверждение.'
  ]},
  {when:c=>c.days>180, title:'Проблема тянется больше полугода', items:[
    'Объяснение, почему ранее принятые меры не дали результата.',
    'ФИО и должность персонально ответственного за адрес.'
  ]}
];
const VIDEO_LINK = '<p>Скачать информационный ролик по промывкам можно <a href="https://cloud.mosreg.ru/s/AgI1k7IvhaMaHY7/download?path=%2F&files=WATER%20SUPPLY.mp4">по ссылке</a></p>';

/* ===================== ВСПОМОГАТЕЛЬНОЕ ===================== */
function shortAddr(full){
  const parts=String(full||'').split(',').map(s=>s.trim()).filter(Boolean)
    .filter(p=>!/^россия$/i.test(p) && !/^московская\s+область$/i.test(p) && !/^\d{6}$/.test(p));
  const out=[];
  for(const p of parts){
    const key=p.toLowerCase().replace(/ё/g,'е');
    if(!out.length || out[out.length-1].toLowerCase().replace(/ё/g,'е')!==key) out.push(p);
  }
  return out.join(', ');
}
function addrKey(full){
  return shortAddr(full).toLowerCase().replace(/ё/g,'е')
    .replace(/[.,]/g,' ').replace(/\s+/g,' ').trim();
}
function parseAnyDate(v){
  if(v==null||v==='') return null;
  if(v instanceof Date) return isNaN(v)?null:v;
  if(typeof v==='number'){ const d=new Date(Math.round((v-25569)*86400*1000)); return isNaN(d)?null:d; }
  const s=String(v).trim();
  let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return new Date(+m[1],+m[2]-1,+m[3]);
  m=s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);   if(m) return new Date(+m[3],+m[2]-1,+m[1]);
  return null;
}
function fmtD(d){ return d? String(d.getDate()).padStart(2,'0')+'.'+String(d.getMonth()+1).padStart(2,'0')+'.'+d.getFullYear() : ''; }

const DB_EPOCH=Date.UTC(2024,0,1);
function makeDB(){
  return {
  v:1, addrs:[], mails:[], rec:{},   // rec: id -> [addrIdx, dayNum, mailIdx]
  _ai:new Map(), _mi:new Map(),
  _idx(arr,map,val){
    if(map.has(val)) return map.get(val);
    const i=arr.length; arr.push(val); map.set(val,i); return i;
  },
  day(d){ return d? Math.round((Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())-DB_EPOCH)/86400000) : 0; },
  date(n){ return new Date(DB_EPOCH+n*86400000); },

  /* возвращает, сколько обращений оказались новыми для базы */
  merge(items){
    let added=0;
    for(const it of items){
      if(!it.addr) continue;
      /* префикс '#' — чтобы ID вида __proto__/constructor не били по прототипу */
      const id='#'+(it.id || (it.addr+'|'+fmtD(it.date)+'|'+it.desc.slice(0,40)));
      if(Object.prototype.hasOwnProperty.call(this.rec,id)) continue;
      this.rec[id]=[ this._idx(this.addrs,this._ai,addrKey(it.addr)),
                     this.day(it.date),
                     it.mail? this._idx(this.mails,this._mi,it.mail) : -1,
                     marksToMask(it.marks) ];
      added++;
    }
    return added;
  },

  /* история по каждому ключу адреса: все события [день, индекс почты] */
  byAddr(){
    const m=new Map();
    for(const id in this.rec){
      const [ai,dn,mi,mk]=this.rec[id];
      const k=this.addrs[ai];
      let h=m.get(k);
      if(!h){ h={ev:[],min:Infinity,max:-Infinity}; m.set(k,h); }
      h.ev.push([dn,mi,mk||0]);
      if(dn){ if(dn<h.min) h.min=dn; if(dn>h.max) h.max=dn; }
    }
    return m;
  },
  stat(){
    let min=Infinity,max=-Infinity,n=0;
    for(const id in this.rec){ const dn=this.rec[id][1]; n++; if(dn){ if(dn<min)min=dn; if(dn>max)max=dn; } }
    return {n, addrs:this.addrs.length, from:isFinite(min)?this.date(min):null, to:isFinite(max)?this.date(max):null};
  }
  };
}


/* обращение из строки детализации: r — объект с полями
   srcnum, ecur, innid, src, type, spam, desc, omsu, addr, date, mail.
   null — спам или предложение (отсеивается, как на старой странице). */
function lowh(s){ return String(s).toLowerCase().replace(/ё/g,'е').trim(); }
function makeItem(r){
  const K={srcnum:'srcnum',ecur:'ecur',innid:'innid',src:'src',spam:'spam',type:'type',
           desc:'desc',omsu:'omsu',addr:'addr',date:'date',mail:'mail'};
    const type=String(r[K.type]).trim(), spam=lowh(r[K.spam]);
    if(type==='1' || spam==='да') return null;
    const desc=String(r[K.desc]||'');
    const marks={};
    for(const m of SUT_MARKS.concat(ZHI_MARKS)){
      const hit=desc.match(m.re);
      if(hit) marks[m.k]=hit[0].replace(/^[^а-яёa-z0-9]+|[^а-яёa-z0-9]+$/gi,'').toLowerCase().slice(0,60);
    }
    return {
      id: String(r[K.srcnum]||'').trim() || String(K.ecur?r[K.ecur]:'').trim() || String(K.innid?r[K.innid]:'').trim(),
      src: String(r[K.src]||'').trim(),
      omsu: String(r[K.omsu]||'').trim(),
      addr: String(r[K.addr]||'').trim(),
      desc, marks,
      date: K.date?parseAnyDate(r[K.date]):null,
      mail: K.mail?String(r[K.mail]||'').trim().toLowerCase():''
    };
}


/* items — обращения текущего периода; DB — база истории, в которую они уже
   подмешаны вместе с архивом. Возвращает кейсы, отсортированные по баллу. */
function buildCases(items, DB){
  let CASES;
  /* --- 3. группировка в кейсы по адресу --- */
  const byKey=new Map();
  let noAddrSeq=0;
  for(const it of items){
    const k = it.addr ? addrKey(it.addr) : '__noaddr__'+(noAddrSeq++);
    if(!byKey.has(k)) byKey.set(k,[]);
    byKey.get(k).push(it);
  }

  /* --- 4. оценка кейсов: текст берём из файла, историю — из базы --- */
  const HIST=DB.byAddr();
  CASES=[];
  for(const [k,list] of byKey){
    list.sort((a,b)=>(a.date?a.date:0)-(b.date?b.date:0));
    const last=list[list.length-1];

    /* история адреса по всей базе; для кейсов без адреса — только текущий файл */
    const h=HIST.get(k);
    let n, authors, nAct, aAct, first, latest;
    if(h){
      n=h.ev.length;
      first =isFinite(h.min)?DB.date(h.min):null;
      latest=isFinite(h.max)?DB.date(h.max):null;
      authors=new Set(h.ev.map(e=>e[1]).filter(i=>i>=0)).size;
      /* активность — только события в окне, отсчитывая от последнего */
      const act=h.ev.filter(e=>e[0] && h.max-e[0]<=ACT_WINDOW);
      nAct=act.length;
      aAct=new Set(act.map(e=>e[1]).filter(i=>i>=0)).size;
    }else{
      const dates=list.map(x=>x.date).filter(Boolean);
      n=nAct=list.length;
      authors=aAct=new Set(list.map(x=>x.mail).filter(Boolean)).size;
      first=dates.length?dates[0]:null; latest=dates.length?dates[dates.length-1]:null;
    }
    const days=(first&&latest)?Math.round((latest-first)/86400000):0;

    /* признаки кейса = признаки обращений из файла + признаки,
       встречавшиеся по этому адресу за окно активности */
    const marks={};
    for(const it of list) for(const mk in it.marks) if(!marks[mk]) marks[mk]=it.marks[mk];
    if(h){
      let hm=0;
      for(const e of h.ev) if(e[0] && h.max-e[0]<=ACT_WINDOW) hm|=e[2];
      for(const key of maskToKeys(hm)) if(marks[key]===undefined) marks[key]='из истории адреса';
    }

    let sut=0; for(const m of SUT_MARKS) if(marks[m.k]) sut+=m.w;
    let zhi=0; for(const m of ZHI_MARKS) if(marks[m.k]) zhi+=m.w;
    zhi+=ptsAuthors(aAct);
    const pc=ptsCount(nAct), pd=ptsDays(days,nAct);
    let adr=pc+pd;
    sut=Math.min(sut,AX_CAP.sut); zhi=Math.min(zhi,AX_CAP.zhi); adr=Math.min(adr,AX_CAP.adr);

    const tags=SUT_MARKS.concat(ZHI_MARKS).filter(m=>marks[m.k]).map(m=>m.k);
    if(pc) tags.push('many');
    if(pd) tags.push('long');

    CASES.push({
      key:k, sel:false, list, n, nNew:list.length, nAct, authors, aAct, first, latest, days,
      addr: last.addr, omsu: last.omsu,
      srcs:[...new Set(list.map(x=>x.src).filter(Boolean))],
      ids: list.map(x=>x.id).filter(Boolean),
      marks, tags, sut, zhi, adr, score:sut+zhi+adr,
      lastDesc:last.desc, open:false
    });
  }
  CASES.sort((a,b)=> b.score-a.score || b.n-a.n);
  return CASES;
}

  function buildDesc(c){
    let n=0, out='';
    out+='<p><b>Что требуется предоставить по адресу:</b></p><p>';
    out+=CHECK_BASE.map(s=>(++n)+'. '+s).join('<br>');
    out+='</p>';
    const blocks=CHECK_COND.filter(b=>{ try{ return b.when(c); }catch(e){ return false; } });
    for(const b of blocks){
      out+='<p><b>'+b.title+':</b><br>';
      out+=b.items.map(s=>(++n)+'. '+s).join('<br>');
      out+='</p>';
    }
    out+='<p>Если проблематика не подтверждается — расписать, почему адрес не является проблемным, со ссылкой на результаты обследования.</p>';
    out+=VIDEO_LINK;
    return out;
  }
  function caseSummary(c){
    const tg=c.tags.map(k=>TAG_NAMES[k]).join(', ');
    return '<p><b>Резонансный кейс, балл '+c.score+'</b> (суть '+c.sut+' / житель '+c.zhi+' / адрес '+c.adr+')<br>'
      +'Обращений по адресу всего: '+c.n+' (за последние 90 дней: '+c.nAct+' от '+(c.aAct||'н/д')+' разных заявителей)'
      +(c.days?(', проблема тянется '+c.days+' дн.'):'')
      +(c.first?('<br>Первое обращение: '+fmtD(c.first)+', последнее: '+fmtD(c.latest)):'')
      +'<br>Признаки: '+(tg||'—')
      +'<br>Источники: '+(c.srcs.join(', ')||'—')
      +'<br>ID обращений: '+c.ids.join(', ')+'</p>';
  }
  const REZ_RULES = {
    escal:  /телевид|(^|[^а-яё])сми([^а-яё]|$)|журналист|прокурат|(^|[^а-яё])суд|президент|губернатор|роспотребнадзор|следствен|обращ[а-яё]*\s+к\s+руководств/i,
    health: /черв|насеком|отравл|сыпь|аллерг|инфекц|зараз|дети\s+бол|опасн[а-яё]*\s+для\s+здоров|здоровь/i,
    emotion:/издеват|ужас|безобраз|невыносим|кошмар|требую|сколько\s+можно|уму\s+не\s+постижимо|умирает/i,
    chronic:/полгода|(\d+|два|три|четыре|пять)\s*(месяц|лет|год)(?!\s*(побед|октябр|влксм|ссср|съезд))|более\s+\d+\s*-?\s*х?\s*(лет|месяц)|каждый\s+(день|вечер|ночь)|ежедневн|постоянно|систематическ|до\s+сих\s+пор|бездейств|игнорир/i,
    collective:/коллектив|весь\s+дом|все\s+жильцы|жители\s+дома|подпис(ей|и|ями)/i
  };
  function rezClassify(src, text){
    const t=String(text||'');
    const b={pos:/^ПОС$|^ППМО$/i.test(String(src||'').trim()), matches:{}};
    for(const k of ['escal','health','emotion','chronic','collective']){
      const m=t.match(REZ_RULES[k]);
      b[k]=!!m;
      if(m) b.matches[k]=m[0].replace(/^[^а-яёa-z0-9]+|[^а-яёa-z0-9]+$/gi,'').toLowerCase();
    }
    return b;
  }
  const REZ_REASONS=[['pos','Федеральная платформа (ПОС/ППМО)','федеральная платформа'],
    ['escal','Угроза эскалации','эскалация'],['health','Риск здоровью','риск здоровью'],
    ['emotion','Сильные эмоции','сильные эмоции'],['chronic','Хроническая проблема','хроника'],
    ['collective','Коллективное обращение','коллектив']];
  function rezReason(o){
    const first=REZ_REASONS.find(([k])=>o[k]);
    const details=REZ_REASONS.filter(([k])=>o[k])
      .map(([k,_,low])=>k==='pos'?low:(low+': «'+(o.matches[k]||'')+'»')).join('; ');
    return {reason:first?first[1]:'', details};
  }

/* CSV для РМ: одна строка = один кейс (как кнопка «CSV для загрузки в РМ») */
function rmRows(sel, ctrl, ds, de){
  const head=['№','улица','ОМСУ','назначено','контролер ','начало ','завершение','Описание','тип','ID жалобы','текст жалобы'];
  const rows=[head];
  sel.forEach((c,i)=>{
    const texts=c.list.slice().reverse().slice(0,3)
      .map(it=>'['+fmtD(it.date)+' · '+it.src+' · ID '+it.id+'] '+it.desc).join('\n\n');
    rows.push([i+1, c.addr, c.omsu, c.omsu.toUpperCase(), ctrl, ds, de,
               buildDesc(c),
               'Резонансные обращения', c.ids[c.ids.length-1]||'', texts]);
  });
  return rows;
}

/* Excel-реестр: одна строка = одно обращение (как кнопка «Excel-реестр») */
function registryRows(sel){
  const head=['Номер обращения','Дата','ОМСУ','Адрес','Источник','Почта заявителя','Описание',
              'Причина резонанса','Детали триггера','Широта','Долгота','Категория',
              'Дата внесения','добавлено на текущей неделе'];
  const today=new Date(), aoa=[head];
  let nObr=0;
  for(const c of sel){
    const adr=shortAddr(c.addr);
    for(const it of c.list){
      const rr=rezReason(rezClassify(it.src, it.desc));
      aoa.push([ it.id, it.date||'', (it.omsu||c.omsu||'').toUpperCase(), adr, it.src, it.mail, it.desc,
                 rr.reason, rr.details, '', '', 'Резонансные жалобы', today, 0 ]);
      nObr++;
    }
  }
  return aoa;
}

const API={SUT_MARKS,ZHI_MARKS,ALL_MARKS,TAG_NAMES,ACT_WINDOW,AX_CAP,CHECK_BASE,CHECK_COND,
  toCP1251,csvField,shortAddr,addrKey,parseAnyDate,fmtD,makeDB,makeItem,buildCases,
  buildDesc,caseSummary,rezClassify,rezReason,rmRows,registryRows};
if(typeof module!=='undefined' && module.exports) module.exports=API; else root.REZ=API;
})(typeof window!=='undefined'?window:globalThis);
