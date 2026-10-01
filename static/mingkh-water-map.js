(() => {

'use strict';
/* ============================================================
   1. Данные
   ============================================================ */
const KINDS = [
  { name: 'ХВС', full: 'Холодное водоснабжение' },
  { name: 'ВО',  full: 'Водоотведение' },
  { name: 'ГВС', full: 'Горячее водоснабжение' },
  { name: 'КР',  full: 'Капремонт МКД' },
  { name: 'Др',  full: 'Другое (МинЖКХ)' },
];
const KIND_KR = 3, KIND_ETC = 4;
/* Группы жалоб — привязка факта ЕЦУР к группе дана пользователем 29.09.2026.
   Три редких факта (крышка люка ЦВС, водозаборный узел, уведомление о работах)
   в его таблице не было — отнесены по аналогии. Факт, которого здесь нет,
   попадает в «Прочее», чтобы новая формулировка портала не пропала молча. */
const GROUPS = ['Отсутствие воды', 'Ржавая вода', 'Канализация', 'Колодцы/Колонки',
                'Модернизация сетей', 'Горячее водоснабжение (ГВС) в МКД', 'Экология', 'Прочее',
                'Капремонт МКД'];
const GROUP_OF = Object.fromEntries([
  ['Восстановить работу внешней системы водоснабжения', 'Отсутствие воды'],
  ['Восстановить работу водозаборного узла', 'Отсутствие воды'],
  ['Принять меры в связи с отсутствием уведомления о запланированных или аварийных работах', 'Отсутствие воды'],
  ['Обеспечить централизованное снабжение водой, соответствующей требованиям СанПиН (ржавая вода)', 'Ржавая вода'],
  ['Принять меры в связи с превышением установленных норм по запаху от сброса стоков с объекта', 'Канализация'],
  ['Устранить повреждение на сети централизованного водоотведения', 'Канализация'],
  ['Устранить повреждение сети централизованного водоснабжения', 'Канализация'],
  ['Ввести в работу дополнительный общественный водоразборный колодец', 'Колодцы/Колонки'],
  ['Демонтировать водоразборную колонку', 'Колодцы/Колонки'],
  ['Отремонтировать водоразборную колонку', 'Колодцы/Колонки'],
  ['Отремонтировать колодец сети централизованного водоснабжения', 'Колодцы/Колонки'],
  ['Установить водоразборную колонку', 'Колодцы/Колонки'],
  ['Установить крышку на люк колодца сети водоотведения', 'Колодцы/Колонки'],
  ['Установить крышку на люк колодца сети централизованного водоснабжения', 'Колодцы/Колонки'],
  ['Устроить дополнительный водоразборный колодец', 'Колодцы/Колонки'],
  ['Модернизировать (реконструировать) канализационные очистные сооружения хозяйственно-бытовых стоков', 'Модернизация сетей'],
  ['Модернизировать (реконструировать) участок водопроводной сети', 'Модернизация сетей'],
  ['Модернизировать (реконструировать) участок сети водоотведения', 'Модернизация сетей'],
  ['Обеспечить централизованную систему горячего водоснабжения качественным ХВС, соответствующим требованиям СанПиН', 'Горячее водоснабжение (ГВС) в МКД'],
  ['Принять меры в связи со сбросом неочищенных стоков на рельеф местности', 'Экология'],
].map(([f, g]) => [normFact(f), g]));
// двойные пробелы и регистр в формулировках портала гуляют — сравниваем без них
function normFact(s){ return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
const groupOf = fact => GROUP_OF[normFact(fact)] || 'Прочее';
/* Таблица групп выше — только для воды. Капремонт — одна группа; прочие жалобы
   МинЖКХ — группа по подкатегории ЕЦУР, чтобы новая тема была видна отдельно. */
const groupOfRec = (kind, fact, subcat) => kind === KIND_KR ? 'Капремонт МКД'
  : kind === KIND_ETC ? (subcat || 'Другое (МинЖКХ)') : groupOf(fact);

const CARD_URL = id => 'https://admin.vmeste.mosreg.ru/CardEditList?show=/Topic?id=' + encodeURIComponent(id);

const esc = s => String(s ?? '').replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtN = (v, d) => Number(v).toLocaleString('ru-RU',
  { minimumFractionDigits: d || 0, maximumFractionDigits: d || 0 });
const fmtDay = s => s ? s.slice(8, 10) + '.' + s.slice(5, 7) + '.' + s.slice(0, 4) : '';
const fmtDist = m => m < 2000 ? fmtN(Math.round(m)) + ' м' : fmtN(m / 1000, m % 1000 ? 1 : 0) + ' км';
const isoDay = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' +
                    String(d.getDate()).padStart(2, '0');

let RECS = [], DMIN = '', DMAX = '';
/* Водные жалобы (ХВС, ВО, ГВС) — всегда «Инженерная инфраструктура». Портал изредка
   ставит им «Инженерные системы МКД» или «Водоснабжение, водоотведение» при тех же
   фактах (4 жалобы за 2024–2026) — отдельной карточкой это выглядело как дубль. */
const catOf = (kind, cat) => kind <= 2 ? 'Инженерная инфраструктура' : cat;
function loadData(W){
  if(!W || !W.rows) return false;
  const c = Object.fromEntries(W.cols.map((n, i) => [n, i]));
  RECS = W.rows.map(r => ({
    id: r[c.id], created: r[c.created], day: String(r[c.created]).slice(0, 10),
    omsu: W.omsu[r[c.omsu]], kind: r[c.kind], address: r[c.address],
    fact: W.fact[r[c.fact]], status: W.status[r[c.status]],
    cat: catOf(r[c.kind], c.cat !== undefined ? W.cat[r[c.cat]] : ''), org: c.org !== undefined ? W.org[r[c.org]] : '',
    group: groupOfRec(r[c.kind], W.fact[r[c.fact]], c.subcat !== undefined ? W.subcat[r[c.subcat]] : ''),
    lat: r[c.lat], lon: r[c.lon], geo: r[c.geo],
  }));
  const days = RECS.map(r => r.day).filter(Boolean).sort();
  DMIN = days[0] || ''; DMAX = days[days.length - 1] || '';
  const m = W.meta || {};
  document.getElementById('meta').textContent =
    'ДоброДел · МинЖКХ · все жалобы (вода, капремонт МКД и прочее) · ' + fmtN(RECS.length) + ' жалоб за ' +
    fmtDay(DMIN) + ' — ' + fmtDay(DMAX) + (m.updated ? ' · данные от ' + m.updated : '');
  return true;
}

/* ============================================================
   2. Фильтры
   ============================================================ */
const F = { from: '', to: '', kinds: KINDS.map(() => true) };
let geoBox = null;      // { kind:'poly', pts } | { kind:'circle', lat, lon, r } и рамка bb

/* Списки с галочками: ОМСУ и факт ЕЦУР. sel: null — все, Set — выбранные. */
const MS = {
  omsu: { field: 'omsu', all: [], sel: null, find: 'Найти округ…',
          label: { all: 'Все ОМСУ', none: 'Ни одного ОМСУ', many: n => n + ' ОМСУ' } },
  fact: { field: 'fact', all: [], sel: null, find: 'Найти по словам факта…',
          label: { all: 'Все факты', none: 'Ни одного факта', many: n => n + ' фактов из ' + MS.fact.all.length } },
  org:  { field: 'org', all: [], sel: null, find: 'Найти исполнителя…',
          label: { all: 'Все исполнители', none: 'Ни одного исполнителя', many: n => n + ' исполнителей' } },
};
/* МОВК и ФКР — те же правила, что в дашборде «Контроль жалоб ЕЦУР»:
   МОВК — округа зоны МОВК и категория «Инженерная инфраструктура»;
   ФКР — исполнитель «Фонд капитального ремонта…». Работают поверх списков. */
const MOVK_STEMS = ['богородск', 'воскресенск', 'орехово-зуевск', 'электросталь', 'лосино-петровск',
                    'павлово-посадск', 'шатур', 'чехов', 'сергиево-посадск'];
const CAT_ENG = 'Инженерная инфраструктура';
const FKR_EXEC = 'Фонд капитального ремонта общего имущества многоквартирных домов';
const isMovk = o => MOVK_STEMS.some(s => String(o).toLowerCase().includes(s));
F.movk = false; F.fkr = false;

// skip — какой список не учитывать: так у каждого пункта видно, сколько жалоб
// он даст при остальных фильтрах
function passesNoBox(r, skip){
  return (!F.from || r.day >= F.from) && (!F.to || r.day <= F.to) && F.kinds[r.kind] &&
         (skip === 'omsu' || !MS.omsu.sel || MS.omsu.sel.has(r.omsu)) &&
         (skip === 'fact' || !MS.fact.sel || MS.fact.sel.has(r.fact)) &&
         (skip === 'org' || !MS.org.sel || MS.org.sel.has(r.org)) &&
         (!F.movk || (r.cat === CAT_ENG && isMovk(r.omsu))) && (!F.fkr || r.org === FKR_EXEC) &&
         (skip === 'group' || !F.grp || F.grp.has(r.group)) &&
         (skip === 'cat' || !F.cats || F.cats.has(r.cat));
}

// ---- группы: фильтр — это сами карточки с цифрами (renderKpis)
F.grp = null;            // null — все группы; Set — включённые
F.cats = null;           // категории ЕЦУР: null — все; Set — включённые
let CATS_IN = [];
/* Те же правила щелчка, что у групп: щелчок — только эта категория, Ctrl+щелчок —
   добавить/убрать, повторный щелчок по единственной выбранной — снова все. */
function catClick(e){
  const b = e.target.closest('.catc[data-c]');
  if(!b) return;
  const c = b.dataset.c;
  let next;
  if(e.ctrlKey || e.metaKey || e.shiftKey){
    next = new Set(F.cats || []);
    next.has(c) ? next.delete(c) : next.add(c);
  } else next = F.cats && F.cats.size === 1 && F.cats.has(c) ? null : new Set([c]);
  F.cats = next && next.size && next.size < CATS_IN.length ? next : null;
  applyFilters();
}
let GROUPS_IN = [];      // группы, что реально есть в данных
function buildGroups(){
  // категории — по числу жалоб, крупные первыми
  const nc = new Map();
  for(const r of RECS) nc.set(r.cat, (nc.get(r.cat) || 0) + 1);
  CATS_IN = [...nc.keys()].sort((a, b) => nc.get(b) - nc.get(a));
  const have = new Set(RECS.map(r => r.group));
  // известные — в заданном порядке, незнакомые (подкатегории «Другого») — следом
  GROUPS_IN = GROUPS.filter(g => have.has(g)).concat([...have].filter(g => !GROUPS.includes(g)).sort());
}
/* Щелчок — только эта группа; Ctrl+щелчок — добавить/убрать; щелчок по
   единственной выбранной или по итогу — снова все. Так же, как строки свода по ОМСУ. */
function groupClick(e){
  if(e.target.closest('.kpi-total')){ if(F.grp || F.cats){ F.grp = null; F.cats = null; applyFilters(); } return; }
  const b = e.target.closest('.kpi[data-g]');
  if(!b) return;
  const g = b.dataset.g;
  let next;
  if(e.ctrlKey || e.metaKey || e.shiftKey){
    next = new Set(F.grp || []);
    next.has(g) ? next.delete(g) : next.add(g);
  } else next = F.grp && F.grp.size === 1 && F.grp.has(g) ? null : new Set([g]);
  F.grp = next && next.size && next.size < GROUPS_IN.length ? next : null;
  applyFilters();
}
const hasPoint = r => r.lat !== null && r.geo === 0;   // точка в пределах МО

function geoMetres(la1, lo1, la2, lo2){
  const R = 6371000, rad = Math.PI / 180;
  const dp = (la2 - la1) * rad, dl = (lo2 - lo1) * rad;
  const h = Math.sin(dp / 2) ** 2 + Math.cos(la1 * rad) * Math.cos(la2 * rad) * Math.sin(dl / 2) ** 2;
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(h)));
}
// луч вправо от точки: нечётное число пересечений со сторонами — точка внутри
function inPoly(la, lo, pts){
  let inside = false;
  for(let i = 0, j = pts.length - 1; i < pts.length; j = i++){
    const [ya, xa] = pts[i], [yb, xb] = pts[j];
    if((ya > la) !== (yb > la) && lo < (xb - xa) * (la - ya) / (yb - ya) + xa) inside = !inside;
  }
  return inside;
}
function shapeBB(s){
  if(s.kind === 'circle'){
    const dLa = s.r / 111320, dLo = s.r / (111320 * Math.cos(s.lat * Math.PI / 180));
    return { s: s.lat - dLa, n: s.lat + dLa, w: s.lon - dLo, e: s.lon + dLo };
  }
  const la = s.pts.map(p => p[0]), lo = s.pts.map(p => p[1]);
  return { s: Math.min(...la), n: Math.max(...la), w: Math.min(...lo), e: Math.max(...lo) };
}
function inGeoBox(r){
  if(!geoBox) return true;
  if(!hasPoint(r)) return false;       // без точки в область попасть не может
  const b = geoBox.bb;
  if(r.lat < b.s || r.lat > b.n || r.lon < b.w || r.lon > b.e) return false;
  if(geoBox.kind === 'circle') return geoMetres(r.lat, r.lon, geoBox.lat, geoBox.lon) <= geoBox.r;
  return inPoly(r.lat, r.lon, geoBox.pts);
}
const boxName = () => !geoBox ? '' :
  geoBox.kind === 'circle' ? 'Радиус ' + fmtDist(geoBox.r) + ' от точки' : 'Область на карте';

let BASE = [], SEL = [];   // BASE — по фильтрам; SEL — по фильтрам и области
function applyFilters(opts){
  const dateError = document.getElementById('load-error');
  dateError.hidden = !(F.from && F.to && F.from > F.to);
  if (!dateError.hidden) { dateError.textContent = 'Начало периода не может быть позже конца.'; return; }
  BASE = RECS.filter(r => passesNoBox(r));
  SEL = geoBox ? BASE.filter(inGeoBox) : BASE;
  renderKpis();
  msCounts('omsu'); msCounts('fact'); msCounts('org');
  document.getElementById('t-movk').classList.toggle('on', F.movk);
  document.getElementById('t-fkr').classList.toggle('on', F.fkr);
  renderPeriodNote();
  renderOmsuTable();
  renderBox();
  renderMap(opts && opts.fit);
  document.getElementById('xls').disabled = !SEL.length;
}

// ---- период
function setPeriod(from, to){
  F.from = from; F.to = to;
  for(const [id, v] of [['d-from', from], ['d-to', to]]){
    const el = document.getElementById(id);
    el.value = v;
    // подсветка — чтобы было видно, что даты сменила кнопка, а не остались прежними
    el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash');
  }
  clearQuick();
}
// на кнопке — название периода и даты: «30 дней · 31.08 — 29.09.2026»
function renderPeriodNote(){
  const on = document.querySelector('#quick .per-opt.on');
  const a = F.from, b = F.to;
  const range = !a && !b ? 'весь период'
    : a && b && a.slice(0, 4) === b.slice(0, 4) ? fmtDay(a).slice(0, 5) + ' — ' + fmtDay(b)
    : (fmtDay(a) || 'с начала') + ' — ' + (fmtDay(b) || 'по конец');
  const days = a && b ? Math.round((new Date(b) - new Date(a)) / 864e5) + 1 : 0;
  const btn = document.getElementById('per-btn');
  btn.textContent = on ? on.firstChild.textContent + ' · ' + range : range + (days > 0 ? ' · ' + fmtN(days) + ' дн.' : '');
  btn.classList.toggle('part', !on);
}
function quick(q){
  if(q === 'all') setPeriod(DMIN, DMAX);
  // «Год» — текущий календарный год с 1 января, а не 365 дней назад
  else if(q === 'year'){
    const jan1 = new Date().getFullYear() + '-01-01';
    setPeriod(jan1 < DMIN ? DMIN : jan1, DMAX);
  } else {
    const to = new Date(DMAX + 'T00:00:00'), from = new Date(to);
    from.setDate(from.getDate() - (+q - 1));
    setPeriod(isoDay(from) < DMIN ? DMIN : isoDay(from), DMAX);
  }
  document.querySelector('#quick [data-q="' + q + '"]').classList.add('on');
}

// ---- списки с галочками (ОМСУ, факт)
const msBox = key => document.querySelector('.ms[data-ms="' + key + '"]');
function msBuild(key){
  const m = MS[key], box = msBox(key);
  const n = new Map();
  for(const r of RECS) n.set(r[m.field], (n.get(r[m.field]) || 0) + 1);
  // ОМСУ — по алфавиту; факты — по частоте: вверху то, о чём пишут чаще всего
  m.all = [...n.keys()].sort(key !== 'omsu' ? (a, b) => n.get(b) - n.get(a) : (a, b) => a.localeCompare(b, 'ru'));
  box.innerHTML =
    '<button class="ms-btn" type="button"></button>' +
    '<div class="ms-pop" hidden><input class="inp ms-q" placeholder="' + esc(m.find) + '" autocomplete="off">' +
    '<div class="ms-acts"><button class="linkbtn" data-a="all" type="button">Выбрать все</button>' +
    '<button class="linkbtn" data-a="none" type="button">Снять все</button></div>' +
    '<div class="ms-list">' + m.all.map(o =>
      '<label class="ms-opt" data-o="' + esc(o) + '" title="' + esc(o) + '"><input type="checkbox" checked>' +
      '<span>' + (esc(o) || '<span class="mute">не указан</span>') + '</span><span class="n"></span></label>').join('') +
    '</div></div>';
  const btn = box.querySelector('.ms-btn'), pop = box.querySelector('.ms-pop');
  btn.onclick = () => {
    document.querySelectorAll('.ms-pop').forEach(p => { if(p !== pop) p.hidden = true; });
    pop.hidden = !pop.hidden;
    if(!pop.hidden) box.querySelector('.ms-q').focus();
  };
  box.querySelector('.ms-q').oninput = e => {
    const q = e.target.value.trim().toLowerCase();
    box.querySelectorAll('.ms-opt').forEach(l => l.classList.toggle('hide', !!q && !l.dataset.o.toLowerCase().includes(q)));
  };
  box.querySelector('.ms-list').onchange = () => msRead(key);
  box.querySelectorAll('[data-a]').forEach(b => b.onclick = () => {
    // «все/снять» — по видимым после поиска: нашёл «Люберцы», снял всё, поставил одну
    const on = b.dataset.a === 'all';
    box.querySelectorAll('.ms-opt:not(.hide) input').forEach(i => i.checked = on);
    msRead(key);
  });
  msLabel(key);
}
function msLabel(key){
  const m = MS[key], b = msBox(key).querySelector('.ms-btn');
  b.textContent = !m.sel ? m.label.all : !m.sel.size ? m.label.none
                : m.sel.size === 1 ? [...m.sel][0] : m.label.many(m.sel.size);
  b.title = m.sel ? [...m.sel].join('\n') : '';
  b.classList.toggle('part', !!m.sel);
}
function msRead(key){
  const on = [...msBox(key).querySelectorAll('.ms-opt')].filter(l => l.querySelector('input').checked).map(l => l.dataset.o);
  msSet(key, new Set(on));
}
// Выбор со стороны (из свода по ОМСУ): галочки в списке ставим те же
function msSet(key, set){
  const m = MS[key];
  m.sel = set && set.size < m.all.length ? set : null;
  msBox(key).querySelectorAll('.ms-opt').forEach(l => l.querySelector('input').checked = !m.sel || m.sel.has(l.dataset.o));
  msLabel(key);
  applyFilters({ fit: key === 'omsu' });
}
// сколько жалоб даст каждый пункт при остальных фильтрах — чтобы выбирать осмысленно
function msCounts(key){
  const m = MS[key], n = new Map();
  for(const r of RECS) if(passesNoBox(r, key)) n.set(r[m.field], (n.get(r[m.field]) || 0) + 1);
  msBox(key).querySelectorAll('.ms-opt').forEach(l => l.querySelector('.n').textContent = fmtN(n.get(l.dataset.o) || 0));
}
const msText = key => !MS[key].sel ? 'все' : [...MS[key].sel].join('; ') || 'ни одного';

// ---- свод по ОМСУ
/* Считаем по всем фильтрам, кроме самого выбора ОМСУ, и с учётом области:
   таблица — это ещё и переключатель, в ней должны быть видны и те округа,
   что сейчас не выбраны. */
let OT_SORT = { key: 'n', asc: false };
function renderOmsuTable(){
  const by = new Map();
  for(const r of RECS){
    if(!passesNoBox(r, 'omsu') || !inGeoBox(r)) continue;
    let o = by.get(r.omsu);
    if(!o) by.set(r.omsu, o = { omsu: r.omsu, n: 0 });
    o.n++;
  }
  const sel = MS.omsu.sel;
  // выбранные округа держим в таблице даже с нулём — иначе их нельзя снять щелчком
  if(sel) for(const o of sel) if(!by.has(o)) by.set(o, { omsu: o, n: 0 });
  const rows = [...by.values()];
  const { key, asc } = OT_SORT;
  rows.sort((a, b) => key === 'omsu' ? a.omsu.localeCompare(b.omsu, 'ru') * (asc ? 1 : -1)
                                     : (asc ? a[key] - b[key] : b[key] - a[key]) || a.omsu.localeCompare(b.omsu, 'ru'));
  const total = rows.reduce((s, o) => s + o.n, 0);
  const max = Math.max(1, ...rows.map(o => o.n));
  const num = v => v ? fmtN(v) : '<span class="z">—</span>';
  const pct = v => total ? (v / total * 100).toFixed(v / total < .01 ? 2 : 1).replace('.', ',') + ' %' : '';
  const th = (s, t, cls) => '<th data-s="' + s + '"' +
    ' class="' + (cls || '') + (s === key ? ' sorted' + (asc ? ' asc' : '') : '') + '">' + t + '</th>';
  const head = '<thead><tr><th class="no">№</th>' + th('omsu', 'ОМСУ', 'l') + th('n', 'Жалоб') +
    '<th class="l" data-s="n">Доля</th></tr></thead>';
  const tr = (o, i) =>
    '<tr data-o="' + esc(o.omsu) + '" class="' + (sel && sel.has(o.omsu) ? 'on' : '') + (o.n ? '' : ' zero') + '">' +
    '<td class="no">' + (i + 1) + '</td><td class="l">' + esc(o.omsu) + '</td><td class="n">' + num(o.n) + '</td>' +
    '<td class="l"><div class="share"><b style="width:' + Math.round(o.n / max * 180) + 'px"></b><span>' + pct(o.n) + '</span></div></td></tr>';
  // две колонки: первая половина слева, вторая справа, нумерация сквозная
  const half = Math.ceil(rows.length / 2);
  const part = (a, off) => '<div><table class="ot">' + head + '<tbody>' + a.map((o, i) => tr(o, i + off)).join('') + '</tbody></table></div>';
  document.getElementById('ot').innerHTML = rows.length
    ? part(rows.slice(0, half), 0) + (rows.length > 1 ? part(rows.slice(half), half) : '')
    : '<div class="mute" style="padding:10px 12px">По этим фильтрам жалоб нет</div>';
  document.getElementById('ot-total').innerHTML = 'Итого: <span><b>' + fmtN(total) + '</b> жалоб в <b>' +
    fmtN(rows.filter(o => o.n).length) + '</b> ОМСУ</span>';
  document.getElementById('ot-note').textContent = (geoBox ? 'в выделенной области · ' : '') +
    'щелчок по строке — только этот округ, Ctrl+щелчок — добавить к выбору, повторный щелчок — снять';
  document.getElementById('ot-reset').hidden = !sel;
}
function omsuTableClick(e){
  const tr = e.target.closest('tr[data-o]');
  if(!tr) return;
  const o = tr.dataset.o, sel = MS.omsu.sel;
  let next;
  if(e.ctrlKey || e.metaKey || e.shiftKey){
    next = new Set(sel || []);
    if(!sel) next.clear();
    next.has(o) ? next.delete(o) : next.add(o);
    if(!next.size) next = null;                          // сняли последний — снова все
  } else {
    next = sel && sel.size === 1 && sel.has(o) ? null : new Set([o]);   // повторный щелчок — снять
  }
  msSet('omsu', next || new Set(MS.omsu.all));
}

/* ============================================================
   3. Цифры
   ============================================================ */
function renderKpis(){
  /* Число у группы — без учёта выбора самих групп (но с областью и прочими
     фильтрами): невыбранная карточка показывает, сколько она добавит, а не ноль. */
  const by = new Map();
  for(const r of RECS)
    if(passesNoBox(r, 'group') && inGeoBox(r)) by.set(r.group, (by.get(r.group) || 0) + 1);
  let noPt = 0;
  for(const r of SEL) if(!hasPoint(r)) noPt++;
  const k = (v, t, tip) =>
    '<div class="kpi"' + (tip ? ' title="' + esc(tip) + '"' : '') + '><div class="v">' +
    fmtN(v) + '</div><div class="t">' + t + '</div></div>';
  const cls = g => !F.grp ? '' : F.grp.has(g) ? ' on' : ' off';
  // категории: число — без учёта выбора самих категорий, как у групп
  const bc = new Map();
  for(const r of RECS)
    if(passesNoBox(r, 'cat') && inGeoBox(r)) bc.set(r.cat, (bc.get(r.cat) || 0) + 1);
  const ccls = c => !F.cats ? '' : F.cats.has(c) ? ' on' : ' off';
  const catsShown = CATS_IN.slice().sort((a, b) => (bc.get(b) || 0) - (bc.get(a) || 0) || CATS_IN.indexOf(a) - CATS_IN.indexOf(b));
  document.getElementById('cats').innerHTML = CATS_IN.length < 2 ? '' : catsShown.map(c =>
    '<button type="button" class="catc' + ccls(c) + '" data-c="' + esc(c) + '"><div class="t">' +
    (esc(c) || 'Категория не указана') + '</div><div class="v">' + fmtN(bc.get(c) || 0) + '</div></button>').join('');
  // при выбранной категории группы с нулём из другой категории только мешают
  // от большего к меньшему по текущим цифрам; при равенстве — в заданном порядке
  const groupsShown = (F.cats ? GROUPS_IN.filter(g => by.get(g) || (F.grp && F.grp.has(g))) : GROUPS_IN.slice())
    .sort((a, b) => (by.get(b) || 0) - (by.get(a) || 0) || GROUPS_IN.indexOf(a) - GROUPS_IN.indexOf(b));
  document.getElementById('kpis').innerHTML =
    '<div class="kpi kpi-total" role="button" tabindex="0" title="' + (F.grp || F.cats ? 'Щелчок — вернуть все категории и группы' : '') + '"><div class="v">' +
    fmtN(SEL.length) + '</div><div class="t">' +
    (geoBox ? 'жалоб в выделенной области' : 'жалоб по фильтрам') + '</div></div>' +
    groupsShown.map(g => '<div class="kpi' + cls(g) + '" role="button" tabindex="0" data-g="' + esc(g) + '"><div class="v">' +
      fmtN(by.get(g) || 0) + '</div><div class="t">' + esc(g) + '</div></div>').join('') +
    (geoBox ? '' : k(noPt, 'без точки на карте',
      'Точки нет в карточке, она вне Московской области или координаты ещё не собраны. ' +
      'В выгрузку такие жалобы попадают, на карту — нет.'));
}

function renderBox(){
  document.getElementById('box-chip').hidden = !geoBox;
  document.getElementById('box-name').textContent = boxName() + ' · ' + fmtN(SEL.length) + ' жалоб';
}

/* ============================================================
   4. Карта
   ============================================================ */
const LEAF_JS = ['/static/vendor/leaflet.min.js',
                 '/static/vendor/leaflet-heat.js',
                 '/static/vendor/leaflet.markercluster.min.js'];
/* Заблокированный сервер не отвечает ошибкой, он просто молчит — без таймаута
   карта висела бы пустой без единого слова. */
const WAIT = 12000;
const withTimeout = (p, ms, what) => Promise.race([p,
  new Promise((_, rej) => setTimeout(() => rej(new Error('нет ответа: ' + what)), ms))]);
const loadJs = src => withTimeout(new Promise((res, rej) => {
  const t = document.createElement('script');
  t.src = src; t.onload = res; t.onerror = () => rej(new Error('не загрузился: ' + src));
  document.head.appendChild(t);
}), WAIT, src);

/* Подложка — 2ГИС. OpenStreetMap из сети ЕДДС отдаёт заглушку «Access blocked»,
   зарубежные серверы без VPN недоступны, а Яндекс в эллиптическом Меркаторе
   сдвигает точки: жалоба встала бы не у своего дома. */
const TILE = { url: 'https://tile{s}.maps.2gis.com/tiles?x={x}&y={y}&z={z}',
               opts: { subdomains: '0123', maxZoom: 18, attribution: '© 2ГИС' } };

let MAP = null, LAYER = null, BOXLAYER = null, DRAW = null, tileNote = '';
let MAPKIND = 'heat';
let mapSel = null, polyPts = null;
let mapRadius = 500;

async function initMap(){
  try{ for(const s of LEAF_JS) await loadJs(s); }
  catch(e){
    document.getElementById('map').innerHTML = '<div class="map-msg"><div><b>Карта не загрузилась.</b><br>' +
      esc(e.message) + '<br><span class="mute">Фильтры, цифры и выгрузка в Excel работают и без неё.</span></div></div>';
    return;
  }
  MAP = L.map('map', { preferCanvas: true }).setView([55.75, 37.6], 8);
  MAP.attributionControl.setPrefix('');
  const tl = L.tileLayer(TILE.url, TILE.opts).addTo(MAP);
  let ok = 0;
  const watch = setTimeout(() => { if(!ok){ tileNote = 'Подложка 2ГИС не отвечает — точки есть, фона под ними нет.'; renderStatus(); } }, WAIT);
  tl.on('tileload', () => { if(!ok++){ clearTimeout(watch); if(tileNote){ tileNote = ''; renderStatus(); } } });
  addSelectTools();
  renderMap(true);
}

const heatGradient = { 0.2: '#3a4fa0', 0.4: '#1fa3c4', 0.55: '#2ecc71', 0.7: '#f1c40f',
                       0.8: '#f76707', 0.9: '#e03131', 1: '#7f0000' };
const CL_LIM = [20, 100];
function clusterIcon(n){
  const cls = n >= CL_LIM[1] ? 'cl-l' : n >= CL_LIM[0] ? 'cl-m' : 'cl-s';
  const d = n >= CL_LIM[1] ? 46 : n >= CL_LIM[0] ? 40 : 34;
  return L.divIcon({ html: '<div class="cl-ico ' + cls + '" style="width:' + d + 'px;height:' + d + 'px">' +
    (n < 1000 ? n : (n / 1000).toFixed(n < 10000 ? 1 : 0).replace('.', ',') + 'к') + '</div>',
    className: '', iconSize: [d, d] });
}
/* Одиночная жалоба — такой же кружок с числом, что и группа, только «1»:
   голая точка рядом с подписанными кружками читалась как что-то другое.
   Цвет один для всех — как у малых групп (разделение по видам убрано 01.10.2026). */
let DOT = null;
const dotIcon = () => DOT || (DOT = L.divIcon({ className: '', iconSize: [28, 28],
  html: '<div class="cl-ico cl-one cl-s" style="width:28px;height:28px">1</div>' }));

function popup(r){
  const row = (k, v) => '<tr><th>' + k + '</th><td>' + v + '</td></tr>';
  return '<div class="pop"><div class="pop-h">Жалоба № <a href="' + CARD_URL(r.id) + '" target="_blank" rel="noopener noreferrer">' +
    r.id + '</a></div><div class="pop-sub">' + esc(r.omsu) + ' · ' + KINDS[r.kind].full + '</div><table class="pop">' +
    row('Подана', esc(fmtDay(r.day) + r.created.slice(10))) +
    row('Адрес', esc(r.address) || '—') +
    row('Категория', esc(r.cat) || '—') +
    row('Группа', esc(r.group)) +
    row('Факт', esc(r.fact) || '—') +
    row('Исполнитель', esc(r.org) || '—') +
    row('Статус', esc(r.status) || '—') +
    row('Координаты', r.lat.toFixed(6) + ', ' + r.lon.toFixed(6)) +
    '</table></div>';
}

let MAP_DIRTY = false;
function renderMap(fit){
  if(!MAP) return renderStatus();
  // вкладка карты скрыта: контейнер нулевой, тепловой слой упадёт на холсте
  // нулевой ширины — перерисуем при возврате на вкладку
  if(!document.getElementById('map').offsetWidth){ MAP_DIRTY = true; return renderStatus(); }
  MAP_DIRTY = false;
  if(LAYER){ MAP.removeLayer(LAYER); LAYER = null; }
  // на карте — вся выборка по фильтрам, а выделение рисуется поверх: так видно,
  // что осталось за границей области
  const pts = BASE.filter(hasPoint);
  if(MAPKIND === 'heat'){
    LAYER = L.heatLayer(pts.map(r => [r.lat, r.lon, 1]),
      { radius: 14, blur: 16, maxZoom: 13, minOpacity: .35, gradient: heatGradient });
  } else {
    LAYER = L.markerClusterGroup({ chunkedLoading: true, showCoverageOnHover: false, maxClusterRadius: 50,
      iconCreateFunction: cl => clusterIcon(cl.getChildCount()) });
    LAYER.addLayers(pts.map(r => L.marker([r.lat, r.lon], { icon: dotIcon() }).bindPopup(() => popup(r))));
  }
  LAYER.addTo(MAP);
  drawBox();
  if(fit && pts.length){
    let s = 90, n = -90, w = 180, e = -180;
    for(const r of pts){ if(r.lat < s) s = r.lat; if(r.lat > n) n = r.lat; if(r.lon < w) w = r.lon; if(r.lon > e) e = r.lon; }
    MAP.fitBounds([[s, w], [n, e]], { padding: [20, 20], maxZoom: 15 });
  }
  renderStatus();
}

function drawBox(){
  if(BOXLAYER){ MAP.removeLayer(BOXLAYER); BOXLAYER = null; }
  if(!geoBox) return;
  const st = { color: '#2a78d6', weight: 2, dashArray: '6,5', fillOpacity: 0.06 };
  BOXLAYER = geoBox.kind === 'circle'
    ? L.layerGroup([L.circle([geoBox.lat, geoBox.lon], Object.assign({ radius: geoBox.r }, st)),
                    L.circleMarker([geoBox.lat, geoBox.lon], { radius: 5, color: '#2a78d6', fillColor: '#fff', fillOpacity: 1, weight: 2 })])
    : L.polygon(geoBox.pts, st);
  BOXLAYER.addTo(MAP);
}

function renderStatus(){
  const el = document.getElementById('map-status');
  const onMap = BASE.filter(hasPoint).length;
  const out = BASE.filter(r => r.geo === 1).length, none = BASE.filter(r => r.geo === 2).length,
        wait = BASE.filter(r => r.geo === 3).length;
  let s = 'На карте <b>' + fmtN(onMap) + '</b> из ' + fmtN(BASE.length) + ' жалоб по фильтрам';
  const miss = [];
  if(none) miss.push(fmtN(none) + ' без точки в карточке');
  if(out) miss.push(fmtN(out) + ' с точкой вне МО');
  if(wait) miss.push(fmtN(wait) + ' — координаты ещё не собраны');
  if(miss.length) s += ' <span class="mute">(' + miss.join(', ') + ')</span>';
  if(geoBox) s += ' · в выделении <b>' + fmtN(SEL.length) + '</b>';
  if(mapSel === 'poly') s += '<br><span class="map-hint">Обводка: щелчками ставьте вершины' +
    (polyPts && polyPts.length ? ' (поставлено ' + polyPts.length + ')' : '') +
    '. Замкнуть — щелчок по первой вершине, двойной щелчок или Enter. Backspace — убрать последнюю, Esc — отмена.</span>';
  if(mapSel === 'circle') s += '<br><span class="map-hint">Щёлкните в точку — выделится всё в радиусе ' + fmtDist(mapRadius) + '.</span>';
  if(tileNote) s += '<br><span class="map-bad">' + esc(tileNote) + '</span>';
  const dot = (c, t) => '<span><i style="background:' + c + '"></i>' + t + '</span>';
  s += MAPKIND === 'heat'
    ? '<div class="lg"><span>Цвет — плотность жалоб:</span><span><i class="lg-grad"></i>реже → чаще</span></div>'
    : '<div class="lg"><span>Число в кружке — сколько жалоб:</span>' + dot('#3987e5', 'до ' + CL_LIM[0]) +
      dot('#e0892a', CL_LIM[0] + ' — ' + (CL_LIM[1] - 1)) + dot('#cc4444', 'от ' + CL_LIM[1]) + '</div>';
  el.innerHTML = s;
}

/* Выделение на карте — как в дашборде ЕДДС:
     «Область» — многоугольник по вершинам: прямоугольник цеплял бы соседние
                 посёлки и дома за рекой;
     «Радиус»  — точка и всё в пределах N метров от неё. */
function setMode(mode){
  mapSel = mapSel === mode ? null : mode;
  document.querySelectorAll('.map-seltools [data-sel]').forEach(a => a.classList.toggle('on', a.dataset.sel === mapSel));
  MAP.getContainer().style.cursor = mapSel ? 'crosshair' : '';
  // Пока выделяем, точки и кружки не ловят щелчки: иначе щелчок по кружку
  // приближает карту вместо того, чтобы поставить вершину
  MAP.getContainer().classList.toggle('selecting', !!mapSel);
  if(mapSel === 'poly') MAP.doubleClickZoom.disable(); else MAP.doubleClickZoom.enable();
  if(mapSel !== 'poly') polyStop();
  renderStatus();
}
function addSelectTools(){
  const ctl = L.control({ position: 'topright' });
  ctl.onAdd = () => {
    const d = L.DomUtil.create('div', 'leaflet-bar map-seltools');
    // ширину — в самом элементе: у Leaflet в .leaflet-bar a она жёстко 26px,
    // и подпись иначе обрезается картой
    const st = ' style="width:auto;min-width:0;padding:0 12px"';
    d.innerHTML = '<a href="#" data-sel="poly" class="map-selbtn"' + st + ' title="Обвести многоугольник по вершинам">⬠ Область</a>' +
                  '<a href="#" data-sel="circle" class="map-selbtn"' + st + ' title="Поставить точку — всё в пределах радиуса">◎ Радиус</a>';
    L.DomEvent.disableClickPropagation(d);
    d.querySelectorAll('[data-sel]').forEach(a => a.onclick = e => { L.DomEvent.preventDefault(e); setMode(a.dataset.sel); });
    return d;
  };
  ctl.addTo(MAP);

  MAP.on('click', e => {
    if(mapSel === 'circle'){ setMode(null); setBox({ kind: 'circle', lat: e.latlng.lat, lon: e.latlng.lng, r: mapRadius }); return; }
    if(mapSel !== 'poly') return;
    if(!polyPts) polyPts = [];
    if(polyPts.length >= 3 && MAP.latLngToContainerPoint(polyPts[0]).distanceTo(e.containerPoint) < 12) return polyFinish();
    // двойной щелчок присылает ещё и два обычных — вторую вершину в ту же точку не ставим
    const last = polyPts[polyPts.length - 1];
    if(last && MAP.latLngToContainerPoint(last).distanceTo(e.containerPoint) < 6) return;
    polyPts.push(e.latlng);
    polyDraw(e.latlng);
    renderStatus();
  });
  MAP.on('mousemove', e => { if(mapSel === 'poly' && polyPts && polyPts.length) polyDraw(e.latlng); });
  MAP.on('dblclick', e => { if(mapSel === 'poly'){ L.DomEvent.stop(e); polyFinish(); } });
  document.addEventListener('keydown', e => {
    if(mapSel !== 'poly' || /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
    if(e.key === 'Escape') setMode(null);
    else if(e.key === 'Enter') polyFinish();
    else if(e.key === 'Backspace' && polyPts && polyPts.length){
      e.preventDefault(); polyPts.pop(); polyDraw(); renderStatus();
    }
  });
}
function polyDraw(cursor){
  if(!DRAW) DRAW = L.layerGroup().addTo(MAP);
  DRAW.clearLayers();
  if(!polyPts || !polyPts.length) return;
  const st = { color: '#2a78d6', weight: 2 };
  if(polyPts.length >= 2) L.polygon(polyPts, Object.assign({ fillOpacity: 0.08, weight: 0 }, st)).addTo(DRAW);
  L.polyline(polyPts, st).addTo(DRAW);
  if(cursor) L.polyline([polyPts[polyPts.length - 1], cursor], Object.assign({ dashArray: '4,5', opacity: .7 }, st)).addTo(DRAW);
  polyPts.forEach((p, i) => L.circleMarker(p, { radius: i ? 4 : 6, color: '#2a78d6', weight: 2,
    fillColor: '#fff', fillOpacity: 1, interactive: false }).addTo(DRAW));
}
function polyStop(){ polyPts = null; if(DRAW){ MAP.removeLayer(DRAW); DRAW = null; } }
function polyFinish(){
  const pts = (polyPts || []).map(p => [p.lat, p.lng]);
  setMode(null);
  if(pts.length < 3) return;          // меньше трёх вершин — не область
  setBox({ kind: 'poly', pts });
}
function setBox(shape){
  shape.bb = shapeBB(shape);
  geoBox = shape;
  applyFilters();
}

/* ============================================================
   5. Выгрузка в Excel — xlsx собираем сами, без библиотек
   ============================================================ */
const CRC_T = (() => { const t = new Uint32Array(256);
  for(let n = 0; n < 256; n++){ let c = n; for(let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
  return t; })();
function crc32(u8){ let c = 0xFFFFFFFF; for(let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
async function makeZip(entries){
  const enc = new TextEncoder(), parts = [], cdir = [];
  let off = 0;
  for(const e of entries){
    const nameB = enc.encode(e.name), orig = enc.encode(e.text), crc = crc32(orig);
    let data = orig, method = 0;
    if(typeof CompressionStream !== 'undefined'){
      try{
        const z = new Uint8Array(await new Response(new Blob([orig]).stream()
          .pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer());
        if(z.length < orig.length){ data = z; method = 8; }
      }catch(_){ /* остаётся без сжатия */ }
    }
    const lh = new Uint8Array(30 + nameB.length), dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(8, method, true);
    dv.setUint32(14, crc, true); dv.setUint32(18, data.length, true); dv.setUint32(22, orig.length, true);
    dv.setUint16(26, nameB.length, true); lh.set(nameB, 30);
    parts.push(lh, data);
    cdir.push({ nameB, crc, comp: data.length, orig: orig.length, off, method });
    off += lh.length + data.length;
  }
  let cdLen = 0;
  for(const c of cdir){
    const h = new Uint8Array(46 + c.nameB.length), dv = new DataView(h.buffer);
    dv.setUint32(0, 0x02014b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 20, true);
    dv.setUint16(10, c.method, true); dv.setUint32(16, c.crc, true); dv.setUint32(20, c.comp, true);
    dv.setUint32(24, c.orig, true); dv.setUint16(28, c.nameB.length, true); dv.setUint32(42, c.off, true);
    h.set(c.nameB, 46); parts.push(h); cdLen += h.length;
  }
  const eo = new Uint8Array(22), dv = new DataView(eo.buffer);
  dv.setUint32(0, 0x06054b50, true); dv.setUint16(8, cdir.length, true); dv.setUint16(10, cdir.length, true);
  dv.setUint32(12, cdLen, true); dv.setUint32(16, off, true);
  parts.push(eo);
  return new Blob(parts, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
const XE = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
                               .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
function colLetter(n){ let s = ''; while(n > 0){ const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; }
// «ГГГГ-ММ-ДД ЧЧ:ММ» → число Excel: дата остаётся датой, по ней работают фильтры и сортировка
function serial(s){
  const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?/.exec(s || '');
  return m ? Date.UTC(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) / 86400000 + 25569 : null;
}
const XS = { def: 0, head: 1, text: 2, int: 3, dt: 4, coord: 5, key: 6 };
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="DD.MM.YYYY HH:MM"/><numFmt numFmtId="165" formatCode="0.000000"/></numFmts>
<fonts count="3"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="10"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1C5CAB"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD0D0CC"/></left><right style="thin"><color rgb="FFD0D0CC"/></right><top style="thin"><color rgb="FFD0D0CC"/></top><bottom style="thin"><color rgb="FFD0D0CC"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

function sheetXml(cols, rows, opts){
  const cell = (ref, v, st) => {
    if(v === null || v === undefined || v === '') return '<c r="' + ref + '" s="' + st + '"/>';
    if(typeof v === 'number') return '<c r="' + ref + '" s="' + st + '"><v>' + v + '</v></c>';
    return '<c r="' + ref + '" s="' + st + '" t="inlineStr"><is><t xml:space="preserve">' + XE(v) + '</t></is></c>';
  };
  let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  if(opts.freeze) x += '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>';
  x += '<cols>' + cols.map((c, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + c.w + '" customWidth="1"/>').join('') + '</cols><sheetData>';
  if(opts.head){
    x += '<row r="1">' + cols.map((c, i) => cell(colLetter(i + 1) + '1', c.h, XS.head)).join('') + '</row>';
  }
  const r0 = opts.head ? 2 : 1;
  rows.forEach((r, j) => {
    x += '<row r="' + (r0 + j) + '">' + r.map((v, i) => cell(colLetter(i + 1) + (r0 + j), v,
      Array.isArray(cols[i].s) ? cols[i].s[j] : cols[i].s)).join('') + '</row>';
  });
  x += '</sheetData>';
  if(opts.head && rows.length) x += '<autoFilter ref="A1:' + colLetter(cols.length) + (rows.length + 1) + '"/>';
  return x + '</worksheet>';
}

async function exportXlsx(){
  const btn = document.getElementById('xls');
  btn.disabled = true; btn.textContent = 'Собираю…';
  try{
    const recs = [...SEL].sort((a, b) => a.created < b.created ? -1 : a.created > b.created ? 1 : a.id - b.id);
    const cols = [
      { h: 'Номер жалобы', w: 13, s: XS.int }, { h: 'Дата создания', w: 16, s: XS.dt },
      { h: 'Адрес', w: 60, s: XS.text }, { h: 'Широта', w: 12, s: XS.coord }, { h: 'Долгота', w: 12, s: XS.coord },
      { h: 'Категория (ЕЦУР)', w: 34, s: XS.text }, { h: 'Факт (ЕЦУР)', w: 60, s: XS.text }, { h: 'Группа', w: 24, s: XS.text },
      { h: 'ОМСУ', w: 22, s: XS.text }, { h: 'Исполнитель', w: 40, s: XS.text }, { h: 'Вид', w: 22, s: XS.text },
      { h: 'Статус', w: 22, s: XS.text }, { h: 'Координаты', w: 30, s: XS.text },
    ];
    const coordNote = r => r.geo === 0 ? '' : r.geo === 1 ? 'точка вне Московской области'
                         : r.geo === 2 ? 'в карточке нет точки' : 'ещё не собраны';
    const rows = recs.map(r => [r.id, serial(r.created), r.address,
      r.lat !== null ? r.lat : null, r.lon !== null ? r.lon : null, r.cat, r.fact, r.group, r.omsu, r.org, KINDS[r.kind].full, r.status, coordNote(r)]);

    const cond = [
      ['Выгрузка', 'Жалобы ДоброДела (МинЖКХ)'],
      ['Сформировано', new Date().toLocaleString('ru-RU')],
      ['Период подачи', fmtDay(F.from) + ' — ' + fmtDay(F.to)],
      ['ОМСУ', msText('omsu')],
      ['Факт (ЕЦУР)', msText('fact')],
      ['Исполнитель', msText('org')],
      ['МОВК', F.movk ? 'да — округа МОВК, категория «' + CAT_ENG + '»' : 'нет'],
      ['ФКР', F.fkr ? 'да — исполнитель «' + FKR_EXEC + '»' : 'нет'],
      ['Категория (ЕЦУР)', F.cats ? [...F.cats].join('; ') : 'все'],
      ['Группа', F.grp ? [...F.grp].join('; ') : 'все'],
      ['Область на карте', geoBox ? (geoBox.kind === 'circle'
        ? 'радиус ' + fmtDist(geoBox.r) + ' от точки ' + geoBox.lat.toFixed(6) + ', ' + geoBox.lon.toFixed(6)
        : 'многоугольник, вершины: ' + geoBox.pts.map(p => p[0].toFixed(6) + ' ' + p[1].toFixed(6)).join('; '))
        : 'не выделена'],
      ['Жалоб в выгрузке', rows.length],
    ];
    const s1 = sheetXml(cols, rows, { head: true, freeze: true });
    const s2 = sheetXml([{ w: 22, s: XS.key }, { w: 90, s: XS.def }], cond, {});
    const blob = await makeZip([
      { name: '[Content_Types].xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Жалобы" sheetId="1" r:id="rId1"/><sheet name="Условия" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'Жалобы\'!$A$1:$' + colLetter(cols.length) + '$' + (rows.length + 1) + '</definedName></definedNames></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', text: STYLES },
      { name: 'xl/worksheets/sheet1.xml', text: s1 },
      { name: 'xl/worksheets/sheet2.xml', text: s2 },
    ]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'Жалобы_МинЖКХ_' + F.from + '_' + F.to + (geoBox ? '_область' : '') + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } finally {
    btn.textContent = '⬇ Скачать Excel'; btn.disabled = !SEL.length;
  }
}

/* ============================================================
   6. Запуск
   ============================================================ */
function wire(){
  const perBox = document.getElementById('per'), perPop = perBox.querySelector('.ms-pop');
  document.getElementById('per-btn').onclick = () => {
    document.querySelectorAll('.ms-pop').forEach(p => { if(p !== perPop) p.hidden = true; });
    perPop.hidden = !perPop.hidden;
  };
  document.getElementById('d-from').onchange = e => { F.from = e.target.value; clearQuick(); applyFilters(); };
  document.getElementById('d-to').onchange = e => { F.to = e.target.value; clearQuick(); applyFilters(); };
  // даты, заданные кнопкой, должны стоять в полях всегда — даже если браузер
  // сбросил поле при открытии календаря
  for(const id of ['d-from', 'd-to']) document.getElementById(id).onblur = e => {
    const want = id === 'd-from' ? F.from : F.to;
    if(e.target.value !== want) e.target.value = want;
  };
  // готовый период выбран — список закрываем; свои даты правятся при открытом
  document.querySelectorAll('#quick .per-opt').forEach(b => b.onclick = () => { quick(b.dataset.q); perPop.hidden = true; applyFilters(); });
  document.getElementById('kpis').onclick = groupClick;
  document.getElementById('cats').onclick = catClick;
  document.getElementById('t-movk').onclick = () => { F.movk = !F.movk; applyFilters({ fit: true }); };
  document.getElementById('t-fkr').onclick = () => { F.fkr = !F.fkr; applyFilters({ fit: true }); };

  // щелчок мимо списка закрывает его
  document.addEventListener('click', e => document.querySelectorAll('.ms').forEach(box => {
    if(!box.contains(e.target)) box.querySelector('.ms-pop').hidden = true;
  }));

  document.querySelectorAll('#mk button').forEach(b => b.onclick = () => {
    MAPKIND = b.dataset.mk;
    document.querySelectorAll('#mk button').forEach(x => x.setAttribute('aria-pressed', x === b));
    renderMap();
  });
  document.getElementById('rad').onchange = e => {
    const n = Math.round(parseFloat(String(e.target.value).replace(',', '.').replace(/\s/g, '')));
    if(!Number.isFinite(n) || n < 10){ e.target.value = mapRadius; return; }
    mapRadius = Math.min(n, 200000);
    if(geoBox && geoBox.kind === 'circle') setBox({ kind: 'circle', lat: geoBox.lat, lon: geoBox.lon, r: mapRadius });
    else renderStatus();
  };
  document.getElementById('box-clear').onclick = () => { geoBox = null; applyFilters(); };
  // таблицы внутри перерисовываются — слушаем на обёртке
  document.getElementById('ot').onclick = e => {
    const th = e.target.closest('th[data-s]');
    if(!th) return omsuTableClick(e);
    const k = th.dataset.s;
    // первый щелчок по числу — по убыванию, по названию — по алфавиту; повторный — наоборот
    OT_SORT = OT_SORT.key === k ? { key: k, asc: !OT_SORT.asc } : { key: k, asc: k === 'omsu' };
    renderOmsuTable();
  };
  document.getElementById('ot-reset').onclick = () => msSet('omsu', null);
  document.getElementById('xls').onclick = exportXlsx;
}
const clearQuick = () => document.querySelectorAll('#quick .per-opt').forEach(b => b.classList.remove('on'));

async function start() {
  try {
    const response = await fetch('/mingkh/api/water-map', {cache:'no-store'});
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || data.detail || 'Не удалось загрузить карту.');
    if (!loadData(data)) throw new Error('В архиве нет данных карты.');
    document.querySelectorAll('[data-dataset-ui]').forEach(el => el.hidden = false);
    msBuild('omsu'); msBuild('fact'); msBuild('org'); buildGroups(); wire(); quick('30'); applyFilters();
    document.querySelectorAll('[data-kind]').forEach(el => el.onchange = () => { F.kinds[Number(el.dataset.kind)] = el.checked; applyFilters(); });
    document.getElementById('fit-map').onclick = () => renderMap(true);
    document.getElementById('kpis').addEventListener('keydown', e => {
      if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('[role=button]')) { e.preventDefault(); groupClick(e); }
    });
    await initMap();
  } catch (error) {
    document.getElementById('meta').textContent = 'Данные карты недоступны';
    const box = document.getElementById('load-error');
    box.hidden = false; box.textContent = error.message;
  }
}
const importForm = document.getElementById('import-form');
const archiveInput = document.getElementById('archive-file');
let archiveUploading = false, archiveRefreshing = false;
function syncArchiveControls() {
  if (!archiveInput) return;
  const busy = archiveUploading || archiveRefreshing;
  archiveInput.disabled = busy;
  document.getElementById('import-submit').disabled = busy;
  document.getElementById('import-submit').textContent = archiveUploading ? 'Загружаю…' : 'Загрузить архив';
  importForm.setAttribute('aria-busy', String(archiveUploading));
  const refresh = document.getElementById('live-refresh');
  if (refresh) refresh.disabled = busy;
}
if (archiveInput) archiveInput.addEventListener('change', () => {
  const file = archiveInput.files[0], filename = document.getElementById('archive-filename');
  filename.textContent = file ? file.name : 'Файл не выбран';
  filename.title = file ? file.name : '';
  filename.classList.toggle('has-file', !!file);
  document.getElementById('import-result').textContent = '';
});
if (importForm) importForm.onsubmit = async event => {
  event.preventDefault();
  const file = document.getElementById('archive-file').files[0];
  if (!file) return;
  const status = document.getElementById('import-result');
  if (file.size > 100 * 1024 * 1024) { status.textContent = 'Файл больше 100 МБ.'; return; }
  archiveUploading = true; syncArchiveControls(); status.textContent = 'Загружаю и проверяю архив…';
  try {
    const response = await fetch('/mingkh/api/water-map/import', {method:'POST',headers:{'Content-Type':'application/octet-stream'},body:file});
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || result.detail || 'Не удалось загрузить архив.');
    status.textContent = 'Импортировано ' + fmtN(result.count) + ' жалоб. Открываю карту…';
    window.location.reload();
  } catch (error) { status.textContent = error.message; archiveUploading = false; syncArchiveControls(); }
};
const refreshButton = document.getElementById('live-refresh');
let refreshTimer = null, followingRefresh = false;
async function refreshStatus() {
  if (!refreshButton) return;
  try {
    const response = await fetch('/mingkh/api/water-map/status', {cache:'no-store'});
    if (!response.ok) throw new Error('Не удалось проверить обновление. Перезагрузите страницу.');
    const state = await response.json();
    const output = document.getElementById('live-status');
    output.textContent = state.error || state.message || '';
    archiveRefreshing = state.busy; syncArchiveControls();
    if (state.busy) { followingRefresh = true; refreshTimer = setTimeout(refreshStatus, 3000); }
    else if (followingRefresh && state.completed_at && !state.error) window.location.reload();
  } catch (error) {
    document.getElementById('live-status').textContent = error.message;
    archiveRefreshing = false; syncArchiveControls();
  }
}
if (refreshButton) {
  refreshButton.onclick = async () => {
    clearTimeout(refreshTimer); archiveRefreshing = true; syncArchiveControls();
    document.getElementById('live-status').textContent = 'Начинаю обновление…';
    try {
      const response = await fetch('/mingkh/api/water-map/refresh', {method:'POST'});
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || result.detail || 'Не удалось запустить обновление.');
      followingRefresh = true; await refreshStatus();
    } catch (error) {
      document.getElementById('live-status').textContent = error.message;
      archiveRefreshing = false; syncArchiveControls();
    }
  };
  refreshStatus();
}
start();


})();
