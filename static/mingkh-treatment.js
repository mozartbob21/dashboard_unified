(() => {
  'use strict';
  const byId = id => document.getElementById(id);
  const form = byId('trFilters');
  const fields = form.elements;
  const workspace = byId('trWorkspace');
  const detail = byId('trDetail');
  const size = 40;
  const number = new Intl.NumberFormat('ru');
  let data, objects, selected, map, points, view = 'objects', page = 1;
  let detailGeneration = 0;
  let exportRows = [];

  const value = v => v === null || v === undefined || v === '' ? '—' : String(v);
  const formattedDate = v => {
    const day = String(v || '').slice(0, 10);
    return /^\d{4}-\d\d-\d\d$/.test(day) ? day.split('-').reverse().join('.') : value(v);
  };
  function element(tag, content) {
    const el = document.createElement(tag);
    if (content !== undefined) el.textContent = value(content);
    return el;
  }
  function button(label, action) {
    const b = element('button', label);
    b.type = 'button';
    b.addEventListener('click', action);
    return b;
  }
  function options(field, label, entries) {
    field.replaceChildren(new Option(label, ''));
    for (const [key, text] of entries) field.add(new Option(text, key));
  }
  function sortedNames(names) {
    return [...new Set(names.filter(Boolean))].sort((a,b) => a.localeCompare(b,'ru')).map(n => [n,n]);
  }
  const title = k => [k.name, k.place].filter(Boolean).join(' · ') + ' (№' + k.n + ')';
  function linked(c) {
    const radius = Math.max(data.radii[c.topic], c.mk ? 10 : 0);
    return c.distance !== null && c.distance <= radius ? objects.get(c.near) : null;
  }
  function matches(c) {
    const k = linked(c);
    if (fields.linked.checked && !k) return false;
    if ((fields.omsu.value && (k ? k.omsu : c.omsu) !== fields.omsu.value) ||
        (fields.obj.value && String(k?.n ?? '') !== fields.obj.value) ||
        (fields.cur.value && c.cur !== fields.cur.value) ||
        (fields.topic.value && c.topic !== fields.topic.value) ||
        (fields.sg.value && c.sg !== fields.sg.value)) return false;
    const day = String(fields.field.value === 'answered' ? (c.sg === 'closed' ? c.changed : '') : c.created).slice(0,10);
    if ((fields.start.value || fields.end.value) &&
        (!day || (fields.start.value && day < fields.start.value) || (fields.end.value && day > fields.end.value))) return false;
    const q = fields.q.value.trim().toLocaleLowerCase('ru');
    return !q || [c.id, c.fact, c.org, c.subcat, c.omsu, k?.name, k?.place].join(' ').toLocaleLowerCase('ru').includes(q);
  }
  function grouped(rows, keyOf) {
    const groups = new Map();
    for (const row of rows) {
      const key = keyOf(row);
      let stat = groups.get(key);
      if (!stat) {
        stat = {all:0, closed:0, work:0, izliv:0, smell:0, kos:0};
        groups.set(key, stat);
      }
      stat.all++;
      stat[row.topic]++;
      stat[row.sg]++;
    }
    return [...groups].sort((a,b) => b[1].all - a[1].all || String(a[0]).localeCompare(String(b[0]),'ru'));
  }
  function detailFields(items) {
    const dl = element('dl');
    for (const [name, val] of items) dl.append(element('dt',name), element('dd',val));
    byId('trDetailBody').replaceChildren(dl);
  }
  function openDetail(label, items) {
    byId('trDetailTitle').textContent = label;
    detailFields(items);
    if (!detail.open) detail.showModal();
  }
  async function fetchDetail(kind, id) {
    const generation = ++detailGeneration;
    const label = kind === 'events' ? 'Событие ЕДДС ' : 'Жалоба №';
    openDetail(label + id, [['Состояние','Загрузка…']]);
    try {
      const response = await fetch('/mingkh/treatment/api/' + kind + '/' + encodeURIComponent(id), {cache:'no-store'});
      if (!response.ok) throw Error('Не удалось получить запись. Проверьте доступ или обновите страницу.');
      const row = await response.json();
      if (generation !== detailGeneration || !detail.open) return;
      detailFields(kind === 'events' ? [
        ['ОМСУ',row.omsu],['РСО',row.rso],['Вид',row.typeName],['Начало',row.start],
        ['Завершение',row.close],['Описание',row.text],['Затронутые объекты',(row.objs || []).join(', ')],
      ] : [
        ['ОМСУ',row.omsu],['Адрес',row.address],['Тема',data.topics[row.topic]],['Основание',row.why],
        ['Получена',formattedDate(row.created)],['Статус',row.status],['Срок',formattedDate(row.deadline)],
        ['Исполнитель',row.org],['Объект',linked(row) ? title(linked(row)) : 'Не привязана'],
        ['Расстояние до ближайшего объекта',row.distance === null ? null : row.distance.toFixed(2) + ' км'],['Текст',row.text],
      ]);
    } catch (e) { if (generation === detailGeneration) detailFields([['Ошибка',e.message]]); }
  }

  function table(headers, rows) {
    const area = byId('trTable'), pages = byId('trPages');
    pages.replaceChildren();
    exportRows = [headers,...rows];
    byId('trExport').disabled = !rows.length;
    if (!rows.length) {
      area.replaceChildren(element('p','По выбранным условиям записей нет.'));
      return;
    }
    const count = Math.ceil(rows.length / size);
    page = Math.max(1, Math.min(page, count));
    const table = element('table'), head = table.createTHead().insertRow();
    for (const h of headers) {
      const th = element('th',h);
      th.scope = 'col';
      head.append(th);
    }
    const body = table.createTBody();
    for (const cells of rows.slice((page-1)*size, page*size)) {
      const tr = body.insertRow();
      for (const v of cells) {
        const td = tr.insertCell();
        // Delayed construction keeps even a full register at one page of buttons.
        if (typeof v === 'function') td.append(v()); else td.textContent = value(v);
      }
    }
    area.replaceChildren(table);
    if (count > 1) {
      const prev = button('← Назад', () => { page--; render(); });
      const next = button('Далее →', () => { page++; render(); });
      prev.disabled = page === 1; next.disabled = page === count;
      pages.append(prev, element('span','Страница ' + page + ' из ' + count), next);
    }
  }
  function showAppeals(key, selectedVal) {
    fields[key].value = selectedVal;
    view = 'appeals'; page = 1; render();
  }
  function objectCard(k) {
    ++detailGeneration;
    openDetail(title(k), [
      ['ОМСУ', k.omsu],['Состояние', k.state],['Эксплуатант',k.operator],['Собственник',k.owner],
      ['Мощность',k.power],['Координаты',k.lat === null || k.lon === null ? null : k.lat + ', ' + k.lon],
      ['Проверка координат',k.check],
    ]);
  }
  function drawMap() {
    const host = byId('trMap');
    host.hidden = false;
    if (!window.L || !window.NeuronaBasemap) {
      byId('trMapStatus').textContent = 'Карта недоступна. Данные доступны в таблицах.';
      return;
    }
    if (!map) {
      map = L.map(byId('trCanvas'), {preferCanvas:true}).setView([55.67,37.62],8);
      window.NeuronaBasemap.create(map, {onStatus: msg => { byId('trMapStatus').textContent = msg; }});
      points = L.layerGroup().addTo(map);
    }
    map.invalidateSize();
    points.clearLayers();
    const kos = new Set();
    function marker(lat, lon, color, label, open) {
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
      const m = L.circleMarker([lat,lon], {radius:5, color, fillOpacity:.8, weight:1});
      m.bindTooltip(element('span', label)).on('click', open).addTo(points);
    }
    for (const c of selected) {
      const k = linked(c);
      if (k) kos.add(k);
      marker(c.lat, c.lon, '#b54624', 'Жалоба №' + c.id, () => fetchDetail('complaints', c.id));
    }
    for (const k of kos) marker(k.lat, k.lon, '#3e50ba', title(k), () => objectCard(k));
    const bounds = L.latLngBounds(points.getLayers().map(pt => pt.getLatLng()));
    if (bounds.isValid()) map.fitBounds(bounds.pad(.07), {maxZoom:13});
  }

  function render() {
    if (!data) return;
    selected = data.complaints.filter(matches);
    const metrics = byId('trMetrics'); metrics.replaceChildren();
    const counts = grouped(selected, () => 'all')[0]?.[1] || {};
    for (const [label, n] of [['Жалоб',selected.length], ['В работе',counts.work || 0], ['Закрыто',counts.closed || 0],
                              ['Объектов с жалобами',new Set(selected.map(linked).filter(Boolean)).size]]) {
      const tile = element('div'); tile.append(element('strong',number.format(n)), element('span',label)); metrics.append(tile);
    }
    const active = byId('trTabs').querySelector('[data-view="' + view + '"]');
    for (const b of byId('trTabs').querySelectorAll('button')) b.removeAttribute('aria-current');
    active.setAttribute('aria-current','page');
    byId('trTitle').textContent = active.textContent;
    byId('trNote').textContent = '';
    byId('trMap').hidden = true;
    if (fields.start.value && fields.end.value && fields.start.value > fields.end.value) {
      byId('trNote').textContent = 'Начало периода должно быть не позже его окончания.';
      table([],[]);return;
    }
    if (view === 'map') {
      byId('trExport').disabled = true;
      byId('trTable').replaceChildren(); byId('trPages').replaceChildren();
      byId('trNote').textContent = 'Синие точки — объекты КОС, оранжевые — жалобы. Нажмите точку, чтобы открыть сведения.';
      drawMap(); return;
    }
    if (view === 'appeals') {
      table(['Жалоба','Дата','ОМСУ','Тематика','Статус','Объект КОС'], [...selected]
        .sort((a,b) => b.created.localeCompare(a.created))
        .map(c => [() => button('№' + c.id, () => fetchDetail('complaints',c.id)), formattedDate(c.created),
          linked(c)?.omsu || c.omsu,data.topics[c.topic],c.status,linked(c) ? title(linked(c)) : 'Без привязки']));
    } else if (view === 'objects') {
      const rows = grouped(selected, c => linked(c)?.n ?? '');
      table(['Объект / карточка','ОМСУ','Жалобы','Изливы','Запах','Работа КОС','В работе'],rows.map(([key, st]) => {
        const k = objects.get(key);
        return [k ? () => button(title(k),() => objectCard(k)) : 'Без привязки', k?.omsu,
          k ? () => button(st.all,() => showAppeals('obj',String(key))) : st.all, st.izliv, st.smell, st.kos, st.work];
      }));
    } else if (view === 'rating') {
      table(['ОМСУ','Жалобы','Изливы','Запах','Работа КОС','Закрыто','В работе'],
        grouped(selected, c => linked(c)?.omsu || c.omsu).map(([key,st]) => [
          () => button(key,() => showAppeals('omsu',key)),st.all,st.izliv,st.smell,st.kos,st.closed,st.work]));
    } else if (view === 'stats') {
      const series = grouped(selected, c => (fields.field.value === 'answered' ? c.changed : c.created).slice(0,7))
        .sort((a,b) => b[0].localeCompare(a[0]));
      table(['Месяц','Всего','Изливы','Запах','Работа КОС','В работе','Закрыто'], series.map(([key,st]) =>
        [key,st.all,st.izliv,st.smell,st.kos,st.work,st.closed]));
    } else if (view === 'operators') {
      const norm = name => String(name || '').toLocaleLowerCase('ru').replace(/ё/g,'е')
        .replace(/["«»“”„']/g,'').replace(/\(арендатор\)/,'').replace(/\s+/g,' ').trim();
      const operatorKey = k => {
        const name = norm(k.operator);
        return /^(муп|зао|ао|ооо|гуп мо|мкп|мп) водоканал$/.test(name) ? name + ' | ' + k.okey : name;
      };
      const operators = new Map();
      for (const k of data.kos) {
        if ((fields.omsu.value && k.omsu !== fields.omsu.value) || (fields.obj.value && String(k.n) !== fields.obj.value)) continue;
        const key = operatorKey(k);
        if (!operators.has(key)) operators.set(key,{key,name:k.operator,places:new Set(),facilities:0,power:0,appeals:0,open:0});
        const op = operators.get(key);
        op.places.add(k.omsu); op.facilities++; op.power += k.power || 0;
      }
      for (const c of selected) {
        const k = linked(c);
        const op = k && operators.get(operatorKey(k));
        if (op) { op.appeals++; if (c.sg === 'work') op.open++; }
      }
      const payments = new Map(data.payments.map(p => [p.key,p]));
      const rub = n => Number.isFinite(n) ? number.format(Math.round(n)) : null;
      byId('trNote').textContent = 'Плата за негативное воздействие на ' + formattedDate(data.pay_as_of) + '. Тире означает отсутствие данных.';
      const needle = fields.q.value.trim().toLocaleLowerCase('ru');
      const groups = [...operators.values()].filter(g => !needle ||
        [g.name,...g.places].join(' ').toLocaleLowerCase('ru').includes(needle)).sort((a,b) => b.facilities-a.facilities);
      table(['Эксплуатант','КОС','ОМСУ','Мощность, м³/сут','Жалобы','В работе','Получено платы, ₽','Начислено, ₽'],
        groups.map(g => {
          const pay = payments.get(g.key) || payments.get(norm(g.name));
          return [g.name,g.facilities,[...g.places].join(', '),number.format(g.power),g.appeals,g.open,rub(pay?.recv),rub(pay?.accr)];
        }));
    } else if (view === 'coordinates') {
      const needle = fields.q.value.trim().toLocaleLowerCase('ru');
      const kos = data.kos.filter(k => (!fields.omsu.value || k.omsu === fields.omsu.value) &&
        (!fields.obj.value || String(k.n) === fields.obj.value) &&
        (!needle || [k.n,k.omsu,k.name,k.place].join(' ').toLocaleLowerCase('ru').includes(needle)));
      byId('trNote').textContent = 'Координаты объектов в сохранённом снимке. Применяются фильтры ОМСУ, объекта и поиска.';
      table(['Объект','ОМСУ','Состояние','Широта','Долгота','Проверка координат'], kos.map(k =>
        [() => button(title(k),() => objectCard(k)),k.omsu,k.state,k.lat,k.lon,k.check]));
    } else if (view === 'records') {
      byId('trNote').textContent = 'Справочные записи снимка. Применяются фильтры ОМСУ и объекта.';
      const recs = data.records.filter(row => (!fields.obj.value || String(row.n) === fields.obj.value) &&
        (!fields.omsu.value || (objects.get(row.n)?.omsu || row.omsu) === fields.omsu.value));
      table(['Вид записи','Объект / мероприятие','ОМСУ','Период','Статус / охват'],
        recs.map(r => [r.t === 'gp' ? 'Госпрограмма' : r.t === 'ox' ? 'Охват' : r.t,
          r.event || r.obj || (objects.has(r.n) ? title(objects.get(r.n)) : null),
          r.omsu || objects.get(r.n)?.omsu, [r.start,r.end].filter(Boolean).join(' — '),
          r.status || (r.people !== null ? number.format(r.people) + ' чел.' : null)]));
    } else if (view === 'edds') {
      byId('trNote').textContent = 'События ЕДДС по ОМСУ, объекту и дате начала. Данные источника на ' + value(data.edds_built);
      const events = data.edds.filter(r => (!fields.omsu.value || r.omsu === fields.omsu.value) &&
        (!fields.obj.value || String(r.n) === fields.obj.value) &&
        (!fields.start.value || (r.start || '').slice(0,10) >= fields.start.value) &&
        (!fields.end.value || (r.start || '').slice(0,10) <= fields.end.value));
      table(['Событие','ОМСУ','Начало','Вид','РСО'],events.map(r => [
        () => button(r.id, () => fetchDetail('events',r.id)),r.omsu,formattedDate(r.start),r.typeName,r.rso]));
    }
  }
  async function load() {
    const res = await fetch('/mingkh/treatment/api/data',{cache:'no-store'});
    if (res.status === 404) {
      workspace.hidden = true; byId('trSource').textContent = 'Снимок ещё не загружен.'; return;
    }
    if (!res.ok) throw Error('Не удалось загрузить снимок. Проверьте доступ и обновите страницу.');
    data = await res.json(); objects = new Map(data.kos.map(k => [k.n,k]));
    byId('trSource').textContent = 'Снимок на ' + data.built + ' · Жалобы: ' + number.format(data.complaints.length) +
      ' · Объекты КОС: ' + number.format(data.kos.length);
    options(fields.omsu,'Все ОМСУ',sortedNames([...data.kos,...data.complaints].map(r => r.omsu)));
    options(fields.obj,'Все объекты',data.kos.map(k => [k.n,title(k)]));
    options(fields.cur,'Все кураторы',sortedNames(data.complaints.map(c => c.cur)));
    options(fields.topic,'Все темы',Object.entries(data.topics));
    fields.start.defaultValue = data.from; fields.end.defaultValue = data.as_of;
    form.reset(); workspace.hidden = false; page = 1; render();
  }
  byId('trExport').addEventListener('click',() => {
    if (!window.XLSX || exportRows.length < 2) return;
    const values = exportRows.map(row => row.map(v => typeof v === 'function' ? v().textContent : v ?? ''));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet(values),'Данные');
    XLSX.writeFile(workbook,'Жалобы по очистным — ' + byId('trTitle').textContent + '.xlsx');
  });
  form.addEventListener('submit', e => e.preventDefault());
  form.addEventListener('input', () => {page=1;render();});
  form.addEventListener('reset', () => setTimeout(() => {page=1;render();},0));
  byId('trTabs').addEventListener('click', e => {
    const tab = e.target.closest('[data-view]');
    if (tab) { view=tab.dataset.view;page=1;render(); }
  });
  byId('trImport')?.addEventListener('submit', async e => {
    e.preventDefault();
    const file = byId('trZip').files[0], submit = e.target.querySelector('button');
    if (!file) return;
    submit.disabled = true; byId('trStatus').textContent = 'Проверяем и сохраняем снимок…';
    try {
      const res = await fetch('/mingkh/treatment/api/import',{method:'POST',body:file,
        headers:{'Content-Type':'application/zip'}});
      const info = await res.json();
      if (!res.ok) throw Error(info.detail || 'Не удалось загрузить снимок.');
      await load(); byId('trStatus').textContent = 'Снимок сохранён.'; byId('trZip').value = '';
    } catch (err) { byId('trStatus').textContent = err.message; }
    finally { submit.disabled = false; }
  });
  load().catch(err => {byId('trStatus').textContent = err.message;});
})();
