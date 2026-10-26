(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  let saved, current = 4, period = null, sort = {key: 'place', direction: 1}, polling, loading = false;
  const normalize = value => String(value || '').toUpperCase().replace(/Ё/g, 'Е').trim();
  const number = value => typeof value === 'number' ? value.toLocaleString('ru-RU', {maximumFractionDigits: 2}) : '—';
  const date = value => value ? new Date(value).toLocaleString('ru-RU') : 'ещё не получены';
  function el(tag, text, cls) { const node = document.createElement(tag); if (text !== undefined) node.textContent = text; if (cls) node.className = cls; return node; }
  async function api(path, method = 'GET') {
    const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 20000);
    try {
      const response = await fetch('/omsu-rating/api/' + path, {method, credentials: 'same-origin', cache: 'no-store', redirect: 'error', signal: controller.signal, headers: {Accept: 'application/json'}});
      if (!response.ok) throw Error(response.status === 401 || response.status === 403 ? 'Сессия завершилась или доступ изменён. Войдите в Нейрону и проверьте права.' : 'Не удалось получить данные. Попробуйте ещё раз.');
      return await response.json();
    } finally { clearTimeout(timer); }
  }
  function alert(message) { $('or-error').textContent = message; $('or-error').hidden = !message; }
  function busy(value) { $('or-refresh').disabled = value; $('or-refresh').textContent = value ? 'Проверяем источники…' : 'Обновить данные'; }
  function schedule() { clearTimeout(polling); polling = setTimeout(check, 3000); }
  async function check() {
    if (document.hidden) { schedule(); return; }
    try {
      const state = await api('status');
      if (state.status === 'running') schedule(); else await load();
    } catch (e) { alert(e.message || 'Связь прервалась. Проверим состояние ещё раз.'); schedule(); }
  }
  async function load() {
    if (loading) return; loading = true;
    try {
      saved = await api('data'); alert(saved.refresh?.error || '');
      if (period === null || !saved.indicators[current].periods[period]?.length) period = saved.indicators[current].default_period;
      render(); busy(saved.refresh.status === 'running');
      if (saved.refresh.status === 'running') schedule();
    } catch (e) { alert(e.message || 'Не удалось загрузить сохранённые данные. Повторите попытку.'); busy(false); }
    finally { loading = false; }
  }
  function render() {
    const item = saved.indicators[current];
    $('or-title').textContent = item.name; $('or-methodology').textContent = item.note;
    $('or-source-error').textContent = item.error || ''; $('or-source-error').hidden = !item.error;
    $('or-updated').textContent = 'Показатель ' + current + ' · получено: ' + date(item.updated_at);
    $('or-history').textContent = saved.history.length ? 'Сохранённых срезов на сервере: ' + saved.history.length + '. Последняя проверка: ' + date(saved.refresh.finished_at) + '.' : 'После первого обновления данные и история срезов будут сохранены на сервере.';
    const tabs = $('or-tabs'); tabs.replaceChildren();
    Object.entries(saved.indicators).forEach(([id, cfg]) => {
      const b = el('button', id + ' · ' + cfg.short, 'or-button'); b.type = 'button'; b.setAttribute('aria-pressed', String(+id === current));
      b.addEventListener('click', () => { current = +id; period = cfg.default_period; sort = {key: 'place', direction: 1}; render(); }); tabs.append(b);
    });
    const periods = $('or-period'); periods.replaceChildren();
    saved.period_labels.forEach((label, p) => { const option = el('option', label + (item.periods[p].length ? '' : ' · нет данных')); option.value = p; option.disabled = !item.periods[p].length; option.selected = p === period; periods.append(option); });
    if (period === null) periods.selectedIndex = -1;
    periods.disabled = period === null;
    $('or-metric').textContent = item.metric + ' ↕';
    document.querySelectorAll('[data-sort]').forEach(button => button.parentElement.setAttribute('aria-sort', button.dataset.sort === sort.key ? sort.direction > 0 ? 'ascending' : 'descending' : 'none'));
    $('or-export').href = '/omsu-rating/export.xlsx?ind=' + current + (period === null ? '' : '&period=' + period);
    const rows = period === null ? [] : item.periods[period], prior = period > 0 ? item.periods[period - 1] : [];
    const prev = new Map(prior.map(r => [normalize(r.omsu), r.place]));
    let shown = rows.slice(); const terms = normalize($('or-search').value).split(/[,;]/).map(t => t.trim()).filter(Boolean);
    if ($('or-all').checked && rows.length) {
      const have = new Set(rows.map(r => normalize(r.omsu)));
      saved.municipalities.filter(o => !have.has(normalize(o))).forEach(o => shown.push({omsu: o, group: ['ВЛАСИХА','ВОСХОД','ЗВЕЗДНЫЙ ГОРОДОК','КРАСНОЗНАМЕНСК','МОЛОДЕЖНЫЙ'].includes(normalize(o)) ? 'zato' : 'main', na: true}));
    }
    shown = shown.filter(r => !terms.length || terms.some(t => normalize(r.omsu).includes(t)));
    shown.sort((a, b) => (a.group === 'zato') - (b.group === 'zato') || (Boolean(a.na) - Boolean(b.na)) || (sort.key === 'omsu' ? a.omsu.localeCompare(b.omsu, 'ru') : (a[sort.key] || 0) - (b[sort.key] || 0)) * sort.direction);
    const summary = $('or-summary'); summary.replaceChildren();
    [rows.length + ' ОМСУ участвует', 'Показано: ' + shown.length, current === 6 ? 'Места условные' : current === 7 ? 'Предварительная оценка' : 'ЗАТО — отдельная группа'].forEach(t => summary.append(el('span', t)));
    const tbody = $('or-rows'); tbody.replaceChildren(); let previousGroup;
    if (!shown.length) { const tr = el('tr'), td = el('td', rows.length ? 'Ничего не найдено. Измените фильтр.' : 'Пока нет данных. Нажмите «Обновить данные»; API-ключ добавляется в админ-панели.', 'or-empty'); td.colSpan = 5; tr.append(td); tbody.append(tr); }
    shown.forEach(row => {
      if (row.group !== previousGroup) { const tr = el('tr', undefined, 'or-group'), td = el('td', row.group === 'zato' ? 'ЗАТО' : 'Муниципалитеты'); td.colSpan = 5; tr.append(td); tbody.append(tr); previousGroup = row.group; }
      const tr = el('tr'); tr.append(el('td', row.na ? '—' : String(row.place), 'or-value'), el('td', row.omsu), el('td', number(row.value), 'or-value'));
      const previous = prev.get(normalize(row.omsu)), delta = previous === undefined || row.na ? null : previous - row.place;
      tr.append(el('td', delta === null ? 'Нет сравнения' : delta === 0 ? 'Без изменений' : (delta > 0 ? '↑ ' : '↓ ') + Math.abs(delta) + ' мест', 'or-delta ' + (delta > 0 ? 'up' : delta < 0 ? 'down' : '')));
      const detail = el('td');
      if (row.na) detail.textContent = 'Не участвует'; else { const b = el('button', 'Подробнее', 'or-button'); b.type = 'button'; b.setAttribute('aria-label', 'Расчёт: ' + row.omsu); b.addEventListener('click', () => showDetail(row, item)); detail.append(b); }
      tr.append(detail); tbody.append(tr);
    });
  }
  function showDetail(row, item) {
    $('or-detail-title').textContent = row.omsu + ' · показатель ' + current;
    const body = $('or-detail-body'); body.replaceChildren(el('p', item.note, 'or-note'));
    if (row.no_docs) body.append(el('p', 'Балл обнулён по всем задачам ОМСУ.', 'or-alert'));
    if (row.extra.total) body.append(el('p', 'Лицевых счетов без ИПУ: ' + number(row.extra.without) + ' из ' + number(row.extra.total)));
    row.details.forEach(d => {
      const card = el('article'); card.append(el('h3', d.name));
      if (current === 4) card.append(el('p', 'Начислено: ' + number(d.amount) + ' тыс. руб. · НВВ: ' + number(d.nvv) + ' тыс. руб.'), el('p', 'Факт: ' + number(d.fact) + '% · План: ' + number(d.plan) + '% · Выполнение: ' + number(d.value) + '%'));
      else card.append(el('p', 'Оценка: ' + number(d.value) + (current === 7 ? '% · Подтверждено мероприятий: ' + d.done + ' из ' + d.plan : '') + (d.zeroed ? ' · Обнулено в источнике' : '')));
      if (d.items) { const details = el('details'), list = el('ol'); details.append(el('summary', 'Как оценены мероприятия')); d.items.forEach(i => { const li = el('li', (i.ok ? 'Подтверждено: ' : 'Не подтверждено: ') + i.name); if (i.raw) li.append(el('p', 'В источнике: ' + i.raw)); if (i.reason) li.append(el('p', i.reason, 'or-note')); list.append(li); }); details.append(list); card.append(details); }
      body.append(card);
    });
    $('or-detail').showModal();
  }
  $('or-detail-close').addEventListener('click', () => $('or-detail').close());
  $('or-refresh').addEventListener('click', async () => { busy(true); alert(''); try { await api('refresh', 'POST'); schedule(); } catch (e) { busy(false); alert(e.message || 'Не удалось начать обновление.'); } });
  $('or-period').addEventListener('change', e => { period = Number(e.target.value); render(); });
  $('or-search').addEventListener('input', () => { if (saved) render(); }); $('or-all').addEventListener('change', () => { if (saved) render(); });
  document.querySelectorAll('[data-sort]').forEach(button => button.addEventListener('click', () => { const key = button.dataset.sort; sort = {key, direction: sort.key === key ? -sort.direction : key === 'value' && current !== 6 ? -1 : 1}; if (saved) render(); }));
  window.addEventListener('pagehide', () => clearTimeout(polling));
  window.addEventListener('pageshow', event => { if (event.persisted) load(); });
  document.addEventListener('visibilitychange', () => { if (!document.hidden && saved?.refresh.status === 'running') check(); });
  load();
})();
