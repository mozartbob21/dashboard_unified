(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const preferenceKey = 'neurona:omsu-rating:v1';
  const snapshotNode = $('or-snapshot-data');
  const offline = Boolean(snapshotNode);
  let saved, current = 5, period = null, sort = {key: 'place', direction: 1};
  let polling, loading = false, exportUrl;
  const normalize = value => String(value ?? '').toUpperCase().replace(/Ё/g, 'Е').trim();
  const number = (value, digits = 2) => typeof value === 'number' && Number.isFinite(value)
    ? value.toLocaleString('ru-RU', {minimumFractionDigits: digits, maximumFractionDigits: digits}) : '—';
  const date = value => value ? new Date(value).toLocaleString('ru-RU') : 'ещё не получены';
  const periodsTo = ['1 кв.', '1 пол. (2 кв.)', '9 мес. (3 кв.)', '11 мес. (4 кв.)'];
  const tagLabels = {achieved: 'План выполнен', not_achieved: 'План не выполнен', zeroed: 'Балл обнулён'};
  const zoneLabels = {z1: 'Верхняя треть рейтинга', z2: 'Средняя треть рейтинга', z3: 'Нижняя треть рейтинга'};
  const descriptions = {
    4: 'Критерий отражает эффективность работы РСО по взиманию платы за негативное воздействие на централизованные системы водоотведения. Факт — доля сбора платы в НВВ РСО на 2026 год, утверждённой Комитетом по ценам и тарифам Московской области. План зависит от категории РСО по НВВ. Несколько РСО — среднее выполнения их планов.',
    5: 'Критерий отражает работу РСО по снижению сверхнормативных расходов воды: потерь при транспортировке и бездоговорного водопотребления. Балл от 0 до 10; для нескольких РСО или задач берётся средний балл. При обнулении всех задач ОМСУ получает последнее место: 51, для ЗАТО — 5.',
    6: 'Цель — стимулировать полный учёт ресурсов, установку ИПУ холодной и горячей воды либо применение повышающего коэффициента 3 при расчётах по нормативу. ЛС факт для расчёта по методике в источнике отсутствует. Поэтому места условные: у всех ОМСУ первое место, балл приведён для справки.',
    7: 'Снижение затрат на оплату электроэнергии оценивается по мероприятиям дорожной карты перехода точек учёта на оптимальную ценовую категорию. В таблице только текущее состояние: учитываются 17 мероприятий за 9 месяцев. Оценка предварительная. Выполнено, «не требуется», «нецелесообразно» засчитываются; частичное выполнение, «в работе» и перенос — нет. Отсутствующая РСО получает 0. При нескольких строках РСО выбирается самая полная.',
  };
  function el(tag, text, cls) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (cls) node.className = cls;
    return node;
  }
  function remember() {
    try { localStorage.setItem(preferenceKey, JSON.stringify({indicator: current, period, showAll: $('or-all').checked})); } catch (_) { /* Storage can be disabled by browser policy. */ }
  }
  try {
    const preference = JSON.parse(localStorage.getItem(preferenceKey) || 'null');
    if (preference && [4, 5, 6, 7].includes(preference.indicator)) {
      current = preference.indicator;
      period = Number.isInteger(preference.period) && preference.period >= 0 && preference.period <= 3 ? preference.period : null;
      $('or-all').checked = preference.showAll === true;
    }
  } catch (_) { /* Invalid preferences never prevent opening the rating. */ }
  async function api(path, method = 'GET') {
    const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 20000);
    try {
      const response = await fetch('/omsu-rating/api/' + path, {
        method, credentials: 'same-origin', cache: 'no-store', redirect: 'error', signal: controller.signal,
        headers: {Accept: 'application/json'},
      });
      if (!response.ok) throw Error(response.status === 401 || response.status === 403
        ? 'Сессия завершилась или доступ изменён. Войдите в Нейрону и проверьте права.'
        : 'Не удалось получить данные. Попробуйте ещё раз.');
      return await response.json();
    } finally { clearTimeout(timer); }
  }
  function alert(message) { $('or-error').textContent = message; $('or-error').hidden = !message; }
  function busy(value) {
    $('or-refresh').disabled = value;
    $('or-refresh').textContent = value ? 'Проверяем источники…' : 'Обновить данные';
    $('or-refresh').setAttribute('aria-busy', String(value));
  }
  function schedule() { clearTimeout(polling); if (!offline) polling = setTimeout(check, 3000); }
  async function check() {
    if (document.hidden) { schedule(); return; }
    try {
      const state = await api('status');
      if (state.status === 'running') schedule(); else await load();
    } catch (e) { alert(e.message || 'Связь прервалась. Проверим состояние ещё раз.'); schedule(); }
  }
  async function load() {
    if (loading || offline) return;
    loading = true;
    try {
      saved = await api('data');
      alert(saved.refresh?.error || '');
      render();
      busy(saved.refresh?.status === 'running');
      if (saved.refresh?.status === 'running') schedule();
    } catch (e) {
      alert(e.name === 'AbortError' ? 'Сервер отвечает дольше обычного. Сохранённые данные остаются на экране; повторите обновление.' : e.message || 'Не удалось загрузить сохранённые данные. Повторите попытку.');
      busy(false);
    } finally { loading = false; }
  }
  function rowValue(row, key) {
    if (key === 'rso') return (row.rso_names || []).join(', ');
    if (key === 'extra') return current === 7 ? row.extra.done / row.extra.plan : row.extra?.fact ?? null;
    if (key === 'plan') return row.extra?.plan?.[0] ?? null;
    if (key === 'lsno') return row.extra?.without ?? null;
    if (key === 'lstot') return row.extra?.total ?? null;
    return row[key] ?? null;
  }
  function sortedRows(rows) {
    return rows.slice().sort((a, b) => {
      const group = (a.group === 'zato') - (b.group === 'zato');
      if (group) return group;
      const x = rowValue(a, sort.key), y = rowValue(b, sort.key);
      if (x === null && y === null) return a.place - b.place;
      if (x === null) return 1;
      if (y === null) return -1;
      const comparison = typeof x === 'string' ? x.localeCompare(y, 'ru') : x - y;
      return comparison * sort.direction || a.place - b.place;
    });
  }
  function columns(item, hasPrevious) {
    const cols = [{key: 'place', label: 'Место', tip: 'Место по методике. Одинаковые места показаны диапазоном; обнулённые баллы — последним местом. ЗАТО ранжируются отдельно.'}];
    if (hasPrevious) cols.push({key: 'delta', label: 'Динамика к ' + periodsTo[period - 1], tip: 'Место в предыдущем периоде минус место в текущем. Положительное значение — рост в рейтинге.'});
    cols.push({key: 'omsu', label: 'ОМСУ', tip: 'Цвет означает верхнюю, среднюю или нижнюю треть рейтинга по месту в своей полной группе.'},
      {key: 'value', label: item.metric, tip: descriptions[current]});
    if (current === 4) cols.push({key: 'extra', label: 'Факт, % от НВВ', tip: 'Сбор платы нарастающим итогом / НВВ × 100. Для нескольких РСО — отношение сумм.'},
      {key: 'plan', label: 'План, %', tip: 'План по категории НВВ каждой РСО: менее 300 млн, от 300 до 600 млн, 600 млн руб. и более.'});
    if (current === 6) cols.push({key: 'lsno', label: 'ЛС без ИПУ', tip: 'Сумма лицевых счетов категорий ЛС1, ЛС2, ЛС3 и ЛС4.'}, {key: 'lstot', label: 'Всего ЛС', tip: 'Общее количество лицевых счетов из источника.'});
    if (current === 7) cols.push({key: 'extra', label: 'Мероприятий выполнено', tip: 'Выполнено / запланировано, сумма по всем РСО ОМСУ. Сортировка — по доле выполнения.'});
    if (current !== 6) cols.push({key: 'rso', label: 'РСО', tip: 'Организации и их баллы. Подробный расчёт доступен по кнопке в строке.'});
    return cols;
  }
  function renderHeader(cols, item) {
    const supported = ['place', 'delta', 'omsu', 'value', 'extra', 'plan', 'lsno', 'lstot', 'rso'];
    if (!supported.includes(sort.key) || (sort.key === 'delta' && !cols.some(col => col.key === 'delta')) || (sort.key === 'plan' && current !== 4)) sort = {key: 'place', direction: 1};
    const tr = el('tr');
    cols.forEach(col => {
      const th = el('th'); th.scope = 'col';
      th.setAttribute('aria-sort', col.key === sort.key ? sort.direction > 0 ? 'ascending' : 'descending' : 'none');
      const button = el('button', col.label + (sort.key === col.key ? sort.direction > 0 ? ' ↑' : ' ↓' : ' ↕'), 'or-sort');
      button.type = 'button'; button.dataset.sort = col.key; button.title = col.tip;
      button.setAttribute('aria-label', 'Сортировать: ' + col.label);
      button.addEventListener('click', () => {
        const descending = (['value', 'extra'].includes(col.key) && item.better !== 'asc') || col.key === 'delta';
        sort = {key: col.key, direction: sort.key === col.key ? -sort.direction : descending ? -1 : 1};
        render();
        document.querySelector('[data-sort="' + col.key + '"]')?.focus();
      });
      th.append(button); tr.append(th);
    });
    const detail = el('th', 'Расчёт'); detail.scope = 'col'; tr.append(detail);
    $('or-head').replaceChildren(tr);
  }
  function renderSummary(stats, rows, matches) {
    const summary = $('or-summary'); summary.replaceChildren();
    const entries = [
      {label: 'ОМСУ в рейтинге', value: String(stats.count ?? rows.length), sub: stats.zato_count ? 'из них ЗАТО: ' + stats.zato_count : 'за выбранный период'},
      {label: 'Среднее по всем ОМСУ', value: number(stats.average), sub: 'фильтр не меняет расчёт'},
    ];
    if (current === 4) entries.push({label: 'Выполнили план', value: String(stats.achieved_count ?? 0), sub: 'из ' + rows.length + ' · выполнение ≥ 100 %'});
    if (normalize($('or-search').value)) entries.push({label: 'По фильтру', value: String(matches), sub: 'участвующих ОМСУ из ' + rows.length});
    entries.forEach(entry => {
      const card = el('div', undefined, 'or-stat');
      card.append(el('span', entry.label), el('strong', entry.value), el('small', entry.sub));
      summary.append(card);
    });
  }
  function cell(row, key, item) {
    const td = el('td'); td.dataset.column = key;
    if (key === 'omsu') {
      const chip = el('span', row.omsu, 'or-chip ' + (row.na ? 'or-na' : row.zone));
      if (!row.na) { chip.title = zoneLabels[row.zone] || ''; chip.setAttribute('aria-label', row.omsu + '. ' + (zoneLabels[row.zone] || '')); }
      td.append(chip);
      if (row.na) td.append(el('span', 'Не участвует', 'or-tag neutral'));
      else if (tagLabels[row.tag]) td.append(el('span', tagLabels[row.tag], 'or-tag ' + (row.tag === 'achieved' ? 'good' : 'bad')));
      return td;
    }
    if (row.na) { td.textContent = key === 'place' ? '—' : ''; return td; }
    if (key === 'place') { td.textContent = row.place_label || String(row.place); td.className = 'or-place'; td.title = row.place_tooltip || ''; }
    else if (key === 'delta') {
      td.className = 'or-delta ' + (row.delta > 0 ? 'up' : row.delta < 0 ? 'down' : 'same');
      td.textContent = row.delta == null ? 'Новый' : row.delta === 0 ? '—' : (row.delta > 0 ? '▲ ' : '▼ ') + Math.abs(row.delta);
      if (row.delta != null) td.title = saved.period_labels[period - 1] + ': ' + row.previous_place + ' место → ' + saved.period_labels[period] + ': ' + row.place + ' место';
      td.setAttribute('aria-label', row.delta == null ? 'Нет данных в предыдущем периоде' : row.delta === 0 ? 'Без изменений' : (row.delta > 0 ? 'Поднялся на ' : 'Опустился на ') + Math.abs(row.delta));
    } else if (key === 'value') {
      td.className = 'or-value'; td.title = row.value_tooltip || '';
      td.append(el('span', number(row.value)));
      const bar = el('span', undefined, 'or-bar'), fill = el('i');
      bar.setAttribute('aria-hidden', 'true'); fill.style.width = Math.max(2, Math.min(100, row.bar_pct ?? 2)) + '%';
      bar.append(fill); td.append(bar);
    } else if (key === 'extra') {
      td.className = 'or-value'; td.textContent = current === 7 ? row.extra.done + ' / ' + row.extra.plan : number(row.extra?.fact);
      td.title = row.extra?.fact_tooltip || 'Выполнено / запланировано мероприятий по всем РСО';
    } else if (key === 'plan') {
      td.className = 'or-value'; td.textContent = (row.extra?.plan || []).map(n => number(n, 0)).join(' / '); td.title = row.extra?.plan_tooltip || '';
    } else if (key === 'lsno' || key === 'lstot') {
      td.className = 'or-value'; td.textContent = number(row.extra?.[key === 'lsno' ? 'without' : 'total'], 0);
      td.title = key === 'lsno' ? row.extra?.without_tooltip || '' : 'Всего лицевых счетов из источника';
    } else if (key === 'rso') {
      td.className = 'or-rso';
      const details = row.details || [];
      if (current === 7 || details.length > 1) details.forEach(detail => {
        const line = el('span', detail.name + ': ', 'or-rso-line');
        line.append(el('strong', number(detail.value) + (current === 4 ? ' %' : '')));
        if (current === 7) line.title = k7Tip(detail);
        else if (current === 4) line.title = 'Факт ' + number(detail.fact) + ' % / план ' + number(detail.plan, 0) + ' % × 100';
        td.append(line);
      });
      else td.textContent = (row.rso_names || details.map(d => d.name)).join(', ') || '—';
    }
    return td;
  }
  function k7Tip(detail) {
    if (!detail.found) return detail.name + ': РСО нет в таблице — 0 из ' + detail.plan + ' мероприятий.';
    return detail.name + '\nВыполнено ' + detail.done + ' из ' + detail.plan + '\n' + detail.items.filter(i => !i.ok).map(i => i.n + '. ' + i.name + ' — ' + (i.reason || 'не подтверждено') + (i.raw ? ': «' + i.raw + '»' : '')).join('\n');
  }
  function updateSourceLink(item) {
    const link = $('or-source-link');
    link.hidden = true;
    if (!item.source_url) return;
    try {
      const url = new URL(item.source_url);
      const allowed = url.protocol === 'https:' && !url.username && !url.password && !url.port &&
        ((url.hostname === 'mgkh.rm.mosreg.ru' && url.pathname === '/issues') ||
         (url.hostname === 'docs.google.com' && url.pathname.startsWith('/spreadsheets/d/')));
      if (!allowed) return;
      link.href = url.href;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      link.textContent = current === 7 ? 'Открыть таблицу дорожных карт ↗' : 'Открыть задачи Redmine ↗';
      link.title = 'Источник откроется в новой вкладке. Может потребоваться отдельный вход.';
      link.hidden = false;
    } catch (_) { /* Only the fixed, valid source addresses may become links. */ }
  }
  function updateExport() {
    if (!offline) {
      $('or-export').href = '/omsu-rating/export.xlsx?ind=' + current + (period === null ? '' : '&period=' + period);
      return;
    }
    if (exportUrl) URL.revokeObjectURL(exportUrl);
    const encoded = saved.offline_exports?.[period === null ? 'default' : current + ':' + period];
    $('or-export').hidden = !encoded;
    if (encoded) {
      const binary = atob(encoded), bytes = Uint8Array.from(binary, character => character.charCodeAt(0));
      exportUrl = URL.createObjectURL(new Blob([bytes], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      $('or-export').href = exportUrl;
      $('or-export').download = 'Рейтинг ОМСУ.xlsx';
    }
  }
  function render() {
    const item = saved.indicators[current];
    if (period === null || !item.periods[period]?.length) period = item.default_period;
    remember();
    const rows = period === null ? [] : item.periods[period];
    const stats = period === null ? {} : item.period_stats?.[period] || {};
    const hasPrevious = Boolean(stats.has_previous);
    const cols = columns(item, hasPrevious);
    $('or-title').textContent = (item.roman ? item.roman + '. ' : '') + item.name;
    $('or-methodology').textContent = item.note;
    $('or-methodology-detail').textContent = item.about ? item.about + '\n\n' + (item.rule || '') + (current === 7 ? '\n\n' + descriptions[7] : '') : descriptions[current];
    $('or-source-error').textContent = item.error || ''; $('or-source-error').hidden = !item.error;
    $('or-updated').textContent = (offline ? 'Снимок · ' : '') + 'показатель ' + current + ' · получено: ' + date(item.updated_at);
    $('or-history').textContent = offline
      ? 'Автономный снимок сохранённых данных. Фильтры, сортировка и детализация работают без подключения. Для получения новых данных откройте рейтинг в Нейроне.'
      : saved.history?.length ? 'Сохранённых срезов на сервере: ' + saved.history.length + '. Последняя проверка: ' + date(saved.refresh?.finished_at) + '.'
        : 'После первого обновления данные и история срезов будут сохранены на сервере.';
    const tabs = $('or-tabs'); tabs.replaceChildren();
    Object.entries(saved.indicators).forEach(([id, cfg]) => {
      const button = el('button', (cfg.roman || id) + '. ' + cfg.short, 'or-button');
      button.type = 'button'; button.title = cfg.name; button.dataset.indicator = id;
      button.setAttribute('aria-pressed', String(+id === current));
      button.addEventListener('click', () => {
        current = +id; period = null; render(); document.querySelector('[data-indicator="' + id + '"]')?.focus();
      }); tabs.append(button);
    });
    const periods = $('or-period'); periods.replaceChildren();
    saved.period_labels.forEach((label, p) => {
      const option = el('option', label + (item.periods[p]?.length ? '' : ' · нет данных'));
      option.value = p; option.disabled = !item.periods[p]?.length; option.selected = p === period; periods.append(option);
    });
    if (period === null) periods.selectedIndex = -1;
    periods.disabled = period === null;
    const suggestions = $('or-omsu-list'); suggestions.replaceChildren();
    [...saved.municipalities].sort((a, b) => a.localeCompare(b, 'ru')).forEach(name => {
      const option = el('option'); option.value = name; suggestions.append(option);
    });
    updateExport();
    updateSourceLink(item);
    const terms = normalize($('or-search').value).split(/[,;]/).map(t => t.trim()).filter(Boolean);
    const matches = row => !terms.length || terms.some(t => normalize(row.omsu).includes(t));
    renderHeader(cols, item);
    const ordered = sortedRows(rows);
    const nonparticipants = $('or-all').checked && rows.length ? stats.nonparticipants || [] : [];
    const shown = ['main', 'zato'].flatMap(group => [...ordered.filter(r => r.group === group), ...nonparticipants.filter(r => r.group === group)]).filter(matches);
    renderSummary(stats, rows, rows.filter(matches).length);
    $('or-legend').hidden = !rows.length;
    const tbody = $('or-rows'); tbody.replaceChildren();
    let previousGroup;
    if (!shown.length) {
      const tr = el('tr'), td = el('td', rows.length ? 'Ничего не найдено. Измените фильтр.' : offline
        ? 'В снимке нет данных по этому показателю.' : 'Пока нет данных. Нажмите «Обновить данные»; API-ключ добавляется в админ-панели.', 'or-empty');
      td.colSpan = cols.length + 1; tr.append(td); tbody.append(tr);
    }
    shown.forEach(row => {
      if (row.group !== previousGroup) {
        const tr = el('tr', undefined, 'or-group'), td = el('td', row.group === 'zato' ? 'ЗАТО — отдельный рейтинг, места с 1 по 5' : 'Муниципалитеты');
        td.colSpan = cols.length + 1; tr.append(td); tbody.append(tr); previousGroup = row.group;
      }
      const tr = el('tr', undefined, row.na ? 'or-nonparticipant' : 'or-rating-row');
      tr.dataset.omsu = row.omsu;
      cols.forEach(col => tr.append(cell(row, col.key, item)));
      const detail = el('td');
      if (!row.na) {
        const button = el('button', 'Подробнее', 'or-button or-detail-button'); button.type = 'button';
        button.setAttribute('aria-label', 'Расчёт: ' + row.omsu);
        button.addEventListener('click', () => showDetail(row, item)); detail.append(button);
      }
      tr.append(detail); tbody.append(tr);
    });
  }
  function showDetail(row, item) {
    $('or-detail-title').textContent = row.omsu + ' · показатель ' + current;
    const body = $('or-detail-body'); body.replaceChildren(el('p', item.note, 'or-note'));
    body.append(el('p', row.place_tooltip, 'or-calculation'), el('p', row.value_tooltip, 'or-calculation'));
    if (row.no_docs) body.append(el('p', 'Балл обнулён по всем задачам ОМСУ.', 'or-alert'));
    if (current === 4) body.append(el('p', row.extra.fact_tooltip, 'or-calculation'), el('p', row.extra.plan_tooltip, 'or-calculation'));
    if (current === 6 && row.extra.total) body.append(el('p', row.extra.without_tooltip, 'or-calculation'), el('p', 'Всего лицевых счетов: ' + number(row.extra.total, 0)));
    (row.details || []).forEach(detail => {
      const card = el('article'); card.append(el('h3', detail.name));
      if (current === 4) card.append(
        el('p', 'Сбор платы нарастающим итогом: ' + number(detail.amount) + ' тыс. руб. · НВВ: ' + number(detail.nvv) + ' тыс. руб.'),
        el('p', 'Категория ' + detail.category + ': ' + detail.category_label),
        el('p', 'Факт: ' + number(detail.fact) + ' % · План: ' + number(detail.plan, 0) + ' % · Выполнение: ' + number(detail.value) + ' %'));
      else card.append(el('p', 'Оценка: ' + number(detail.value) + (current === 7 ? ' · Выполнено мероприятий: ' + detail.done + ' из ' + detail.plan : '') + (detail.zeroed ? ' · Обнулено в источнике' : '')));
      if (current === 7) {
        if (!detail.found) card.append(el('p', 'РСО нет в таблице. Все мероприятия засчитаны как невыполненные.', 'or-alert'));
        if (detail.duplicates > 1) card.append(el('p', 'Строк РСО в источнике: ' + detail.duplicates + '. Выбрана строка с наибольшим количеством выполненных мероприятий.', 'or-note'));
        if (detail.sheet_name) card.append(el('p', 'В источнике: ' + detail.sheet_name + (detail.sheet_omsu ? ' · ' + detail.sheet_omsu : ''), 'or-note'));
      }
      if (detail.items) {
        const disclosure = el('details'), list = el('ol');
        disclosure.append(el('summary', 'Как оценены все мероприятия (' + detail.items.length + ')'));
        detail.items.forEach(i => {
          const li = el('li'); li.value = i.n;
          li.append(el('strong', i.name), el('span', i.ok ? 'Засчитано' : 'Не засчитано', 'or-tag ' + (i.ok ? 'good' : 'bad')));
          if (i.due) li.append(el('p', 'Срок по методике: ' + i.due, 'or-note'));
          if (i.raw !== '' && i.raw != null) li.append(el('p', 'В источнике: ' + i.raw));
          if (i.reason) li.append(el('p', i.reason, 'or-note'));
          list.append(li);
        });
        disclosure.append(list); card.append(disclosure);
      }
      body.append(card);
    });
    $('or-detail').showModal();
  }
  $('or-detail-close').addEventListener('click', () => $('or-detail').close());
  $('or-refresh').addEventListener('click', async () => {
    if (offline) return;
    busy(true); alert('');
    try { await api('refresh', 'POST'); schedule(); }
    catch (e) { busy(false); alert(e.name === 'AbortError' ? 'Сервер не подтвердил запуск. Проверяем состояние — повторная загрузка не требуется.' : e.message || 'Не удалось начать обновление.'); if (e.name === 'AbortError') schedule(); }
  });
  $('or-period').addEventListener('change', e => { period = Number(e.target.value); render(); });
  $('or-search').addEventListener('input', () => { if (saved) render(); });
  $('or-all').addEventListener('change', () => { if (saved) render(); });
  window.addEventListener('pagehide', () => { clearTimeout(polling); if (exportUrl) { URL.revokeObjectURL(exportUrl); exportUrl = null; } });
  window.addEventListener('pageshow', event => { if (event.persisted) { if (offline) updateExport(); else load(); } });
  document.addEventListener('visibilitychange', () => { if (!offline && !document.hidden && saved?.refresh?.status === 'running') check(); });
  if (offline) {
    document.documentElement.dataset.ratingMode = 'snapshot';
    ['or-refresh', 'or-export-html', 'or-home'].forEach(id => { if ($(id)) $(id).hidden = true; });
    try { saved = JSON.parse(snapshotNode.textContent); render(); }
    catch (_) { alert('Не удалось прочитать данные снимка. Скачайте его заново в Нейроне.'); }
  } else load();
})();
