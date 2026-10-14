/* Shared, read-only totals; labels are rendered as text, never HTML. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else if (root.document.getElementById('modulePopularity')) api.createController(root).start();
})(typeof window === 'undefined' ? globalThis : window, function () {
  'use strict';
  const POLL_MS = 10000;
  const REQUEST_TIMEOUT_MS = 12000;
  const ENDPOINT = '/api/users/module-popularity';
  const integer = value => Number.isSafeInteger(value) && value >= 0;
  const date = value => typeof value === 'string' && !Number.isNaN(Date.parse(value));
  const number = value => new Intl.NumberFormat('ru-RU').format(value);
  const dateTime = value => new Intl.DateTimeFormat('ru-RU', {
    timeZone:'Europe/Moscow', day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit', second:'2-digit'
  }).format(new Date(value));
  function validPayload(data) {
    return data && data.period === 'all_time' && date(data.tracking_since) && date(data.updated_at) &&
      integer(data.visit_gap_minutes) && data.visit_gap_minutes > 0 && Array.isArray(data.modules) && data.modules.length > 0 &&
      new Set(data.modules.map(row => row?.module_id)).size === data.modules.length &&
      data.modules.every((row, index) => row && row.rank === index + 1 && typeof row.module_id === 'string' &&
        row.module_id.length > 0 && typeof row.label === 'string' &&
        ['visits','unique_users','page_views','actions'].every(key => integer(row[key])) &&
        (row.last_seen_at === null || date(row.last_seen_at)) &&
        (index === 0 || row.visits <= data.modules[index - 1].visits));
  }
  function createController(env) {
    const document = env.document;
    const node = id => document.getElementById(id);
    let started = false, stopped = false, interval = null, request = null, timeout = null;
    let inFlight = false, lastSuccess = null, fingerprint = null;
    function sectionActive() {
      const section = document.body?.dataset?.adminSection;
      const panel = node('modulePopularity')?.closest?.('[data-admin-panel]');
      // The standalone statistics page has no shared admin navigation.
      return !section || !panel || (section === 'statistics' && !panel.hidden);
    }
    function cell(tag, text, className) {
      const result = document.createElement(tag);
      result.textContent = String(text);
      if (className) result.className = className;
      return result;
    }
    function render(data) {
      // Preserve selected text and table focus when polling returns unchanged totals.
      const nextFingerprint = JSON.stringify([data.tracking_since, data.visit_gap_minutes, data.modules]);
      if (nextFingerprint !== fingerprint) {
        const rows = data.modules.map(item => {
          const row = document.createElement('tr');
          const label = cell('th', item.label);
          label.setAttribute('scope', 'row');
          row.append(cell('td', item.rank, 'popularity-rank'), label,
            cell('td', number(item.visits), 'popularity-visits'), cell('td', number(item.unique_users)),
            cell('td', number(item.page_views)), cell('td', number(item.actions)),
            cell('td', item.last_seen_at === null ? 'Нет активности' : dateTime(item.last_seen_at), 'popularity-date'));
          return row;
        });
        node('popularityRows').replaceChildren(...rows);
        node('popularitySince').textContent = `За всё время · сбор с ${dateTime(data.tracking_since)} (МСК). Блоков в рейтинге: ${number(data.modules.length)}.`;
        const overview = {
          popularityTotalVisits:data.modules.reduce((total, item) => total + BigInt(item.visits), 0n),
          popularityTrackedModules:data.modules.length,
          popularityUnusedModules:data.modules.filter(item => item.visits === 0).length
        };
        for (const [id, value] of Object.entries(overview)) {
          if (node(id)) node(id).textContent = number(value);
        }
        fingerprint = nextFingerprint;
      }
      node('popularityContent').hidden = false;
    }
    function status(text, state) {
      if (node('popularityStatus').textContent !== text) node('popularityStatus').textContent = text;
      node('popularityStatus').dataset.state = state;
    }
    function clearRestrictedData() {
      node('popularityContent').hidden = true;
      node('popularityRows').replaceChildren();
      node('popularitySince').textContent = '';
      node('popularityUpdated').textContent = 'Статистика скрыта до подтверждения прав доступа.';
      for (const id of ['popularityTotalVisits','popularityTrackedModules','popularityUnusedModules']) {
        if (node(id)) node(id).textContent = '—';
      }
      fingerprint = null;
      lastSuccess = null;
    }
    async function refresh() {
      if (inFlight || stopped || document.hidden || !sectionActive()) return;
      inFlight = true;
      node('popularityRefresh').disabled = true;
      node('popularityRefresh').setAttribute('aria-busy', 'true');
      if (!lastSuccess) status('Загрузка статистики…', 'loading');
      request = new env.AbortController();
      timeout = env.setTimeout(() => request?.abort(), REQUEST_TIMEOUT_MS);
      try {
        const response = await env.fetch(ENDPOINT, {
          credentials:'same-origin', cache:'no-store', signal:request.signal, headers:{'Accept':'application/json'}
        });
        if (stopped) return;
        if (response.redirected || response.status === 401 || response.status === 403) {
          clearRestrictedData();
          node('popularityAccessNotice').hidden = false;
          status('Для просмотра статистики нужно право управления пользователями.', 'error');
          return;
        }
        if (!response.ok) throw new Error('Unavailable');
        const data = await response.json();
        if (stopped) return;
        if (!validPayload(data)) throw new Error('Invalid response');
        render(data);
        lastSuccess = data.updated_at;
        node('popularityAccessNotice').hidden = true;
        status('Рейтинг обновляется автоматически.', 'success');
        node('popularityUpdated').textContent = `Последнее обновление: ${dateTime(lastSuccess)} (МСК). Проверка каждые 10 секунд, пока раздел открыт и вкладка видна.`;
      } catch (_) {
        if (!stopped) status(lastSuccess
          ? 'Связь с сервером прервалась. Показаны последние полученные данные; повторим запрос автоматически.'
          : 'Не удалось загрузить статистику. Повторим запрос автоматически; можно нажать «Обновить сейчас».', 'error');
      } finally {
        env.clearTimeout(timeout);
        timeout = null;request = null;inFlight = false;
        node('popularityRefresh').disabled = false;
        node('popularityRefresh').removeAttribute('aria-busy');
      }
    }
    function resume() { if (!document.hidden) return refresh(); }
    function start() {
      if (started) return;
      started = true;
      node('popularityRefresh').addEventListener('click', refresh);
      document.addEventListener('visibilitychange', resume);
      document.addEventListener('admin:sectionchange', resume);
      env.addEventListener('online', resume);
      interval = env.setInterval(refresh, POLL_MS);
      return refresh();
    }
    function stop() {
      stopped = true;
      env.clearInterval(interval);env.clearTimeout(timeout);request?.abort();
      node('popularityRefresh').removeEventListener('click', refresh);
      document.removeEventListener('visibilitychange', resume);
      document.removeEventListener('admin:sectionchange', resume);
      env.removeEventListener('online', resume);
    }
    return {start, stop, refresh};
  }
  return {createController, validPayload, POLL_MS, REQUEST_TIMEOUT_MS};
});
