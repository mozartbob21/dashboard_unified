/* Read-only survey analytics. Comments are always inserted with textContent. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else if (root.document.getElementById('feedbackAdmin')) api.createController(root).start();
})(typeof window === 'undefined' ? globalThis : window, function () {
  'use strict';
  const POLL_MS = 10000;
  const REQUEST_TIMEOUT_MS = 12000;
  const ENDPOINT = '/api/ux-survey/admin';
  const integer = value => Number.isInteger(value) && value >= 0;
  const decimal = value => new Intl.NumberFormat('ru-RU', {minimumFractionDigits:1, maximumFractionDigits:1}).format(value);
  const number = value => new Intl.NumberFormat('ru-RU').format(value);
  function dateTime(value) {
    const date = new Date(value);
    return Number.isNaN(date.valueOf()) ? 'Дата не указана' : new Intl.DateTimeFormat('ru-RU', {
      timeZone:'Europe/Moscow', day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit', second:'2-digit'
    }).format(date);
  }
  function day(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return String(value);
    return value.split('-').reverse().join('.');
  }
  function validPayload(data) {
    return data && integer(data.completed_count) && integer(data.invited_count) &&
      (data.completion_rate === null || (typeof data.completion_rate === 'number' && Number.isFinite(data.completion_rate) && data.completion_rate >= 0 && data.completion_rate <= 100)) &&
      Array.isArray(data.questions) && data.questions.length === 6 && data.questions.every(q =>
        typeof q.label === 'string' && integer(q.count) && (q.average === null || (typeof q.average === 'number' && q.average >= 0 && q.average <= 5)) &&
        Array.isArray(q.distribution) && q.distribution.length === 6 && [0,1,2,3,4,5].every(score => q.distribution.filter(x => x.score === score && integer(x.count)).length === 1)) &&
      Array.isArray(data.daily) && data.daily.every(item => typeof item.date === 'string' && integer(item.count)) &&
      Array.isArray(data.comments) && data.comments.every(item => typeof item.text === 'string' && typeof item.submitted_at === 'string');
  }
  function createController(env) {
    const document = env.document;
    const node = id => document.getElementById(id);
    let started = false, stopped = false, interval = null, request = null, timeout = null;
    let inFlight = false, lastSuccess = null, fingerprint = null;
    function sectionActive() {
      const section = document.body?.dataset?.adminSection;
      const panel = node('feedbackAdmin')?.closest?.('[data-admin-panel]');
      // The standalone survey page keeps its original polling behaviour.
      return !section || !panel || (section === 'feedback' && !panel.hidden);
    }
    function element(tag, className, text) {
      const result = document.createElement(tag);
      if (className) result.className = className;
      if (text !== undefined) result.textContent = String(text);
      return result;
    }
    function bar(count, total) {
      const outer = element('span', 'feedback-bar');
      outer.setAttribute('aria-hidden', 'true');
      const fill = element('span');
      fill.style.width = `${total ? Math.min(100, 100 * count / total) : 0}%`;
      outer.append(fill);
      return outer;
    }
    function renderQuestions(data) {
      node('feedbackRatingsEmpty').hidden = data.completed_count !== 0;
      const cards = data.questions.map((question, index) => {
        const card = element('article', 'feedback-panel feedback-question');
        const heading = element('div', 'feedback-question-heading');
        const title = element('h3', '', question.label);
        title.id = `feedback-question-${index + 1}`;
        card.setAttribute('aria-labelledby', title.id);
        heading.append(element('span', 'feedback-question-number', index + 1), title);
        const score = element('div', 'feedback-question-score');
        const average = element('div', 'feedback-average', question.average === null ? '—' : decimal(question.average));
        average.setAttribute('aria-label', question.average === null ? 'Средняя оценка: пока нет ответов' : `Средняя оценка: ${decimal(question.average)} из 5`);
        average.append(element('small', '', ' / 5'));
        score.append(average, element('p', 'feedback-question-count', `Ответов: ${number(question.count)}`));
        const distribution = element('dl', 'feedback-distribution');
        distribution.setAttribute('aria-label', 'Распределение оценок от 0 до 5');
        for (let value = 0; value <= 5; value++) {
          const count = question.distribution.find(item => item.score === value).count;
          const row = element('div', 'feedback-distribution-row');
          const label = element('dt', '', value);
          const graphic = element('dd');
          graphic.append(bar(count, question.count));
          const percent = question.count ? Math.round(100 * count / question.count) : 0;
          row.append(label, graphic, element('dd', 'feedback-bar-value', `${number(count)} · ${percent}%`));
          distribution.append(row);
        }
        card.append(heading, score, distribution);
        return card;
      });
      node('feedbackQuestions').replaceChildren(...cards);
    }
    function renderDaily(data) {
      node('feedbackDailyEmpty').hidden = data.daily.length > 0;
      node('feedbackDaily').hidden = data.daily.length === 0;
      const table = element('table');
      table.append(element('caption', '', 'Отправленные анкеты по дням · московское время'));
      const head = element('thead'), row = element('tr');
      const dateHeading = element('th', '', 'Дата');dateHeading.setAttribute('scope', 'col');
      dateHeading.colSpan = 2;
      const countHeading = element('th', '', 'Ответов');countHeading.setAttribute('scope', 'col');
      row.append(dateHeading, countHeading);head.append(row);table.append(head);
      const body = element('tbody');
      const maximum = Math.max(1, ...data.daily.map(item => item.count));
      for (const item of [...data.daily].sort((a,b) => a.date.localeCompare(b.date))) {
        const dailyRow = element('tr');
        const date = element('th', '', day(item.date));date.setAttribute('scope', 'row');
        const graphic = element('td', 'feedback-daily-bar');graphic.append(bar(item.count, maximum));
        dailyRow.append(date, graphic, element('td', '', number(item.count)));body.append(dailyRow);
      }
      table.append(body);node('feedbackDaily').replaceChildren(table);
    }
    function renderComments(data) {
      const comments = data.comments.filter(item => item.text.trim()).sort((a,b) => b.submitted_at.localeCompare(a.submitted_at));
      node('feedbackCommentsEmpty').hidden = comments.length > 0;
      node('feedbackCommentsCaption').textContent = 'Непустые ответы, от новых к старым. Время московское.';
      const items = comments.map(comment => {
        const item = element('li', 'feedback-comment');
        const date = element('time', '', dateTime(comment.submitted_at));
        date.dateTime = comment.submitted_at;
        item.append(date, element('p', '', comment.text));
        return item;
      });
      node('feedbackComments').replaceChildren(...items);
    }
    function render(data) {
      // Polls without changes do not replace text being read or selected by an administrator.
      const nextFingerprint = JSON.stringify([data.completed_count, data.invited_count, data.completion_rate, data.questions, data.daily, data.comments]);
      if (nextFingerprint !== fingerprint) {
        node('feedbackCompleted').textContent = number(data.completed_count);
        node('feedbackInvited').textContent = number(data.invited_count);
        node('feedbackConversion').textContent = data.completion_rate === null ? '—' : `${decimal(data.completion_rate)}%`;
        renderQuestions(data);renderDaily(data);renderComments(data);
        fingerprint = nextFingerprint;
      }
      node('feedbackResults').hidden = false;
    }
    function status(text, state) {
      if (node('feedbackConnection').textContent !== text) node('feedbackConnection').textContent = text;
      node('feedbackConnection').dataset.state = state;
    }
    function clearRestrictedData() {
      node('feedbackResults').hidden = true;
      for (const id of ['feedbackQuestions','feedbackDaily','feedbackComments']) node(id).replaceChildren();
      for (const id of ['feedbackCompleted','feedbackInvited','feedbackConversion']) node(id).textContent = '—';
      fingerprint = null;lastSuccess = null;
      node('feedbackUpdated').textContent = 'Результаты скрыты до подтверждения прав доступа.';
    }
    async function refresh() {
      if (inFlight || stopped || document.hidden || !sectionActive()) return;
      inFlight = true;
      node('feedbackRefresh').disabled = true;
      node('feedbackRefresh').setAttribute('aria-busy','true');
      if (!lastSuccess) status('Загрузка результатов…', 'loading');
      request = new env.AbortController();
      timeout = env.setTimeout(() => request?.abort(), REQUEST_TIMEOUT_MS);
      try {
        const response = await env.fetch(ENDPOINT, {credentials:'same-origin', cache:'no-store', signal:request.signal, headers:{'Accept':'application/json'}});
        if (stopped) return;
        if (response.redirected || response.status === 401 || response.status === 403) {
          clearRestrictedData();
          node('feedbackAccessNotice').hidden = false;
          status('Доступ не подтверждён. Результаты доступны только администраторам.', 'error');
          return;
        }
        if (!response.ok) throw new Error('Unavailable');
        const data = await response.json();
        if (!validPayload(data)) throw new Error('Invalid response');
        if (stopped) return;
        render(data);
        lastSuccess = new Date();
        node('feedbackAccessNotice').hidden = true;
        status('Результаты обновляются автоматически.', 'success');
        node('feedbackUpdated').textContent = `Последнее обновление: ${dateTime(lastSuccess)} (МСК). Проверка каждые 10 секунд, пока раздел открыт и вкладка видна.`;
      } catch (_) {
        if (!stopped) status(lastSuccess ? 'Связь с сервером прервалась. Показаны последние полученные данные; повторим запрос автоматически.' : 'Не удалось загрузить результаты. Повторим запрос автоматически; можно нажать «Обновить сейчас».', 'error');
      } finally {
        env.clearTimeout(timeout);timeout = null;request = null;inFlight = false;
        node('feedbackRefresh').disabled = false;
        node('feedbackRefresh').removeAttribute('aria-busy');
      }
    }
    function resume() { if (!document.hidden) return refresh(); }
    function start() {
      if (started) return;
      started = true;stopped = false;
      node('feedbackRefresh').addEventListener('click', refresh);
      document.addEventListener('visibilitychange', resume);
      document.addEventListener('admin:sectionchange', resume);
      env.addEventListener('online', resume);
      interval = env.setInterval(refresh, POLL_MS);
      return refresh();
    }
    function stop() {
      stopped = true;
      env.clearInterval(interval);env.clearTimeout(timeout);request?.abort();
      node('feedbackRefresh').removeEventListener('click', refresh);
      document.removeEventListener('visibilitychange', resume);
      document.removeEventListener('admin:sectionchange', resume);
      env.removeEventListener('online', resume);
    }
    return {start,stop,refresh};
  }
  return {createController,validPayload,POLL_MS,REQUEST_TIMEOUT_MS};
});
