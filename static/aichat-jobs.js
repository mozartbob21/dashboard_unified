/* Durable chat tasks: short requests only. Never retry a command automatically. */
(function(root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.NeuronaChatJobs = api;
})(typeof window === 'undefined' ? globalThis : window, function() {
  'use strict';
  const ACTIVE = new Set(['queued', 'running']);
  const TERMINAL = new Set(['succeeded', 'failed', 'interrupted', 'timed_out']);
  const LABELS = {queued:'Принято в работу', running:'Выполняется', succeeded:'Завершено',
    failed:'Не удалось завершить', interrupted:'Выполнение прервано', timed_out:'Время выполнения истекло'};
  function requestKey(env) {
    if (env.crypto && env.crypto.randomUUID) return env.crypto.randomUUID();
    const bytes = new Uint8Array(16);
    env.crypto.getRandomValues(bytes);
    return Array.from(bytes, value => value.toString(16).padStart(2, '0')).join('');
  }
  async function request(env, url, options) {
    const controller = new env.AbortController();
    const timer = env.setTimeout(() => controller.abort(), 20000);
    try {
      const response = await env.fetch(url, {credentials:'same-origin', cache:'no-store', ...options,
        signal:controller.signal});
      if (response.status === 401 || response.redirected) {
        const error = new Error('Сессия завершилась. Войдите в систему снова.');
        error.status = 401;throw error;
      }
      let data;
      try { data = await response.json(); }
      catch (_) { throw new Error('Не удалось прочитать ответ сервера.'); }
      if (!response.ok) {
        const error = new Error(typeof data.detail === 'string' ? data.detail : 'Сервер не смог обработать запрос.');
        error.status = response.status;throw error;
      }
      return data;
    } finally { env.clearTimeout(timer); }
  }
  function create(env, options) {
    options = options || {};
    let dialog = null, generation = 0, timer = null, destroyed = false, fetching = false;
    let failures = 0, connection = '', lastUpdated = '', jobs = [], pending = new Map();
    let terminalSeen = new Set(), running = false, confirmations = new Map();
    function notify() {
      const active = jobs.some(job => ACTIVE.has(job.status));
      if (active !== running) { running = active;options.onActivity?.(active); }
      options.onChange?.({jobs:jobs.slice(), pending:Array.from(pending.values()).map(item => ({
        id:item.id, text:item.text, error:item.error, sending:item.sending, retryable:item.retryable
      })), connection, lastUpdated});
    }
    function schedule(delay) {
      env.clearTimeout(timer);
      if (!destroyed) timer = env.setTimeout(refresh, delay);
    }
    async function refresh() {
      if (destroyed || fetching) return;
      fetching = true;
      const version = generation;
      try {
        const data = await request(env, '/aichat/api/jobs' + (dialog ? '?dialog_id=' + encodeURIComponent(dialog) : ''));
        if (destroyed || version !== generation) return;
        jobs = (Array.isArray(data.items) ? data.items : []).filter(job => job && typeof job.job_id === 'string' &&
          (ACTIVE.has(job.status) || TERMINAL.has(job.status)));
        failures = 0;connection = '';lastUpdated = new Date().toISOString();
        const completed = jobs.filter(job => TERMINAL.has(job.status) && !terminalSeen.has(job.job_id));
        notify();
        if (completed.length) await options.onTerminal?.(completed);
        if (version === generation) completed.forEach(job => terminalSeen.add(job.job_id));
      } catch(error) {
        if (version !== generation || destroyed) return;
        failures++;
        connection = error.status === 401 ? error.message : 'Связь с сервером прервана. Проверим статус снова; повторного запуска не будет.';
        notify();
        if (error.status === 401) { destroyed = true;options.onActivity?.(false); }
      } finally {
        fetching = false;
        schedule(failures ? Math.min(10000, 2000 * Math.pow(2, Math.min(failures - 1, 3))) :
          jobs.some(job => ACTIVE.has(job.status)) ? 2000 : 10000);
      }
    }
    async function deliver(item) {
      if (item.sending) return null;
      item.sending = true;item.error = '';notify();
      try {
        const ticket = await request(env, '/aichat/api/send', {method:'POST', body:item.body,
          headers:{Accept:'application/json'}});
        if (!ticket || typeof ticket.job_id !== 'string' || typeof ticket.dialog_id !== 'string') {
          throw new Error('Сервер не подтвердил принятие задачи.');
        }
        pending.delete(item.id);
        if (dialog === ticket.dialog_id || dialog === null) {
          jobs = [...jobs.filter(job => job.job_id !== ticket.job_id), {...ticket, status:ticket.status || 'queued'}];
        }
        await options.onAccepted?.(ticket, item.id);
        notify();schedule(0);
        return ticket;
      } catch(error) {
        item.sending = false;
        item.retryable = !error.status || error.status >= 500 || error.status === 429;
        item.error = item.retryable ? 'Ответ о приёме не получен. Проверьте список задач или повторите отправку этой кнопкой — идентификатор запроса сохранён.' : error.message;
        notify();schedule(0);
        return null;
      }
    }
    function submit(body, id, text) {
      id = id || requestKey(env);
      if (!pending.has(id)) {
        body.set('background', 'true');body.set('request_id', id);
        pending.set(id, {id, body, text:String(text || 'Сообщение с вложением').slice(0, 120),
          error:'', sending:false, retryable:true});
      }
      return deliver(pending.get(id));
    }
    async function confirm(id) {
      if (confirmations.get(id)) return null;
      confirmations.set(id, true);notify();
      try {
        const ticket = await request(env, '/aichat/api/jobs/' + encodeURIComponent(id) + '/confirm', {
          method:'POST', headers:{'Content-Type':'application/json'}, body:'{}'});
        if (!ticket || typeof ticket.job_id !== 'string' || typeof ticket.dialog_id !== 'string') {
          throw new Error('Сервер не подтвердил состояние проверки. Повторите подтверждение, чтобы получить сохранённый статус.');
        }
        jobs = jobs.map(job => job.job_id === id ? {...job, result:{...job.result,
          action:{...job.result.action, confirmed_job_id:ticket.job_id}}} : job);
        await options.onAccepted?.(ticket);
        connection = '';schedule(0);return ticket;
      } catch(error) {
        connection = error.status ? error.message : 'Подтверждение запуска не получено. Можно повторить подтверждение: сервер сохранит одну задачу.';
        return null;
      } finally { confirmations.delete(id);notify(); }
    }
    function select(id) {
      dialog = id || null;generation++;jobs = [];terminalSeen = new Set();notify();schedule(0);
    }
    function online() { schedule(0); }
    env.addEventListener?.('online', online);
    const api = {submit, confirm, select, refresh, requestKey:() => requestKey(env),
      retry:id => pending.has(id) ? deliver(pending.get(id)) : Promise.resolve(null),
      dismiss:id => { const item = pending.get(id);if (item && !item.sending && !item.retryable) { pending.delete(id);options.onDismissed?.(id);notify(); } },
      confirming:id => confirmations.has(id),
      destroy:() => { destroyed = true;env.clearTimeout(timer);env.removeEventListener?.('online', online);options.onActivity?.(false); }
    };
    schedule(0);return api;
  }
  function render(document, host, state, controller, openDialog) {
    // Keep stable controls during polling so keyboard focus is never replaced.
    const focused = document.activeElement;
    const focusKey = host.contains?.(focused) ? focused.getAttribute('data-control') : null;
    const fragment = document.createDocumentFragment();
    function el(tag, className, text) {
      const node = document.createElement(tag);node.className = className || '';
      if (text != null) node.textContent = text;return node;
    }
    function button(text, key, click, disabled) {
      const node = el('button', 'ac-job-button', text);node.type = 'button';
      node.setAttribute('data-control', key);node.disabled = !!disabled;
      node.addEventListener('click', click);return node;
    }
    const cards = state.jobs.filter(job => ACTIVE.has(job.status) || (job.result?.action && !job.result.action.confirmed_job_id));
    for (const job of cards) {
      const card = el('section', 'ac-job-card');
      const action = job.result?.action;
      card.appendChild(el('strong', 'ac-job-title', action ? action.label || 'Запуск блока' : LABELS[job.status]));
      card.appendChild(el('p', '', action ? action.summary || 'Проверьте параметры и подтвердите запуск.' : job.stage || LABELS[job.status]));
      const meta = el('div', 'ac-job-meta', 'Задача ' + job.job_id.slice(0, 8));card.appendChild(meta);
      if (action) card.appendChild(button(controller.confirming(job.job_id) ? 'Подтверждаем…' : 'Запустить блок',
        'confirm-' + job.job_id, () => controller.confirm(job.job_id), controller.confirming(job.job_id)));
      if (openDialog) card.appendChild(button('Открыть диалог', 'open-' + job.job_id, () => openDialog(job.dialog_id)));
      fragment.appendChild(card);
    }
    for (const item of state.pending) {
      const card = el('section', 'ac-job-card ac-job-pending');
      card.appendChild(el('strong', 'ac-job-title', item.sending ? 'Отправляем запрос…' : 'Отправка требует внимания'));
      card.appendChild(el('p', '', item.text));
      if (item.error) card.appendChild(el('p', 'ac-job-error', item.error));
      if (!item.sending && item.retryable) card.appendChild(button('Повторить отправку безопасно',
        'retry-' + item.id, () => controller.retry(item.id)));
      if (!item.sending && !item.retryable) card.appendChild(button('Скрыть', 'dismiss-' + item.id, () => controller.dismiss(item.id)));
      fragment.appendChild(card);
    }
    if (state.connection) fragment.appendChild(el('p', 'ac-job-connection', state.connection));
    if (cards.length && state.lastUpdated) fragment.appendChild(el('p', 'ac-job-updated', 'Статус проверен в ' +
      new Date(state.lastUpdated).toLocaleTimeString('ru-RU', {hour:'2-digit', minute:'2-digit', second:'2-digit'})));
    host.replaceChildren(fragment);host.hidden = !host.childNodes.length;
    if (focusKey) {
      const control = Array.from(host.querySelectorAll('[data-control]')).find(node => node.getAttribute('data-control') === focusKey);
      if (control && !control.disabled) control.focus({preventScroll:true});
    }
  }
  return {create, render, requestKey, ACTIVE, TERMINAL};
});
