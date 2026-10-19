/* Same-origin transport and accessible controls for the MinZhKH SMS workspace. */
(function (root) {
  'use strict';
  const BASE = '/mingkh/sms/api';
  const HEAVY = new Set(['/build', '/daily', '/rez/data', '/sys/run', '/rm/measures']);
  const escape = value => String(value == null ? '' : value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  function csvCell(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : '';
    const text = String(value == null ? '' : value);
    return /^[ \t\r\n\uFEFF]*[=+@-]|^[\t\r\n]/.test(text) ? "'" + text : text;
  }
  function safeLink(value) {
    try { const url = new URL(String(value)); return ['https:', 'http:'].includes(url.protocol) && !url.username && !url.password ? url.href : '#'; }
    catch (_) { return '#'; }
  }
  function endpoint(path) {
    if (typeof path !== 'string') throw new Error('Некорректный адрес запроса.');
    if (path.startsWith('/api/')) path = BASE + path.slice(4);
    if (!path.startsWith(BASE + '/') || /[\\\s]/.test(path) || path.includes('..')) throw new Error('Некорректный адрес запроса.');
    return path;
  }
  function requestKey(url, body) {
    let h = 2166136261;
    for (const ch of url + '\n' + String(body || '')) { h ^= ch.charCodeAt(0); h = Math.imul(h, 16777619); }
    return 'neurona.sms.job.' + (h >>> 0).toString(16);
  }
  function createClient(config) {
    const transport = config.fetch;
    const pause = config.sleep || (ms => new Promise(resolve => setTimeout(resolve, ms)));
    const storage = config.storage;
    const progress = config.progress || (() => {});
    const now = config.now || Date.now;
    function remember(key, ticket) { try { storage?.setItem(key, JSON.stringify(ticket)); } catch (_) {} }
    function forget(key) { try { storage?.removeItem(key); } catch (_) {} }
    function saved(key) {
      try { const ticket = JSON.parse(storage?.getItem(key) || 'null'); return ticket && now() - ticket.created < 86400000 ? ticket : null; }
      catch (_) { return null; }
    }
    async function read(url, options) {
      const abort = new AbortController();
      const timeout = setTimeout(() => abort.abort(), 45000);
      let response;
      try { response = await transport(endpoint(url), {...options, credentials:'same-origin', cache:'no-store', signal:abort.signal, headers:{Accept:'application/json', ...options?.headers}}); }
      catch (_) { const err = new Error('Связь с сервером прервалась. Проверьте подключение и повторите действие.'); err.network = true; throw err; }
      finally { clearTimeout(timeout); }
      if (response.redirected || response.status === 401) { const error=new Error('Сессия завершилась. Войдите в Нейрону снова.'); error.status=401; throw error; }
      let data;
      try { data = await response.json(); }
      catch (_) { throw new Error('Сервер вернул неполный ответ. Повторите действие позже.'); }
      if (!response.ok) {
        const detail = data?.error || data?.detail;
        const message = typeof detail === 'string' ? detail : response.status === 403 ? 'Недостаточно прав для этого действия.' : 'Не удалось выполнить запрос. Повторите позже.';
        const err = new Error(message); err.status=response.status; err.retryable = response.status >= 500; throw err;
      }
      return {status: response.status, data};
    }
    async function get(url) {
      for (let attempt = 0; ; attempt++) {
        try { return await read(url); }
        catch (error) {
          if (!(error.network || error.retryable) || attempt >= 3) throw error;
          progress('Восстанавливаем связь с сервером…');
          await pause(1000 * 2 ** attempt);
        }
      }
    }
    async function poll(ticket, key) {
      // Each GET is bounded; the server operation is not cancelled after a minute.
      const url = endpoint(ticket.status_url);
      if (!url.startsWith(BASE + '/jobs/')) throw new Error('Сервер вернул некорректный номер операции.');
      for (;;) {
        let data;
        try { ({data} = await get(url)); }
        catch (error) { if ([401,403,404].includes(error.status)) forget(key); progress('Связь прервана. Повторите формирование, чтобы проверить уже запущенный сбор.'); throw error; }
        if (data.status === 'success') { forget(key); progress(''); return data.result; }
        if (data.status === 'error') { forget(key); progress(''); throw new Error(typeof data.error === 'string' ? data.error : 'Не удалось завершить сбор данных.'); }
        if (!['queued','running'].includes(data.status)) throw new Error('Не удалось определить состояние операции. Повторите действие.');
        progress(data.status === 'queued' ? 'Запрос в очереди. Результат появится автоматически.' : 'Сбор данных продолжается. Можно дождаться результата на этой странице.');
        await pause(2000);
      }
    }
    async function fetchAPI(path, options = {}) {
      const url = endpoint(path), method = String(options.method || 'GET').toUpperCase();
      const heavy = method === 'POST' && HEAVY.has(url.slice(BASE.length));
      let data;
      if (heavy) {
        const key = requestKey(url, options.body);
        const fingerprint = url + '\n' + String(options.body || '');
        let ticket = saved(key);
        if (ticket?.fingerprint !== fingerprint) ticket = null;
        if (!ticket) {
          ticket = {request_id: config.uuid?.() || root.crypto?.randomUUID?.() || 'sms-' + now() + '-' + Math.random().toString(36).slice(2), created:now(), fingerprint};
          remember(key, ticket);
        }
        if (!ticket.status_url) {
          let response;
          for (let attempt = 0; ; attempt++) {
            try { response = await read(url, {...options, headers:{...options.headers, 'Idempotency-Key':ticket.request_id}}); break; }
            catch (error) {
              if (!(error.network || error.retryable)) { forget(key); throw error; }
              if (attempt >= 2) throw error;
              progress('Восстанавливаем связь с сервером…'); await pause(1000 * 2 ** attempt);
            }
          }
          data = response.data;
          if (response.status === 202) {
            if (!data.job_id || !data.status_url) throw new Error('Сервер не подтвердил запуск сбора.');
            ticket = {...ticket, job_id:data.job_id, status_url:endpoint(data.status_url)};
            remember(key, ticket);
          } else { forget(key); ticket = null; }
        }
        if (ticket) data = await poll(ticket, key);
      } else data = (method === 'GET' ? await get(url) : await read(url, options)).data;
      if (url === BASE + '/options' && config.warning) config.warning(typeof data?.warning === 'string' ? data.warning : '');
      return {ok:true, status:200, json:async () => data};
    }
    return {fetch:fetchAPI};
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = {escape, csvCell, safeLink, endpoint, requestKey, createClient};
  if (!root.document) return;
  const doc = root.document;
  const byId = id => doc.getElementById(id);
  function progress(message) { const node = byId('smsConnection'); if (node) { node.textContent = message; node.hidden = !message; } }
  let storage;
  try { storage = root.sessionStorage; } catch (_) {}
  const SMS = root.SMS = {...createClient({fetch:root.fetch.bind(root), storage, progress, warning:message=>{const node=byId('smsOptionsWarning');if(node){node.textContent=message;node.hidden=!message;}}}), escape, csvCell, safeLink};
  SMS.showError = error => {
    const node = byId('msg'); if (node) { node.className='status err'; node.textContent=error.message || 'Не удалось загрузить данные.'; }
  };
  SMS.initialize = callback => {
    const go = byId('go');
    const controls=[...doc.querySelectorAll('.wrap button,.wrap input,.wrap select')];
    controls.forEach(control=>control.disabled=true);
    SMS.ready = Promise.resolve().then(callback).then(() => { controls.forEach(control=>control.disabled=false); });
    SMS.ready.catch(error => {
      SMS.showError(error);
      if (!doc.querySelector('.sms-init-retry')) {
        const retry = doc.createElement('button'); retry.className='mini sms-init-retry'; retry.type='button'; retry.textContent='Повторить загрузку настроек';
        retry.onclick = () => { retry.remove(); SMS.initialize(callback); };
        byId('msg')?.after(retry);
      }
    });
    return SMS.ready;
  };
  SMS.bind = () => {
    for (const id of ['go','dGo']) {
      const button=byId(id); if (!button?.onclick) continue;
      const handler = button.onclick; let busy = false;
      button.onclick = async event => {
        if (busy || button.disabled) return;
        busy=true; button.setAttribute('aria-busy','true');
        try { return await handler.call(button, event); }
        finally { busy=false; button.removeAttribute('aria-busy'); progress(''); }
      };
    }
    doc.querySelectorAll('.twrap').forEach(node => { node.tabIndex=0; node.setAttribute('role','region'); node.setAttribute('aria-label','Таблица результатов, прокручивается по горизонтали'); });
    SMS.run = async fields => {
      await SMS.ready;
      for (const [id,value] of Object.entries(fields)) {
        const field=byId(id); if (!field) throw new Error('Не найдено поле периода.');
        field.value=value; field.dispatchEvent(new Event('change', {bubbles:true}));
      }
      if (fields.date && !fields.from && byId('from')) byId('from').value=fields.date.slice(0,4)+'-01-01';
      const go=byId('go'); if (!go || go.disabled) throw new Error('Формирование уже выполняется.');
      const output=byId('out'); if (output) output.value='';
      await go.onclick();
      const message=byId('msg');
      if (message?.classList.contains('err') || !output?.value.trim()) throw new Error(message?.textContent || 'Источник не вернул текст.');
      return {text:output.value.trim(), message:message?.textContent || 'Готово'};
    };
  };
  const importer=byId('smsImport');
  if (importer) {
    const input=byId('smsArchive'), selected=byId('smsArchiveName'), submit=byId('smsImportSubmit'), status=byId('smsImportStatus');
    input.onchange=()=>{ selected.textContent=input.files[0]?.name || 'Архив не выбран'; submit.disabled=!input.files.length; };
    importer.onsubmit=async event=>{
      event.preventDefault(); const file=input.files[0]; if (!file || submit.disabled) return;
      if (!/\.zip$/i.test(file.name)) { status.textContent='Выберите ZIP-архив для переноса.'; return; }
      submit.disabled=true; status.className='status'; status.textContent='Переносим сохранённые данные…';
      try {
        const result=await (await SMS.fetch('/api/import', {method:'POST', headers:{'Content-Type':'application/zip'}, body:file})).json();
        status.textContent=result.message || 'Данные перенесены. Можно продолжить работу с реестрами.';
        input.value=''; selected.textContent='Архив не выбран';
        doc.dispatchEvent(new Event('sms:imported'));
        const refresh=doc.createElement('button'); refresh.type='button'; refresh.className='mini'; refresh.textContent='Обновить раздел'; refresh.onclick=()=>root.location.reload(); status.after(refresh);
      } catch(error) { status.className='status err'; status.textContent=error.message; submit.disabled=false; }
    };
  }
})(typeof window !== 'undefined' ? window : globalThis);
