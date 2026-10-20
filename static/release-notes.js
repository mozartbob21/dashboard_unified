/* Release content is filtered by the server. Read receipts belong to the account. */
(function (global) {
  'use strict';
  class ReleaseAPI {
    constructor(fetcher) { this.fetcher = fetcher; }
    async request(path, method = 'GET') {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 10000);
      try {
        const response = await this.fetcher('/api/releases' + path, {
          method, credentials: 'same-origin', cache: 'no-store', signal: controller.signal,
          headers: {Accept: 'application/json'}, ...(method === 'POST' ? {keepalive: true} : {})
        });
        if (!response.ok || response.redirected) throw new Error('release_request_failed');
        return await response.json();
      } finally { clearTimeout(timer); }
    }
    async list(history = false) {
      const data = await this.request(history ? '/history' : '/pending');
      if (!Array.isArray(data.releases)) throw new Error('invalid_releases');
      return data.releases.filter(item => typeof item.id === 'string' && Array.isArray(item.entries) && item.entries.length);
    }
    async read(id) {
      const data = await this.request('/' + encodeURIComponent(id) + '/read', 'POST');
      if (data.ok !== true) throw new Error('receipt_not_saved');
    }
  }
  function safeToShow(doc, edited, lastActivity, now = Date.now()) {
    if (edited || doc.visibilityState !== 'visible' || !doc.hasFocus() || now - lastActivity < 5000) return false;
    if (doc.activeElement?.closest('input, textarea, select, [contenteditable="true"], form')) return false;
    return !Array.from(doc.querySelectorAll('dialog[open], [role="dialog"], [aria-busy="true"], [aria-expanded="true"], .modal.show, .modal.open, .whats-new-modal.is-open'))
      .some(node => !node.hidden && node.getClientRects().length && global.getComputedStyle(node).visibility !== 'hidden');
  }
  // A failed POST is never treated as an acknowledgement. Successful IDs need not be resent.
  class ReadSession {
    constructor(api, releases) { this.api = api; this.ids = releases.map(item => item.id); this.saved = new Set(); this.inflight = null; }
    acknowledge() {
      if (this.inflight) return this.inflight;
      this.inflight = (async () => {
        for (const id of this.ids) {
          if (!this.saved.has(id)) { await this.api.read(id); this.saved.add(id); }
        }
      })().finally(() => { this.inflight = null; });
      return this.inflight;
    }
  }
  function node(tag, cls, text) {
    const result = document.createElement(tag); result.className = cls;
    if (text !== undefined) result.textContent = text;
    return result;
  }
  function showDialog(api, releases, history = false, loadError = false) {
    if (document.getElementById('releaseNotesDialog')) return;
    const previous = document.activeElement;
    const dialog = node('dialog', 'release-notes'); dialog.id = 'releaseNotesDialog';
    dialog.setAttribute('aria-labelledby', 'releaseNotesTitle');
    dialog.setAttribute('aria-describedby', 'releaseNotesIntro');
    const header = node('div', 'release-notes-head');
    const heading = node('div', '');
    const version = releases[0]?.version;
    heading.append(node('span', 'release-notes-kicker', 'НЕЙРОНА' + (version ? ' · ВЕРСИЯ ' + version : '')));
    const title = node('h2', '', history ? 'Что нового в сервисе' : 'Сервис стал удобнее'); title.id = 'releaseNotesTitle';
    const intro = node('p', '', 'Общие улучшения сервиса и новости доступных вам разделов.'); intro.id = 'releaseNotesIntro';
    heading.append(title, intro);
    const close = node('button', 'release-notes-close', '×'); close.type = 'button'; close.setAttribute('aria-label', 'Закрыть обновления');
    header.append(heading, close);
    const body = node('div', 'release-notes-body'); body.tabIndex = 0; body.setAttribute('aria-label', 'Список изменений');
    releases.forEach(release => {
      const section = node('section', 'release-notes-release');
      section.append(node('div', 'release-notes-date', new Date(release.published_at).toLocaleDateString('ru-RU', {day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Moscow'})));
      section.append(node('h3', '', release.title));
      const list = node('ul', 'release-notes-list');
      release.entries.forEach(entry => {
        const item = node('li', 'release-notes-entry');
        item.append(node('span', 'release-notes-module', entry.module_name || 'Платформа'));
        item.append(node('h4', '', entry.title), node('p', '', entry.text)); list.append(item);
      });
      section.append(list); body.append(section);
    });
    if (!releases.length) body.append(node('p', 'release-notes-empty', loadError
      ? 'Не удалось загрузить список обновлений. Проверьте соединение и откройте «Что нового» ещё раз.'
      : 'Для ваших разделов пока нет опубликованных обновлений.'));
    const footer = node('div', 'release-notes-footer');
    const status = node('p', 'release-notes-status'); status.setAttribute('role', 'status'); status.setAttribute('aria-live', 'polite');
    const skip = node('button', 'release-notes-secondary', 'Закрыть без отметки'); skip.type = 'button'; skip.hidden = true;
    const done = node('button', 'release-notes-primary', history ? 'Закрыть' : 'Понятно'); done.type = 'button';
    footer.append(status, skip, done); dialog.append(header, body, footer); document.body.append(dialog);
    const receipt = new ReadSession(api, releases);
    let saving = false;
    function remove() { dialog.close(); dialog.remove(); previous?.focus?.(); }
    async function dismiss() {
      if (saving) return;
      if (!releases.length) { remove(); return; }
      saving = true; done.disabled = close.disabled = true; done.textContent = 'Сохраняем…'; status.textContent = '';
      try { await receipt.acknowledge(); remove(); }
      catch (_) {
        status.textContent = 'Не удалось сохранить прочтение. Повторите попытку или закройте окно — обновление останется непрочитанным.';
        done.textContent = 'Повторить'; skip.hidden = false;
      } finally { saving = false; done.disabled = close.disabled = false; }
    }
    close.addEventListener('click', dismiss); done.addEventListener('click', dismiss); skip.addEventListener('click', remove);
    dialog.addEventListener('cancel', event => { event.preventDefault(); dismiss(); });
    try { dialog.showModal(); done.focus(); } catch (_) { dialog.remove(); }
  }
  function mount() {
    if (global.top !== global.self || typeof HTMLDialogElement === 'undefined' || global.NeuronaReleaseNotes) return;
    const api = new ReleaseAPI(global.fetch.bind(global));
    let edited = false, lastActivity = Date.now(), checking = false, finished = false, retryAt = 0;
    const record = () => { lastActivity = Date.now(); };
    ['pointerdown', 'keydown', 'scroll'].forEach(name => document.addEventListener(name, record, {passive: true}));
    document.addEventListener('input', event => { if (!event.target.closest('.release-notes')) edited = true; }, {once: true});
    async function openHistory() {
      if (document.getElementById('releaseNotesDialog')) return;
      finished = true;
      global.NeuronaProfileMenu?.close(true);
      document.querySelectorAll('#profileMenuDropdown').forEach(menu => { menu.hidden = true; menu.classList.remove('open'); });
      document.querySelectorAll('#profileMenuButton').forEach(button => button.setAttribute('aria-expanded', 'false'));
      try { showDialog(api, await api.list(true), true); } catch (_) { showDialog(api, [], true, true); }
    }
    global.NeuronaReleaseNotes = {open: openHistory};
    document.querySelectorAll('[data-release-notes-open]').forEach(button => button.addEventListener('click', openHistory));
    const timer = setInterval(async () => {
      if (finished) { clearInterval(timer); return; }
      if (checking || Date.now() < retryAt || !safeToShow(document, edited, lastActivity)) return;
      checking = true;
      try {
        const releases = await api.list();
        if (!releases.length) { finished = true; return; }
        if (!safeToShow(document, edited, lastActivity)) return;
        finished = true; showDialog(api, releases);
      } catch (_) { retryAt = Date.now() + 60000; }
      finally { checking = false; }
    }, 2000);
    global.addEventListener('pagehide', () => clearInterval(timer), {once: true});
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = {ReleaseAPI, ReadSession, safeToShow};
  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount, {once: true}); else mount();
  }
})(typeof window !== 'undefined' ? window : globalThis);
