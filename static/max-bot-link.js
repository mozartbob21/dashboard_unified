/* Account linking is confirmed in the authenticated platform, never by the bot alone. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else api.boot(root);
})(typeof window !== 'undefined' ? window : null, function () {
  'use strict';
  const OPERATIONS = new Set(['status', 'link', 'confirm', 'unlink']);
  class LinkError extends Error {
    constructor(code, status = 0) { super(code); this.code = code; this.status = status; }
  }
  function expiresAt(value) {
    const result = typeof value === 'string' ? Date.parse(value) : NaN;
    return Number.isFinite(result) ? result : 0;
  }
  function timeLabel(value) {
    const date = expiresAt(value);
    return date ? new Intl.DateTimeFormat('ru-RU', { hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Moscow' }).format(date) + ' мск' : '';
  }
  function errorMessage(error) {
    if (error.code === 'MAX_BOT_HTTPS_REQUIRED') return 'Для подключения MAX откройте Нейрону по HTTPS или локально на компьютере сервера. Через обычное HTTP-подключение по сети привязка недоступна.';
    if (error.status === 401 || error.status === 403) return 'Сессия завершилась или доступ изменился. Обновите страницу и войдите в Нейрону снова.';
    if (error.status === 429) return 'Слишком много попыток. Подождите немного и повторите действие.';
    if ([404, 409, 410].includes(error.status)) return 'Запрос уже завершён или истёк. Проверьте состояние подключения; при необходимости получите новый код.';
    if (error.code === 'TIMEOUT') return 'Сервер отвечает дольше обычного. Проверьте состояние подключения перед повторной попыткой.';
    return 'Не удалось связаться с сервером. Подключение не подтверждено; попробуйте снова.';
  }
  function createRequest(win, csrf, timeoutMs = 15000) {
    return async (operation, payload = {}) => {
      if (!OPERATIONS.has(operation)) throw new LinkError('INVALID_OPERATION');
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      try {
        const headers = { Accept: 'application/json' };
        const options = { method: operation === 'status' ? 'GET' : 'POST', credentials: 'same-origin', cache: 'no-store', redirect: 'error', headers, signal: controller.signal };
        if (operation !== 'status') {
          if (!csrf) throw new LinkError('CSRF_MISSING', 403);
          headers['Content-Type'] = 'application/json'; headers['X-CSRF-Token'] = csrf;
          options.body = JSON.stringify(payload);
        }
        const response = await win.fetch('/max-bot/api/' + operation, options);
        let data;
        try { data = await response.json(); } catch (_) { throw new LinkError('UNAVAILABLE', response.status); }
        if (!response.ok) {
          const detail = data && typeof data.detail === 'object' ? data.detail : data;
          throw new LinkError(detail && typeof detail.code === 'string' ? detail.code : 'HTTP_ERROR', response.status);
        }
        return data;
      } catch (error) {
        if (controller.signal.aborted) throw new LinkError('TIMEOUT');
        if (error instanceof LinkError) throw error;
        throw new LinkError('UNAVAILABLE');
      } finally { clearTimeout(timer); }
    };
  }
  class LinkPage {
    constructor(win) {
      this.win = win; this.doc = win.document;
      const meta = this.doc.querySelector('meta[name="max-bot-csrf"]');
      this.request = createRequest(win, meta ? meta.content : '');
      this.state = { enabled: false, active: false, pending: [], code: '', expires: 0, waiting: false, blocked: false };
      this.generation = 0; this.busy = false; this.loading = false; this.timer = null; this.lastPoll = 0; this.copySequence = 0;
    }
    el(id) { return this.doc.getElementById('max-bot-' + id); }
    hide(id, value) { this.el(id).hidden = value; }
    text(id, value) { this.el(id).textContent = value == null ? '' : String(value); }
    on(id, action) { this.el(id).addEventListener('click', action); }
    node(tag, cls, value) { const node = this.doc.createElement(tag); if (cls) node.className = cls; if (value != null) node.textContent = String(value); return node; }
    init() {
      this.on('create', () => this.createCode()); this.on('copy', () => this.copyCode());
      this.on('retry', () => this.loadStatus()); this.on('unlink', () => this.unlink());
      this.doc.addEventListener('visibilitychange', () => { if (!this.doc.hidden && this.state.waiting && !this.state.blocked) this.loadStatus(); });
      this.win.addEventListener('pagehide', () => { this.generation++; this.clearCode(); this.state.pending = []; this.state.waiting = false; this.loading = false; this.setBusy(false); this.el('pending-list').replaceChildren(); this.hide('pending', true); this.stop(); });
      this.win.addEventListener('pageshow', event => { if (event.persisted) { this.start(); this.loadStatus(); } });
      this.start(); this.loadStatus();
    }
    start() { if (!this.timer) this.timer = this.win.setInterval(() => this.tick(), 1000); }
    stop() { if (this.timer) this.win.clearInterval(this.timer); this.timer = null; }
    tick() {
      const now = Date.now();
      if (this.state.code && now >= this.state.expires) { this.clearCode(); this.hide('expired', false); this.text('announcement', 'Срок одноразового кода истёк.'); }
      const validPending = this.state.pending.filter(item => expiresAt(item.expires_at) > now);
      if (validPending.length !== this.state.pending.length) { this.state.pending = validPending; this.renderPending(); }
      this.state.waiting = !this.state.active && (!!this.state.code || this.state.pending.length > 0);
      if (this.state.waiting && !this.doc.hidden && !this.state.blocked && now - this.lastPoll >= 10000) this.loadStatus();
    }
    clearCode() { this.copySequence++; this.state.code = ''; this.state.expires = 0; this.text('code', ''); this.text('expiry', ''); this.hide('code-panel', true); this.text('copy', 'Скопировать'); }
    setBusy(value) {
      this.busy = value;
      this.el('create').disabled = value || this.state.blocked;
      this.el('unlink').disabled = value || this.state.blocked;
      for (const button of this.el('pending-list').querySelectorAll('button')) button.disabled = value || this.state.blocked;
    }
    clearError() { this.hide('error', true); this.hide('retry', true); }
    fail(error) {
      this.text('error', errorMessage(error)); this.hide('error', false);
      if ([401, 403].includes(error.status)) {
        this.state.blocked = true; this.state.waiting = false; this.clearCode(); this.state.pending = []; this.renderPending(); this.setBusy(false); this.hide('retry', true);
      } else this.hide('retry', false);
    }
    applyStatus(data) {
      if (!data || typeof data.enabled !== 'boolean' || typeof data.active !== 'boolean' || !Array.isArray(data.pending)) throw new LinkError('INVALID_RESPONSE');
      this.state.enabled = data.enabled; this.state.active = data.active;
      this.state.pending = data.pending.filter(item => item && typeof item.id === 'string' && typeof item.label === 'string' && expiresAt(item.expires_at) > Date.now()).slice(0, 10);
      if (data.active || !data.enabled || this.state.pending.length) this.clearCode();
      this.state.waiting = !data.active && (!!this.state.code || this.state.pending.length > 0);
      this.hide('loading', true); this.hide('disabled', data.enabled); this.hide('unlinked', !data.enabled || data.active); this.hide('active', !data.active);
      this.text('status', data.active ? 'Подключён' : data.enabled ? 'Не подключён' : 'Не настроен');
      this.el('status').classList.toggle('connected', data.active);
      this.text('active-label', typeof data.active_label === 'string' ? data.active_label : 'Доступ к опубликованным остаткам');
      const destination = this.el('name');
      if (destination) {
        const name = typeof data.bot_public_name === 'string' ? data.bot_public_name.trim().replace(/[\u0000-\u001f\u007f]/g, '').slice(0, 120) : '';
        destination.textContent = name ? 'Ваш бот в MAX: ' + name : 'Откройте в MAX бота, которого вам предоставил администратор.';
      }
      this.text('create', this.state.code ? 'Получить новый код' : 'Получить код');
      if (data.active) { this.hide('expired', true); this.text('announcement', 'Аккаунт MAX подключён.'); }
      this.renderPending();
      this.text('updated', 'Проверено в ' + new Intl.DateTimeFormat('ru-RU', { hour: '2-digit', minute: '2-digit', second: '2-digit', timeZone: 'Europe/Moscow' }).format(new Date()) + ' мск' + (this.state.waiting ? ' · проверяем каждые 10 секунд' : ''));
    }
    renderPending() {
      const target = this.el('pending-list'); target.replaceChildren();
      this.hide('pending', !this.state.pending.length || this.state.active || this.state.blocked);
      for (const item of this.state.pending) {
        const card = this.node('div', 'max-bot-pending-item');
        card.append(this.node('strong', '', item.label));
        card.append(this.node('span', 'max-bot-muted', 'Подтвердите до ' + timeLabel(item.expires_at)));
        const button = this.node('button', 'max-bot-button primary', 'Это мой аккаунт — подключить'); button.type = 'button'; button.disabled = this.busy || this.state.blocked;
        button.addEventListener('click', () => this.confirm(item.id)); card.append(button); target.append(card);
      }
    }
    async loadStatus() {
      if (this.loading || this.busy || this.state.blocked) return;
      this.loading = true; this.lastPoll = Date.now(); const generation = this.generation;
      try { const data = await this.request('status'); if (generation !== this.generation) return; this.clearError(); this.applyStatus(data); }
      catch (error) { if (generation === this.generation) { this.hide('loading', true); this.fail(error); } }
      finally { this.loading = false; }
    }
    async createCode() {
      if (this.busy || this.state.blocked || !this.state.enabled || this.state.active) return;
      this.generation++; const generation = this.generation; this.setBusy(true); this.clearError(); this.clearCode(); this.hide('expired', true);
      try {
        const data = await this.request('link');
        if (generation !== this.generation) return;
        if (!data || typeof data.code !== 'string' || !/^[A-Za-z0-9-]{4,64}$/.test(data.code) || expiresAt(data.expires_at) <= Date.now()) throw new LinkError('INVALID_RESPONSE');
        this.state.pending = []; this.renderPending();
        this.state.code = data.code; this.state.expires = expiresAt(data.expires_at); this.state.waiting = true; this.lastPoll = Date.now();
        this.text('code', '/link ' + data.code); this.text('expiry', 'Код действует до ' + timeLabel(data.expires_at)); this.hide('code-panel', false); this.text('create', 'Получить новый код');
        this.text('announcement', 'Одноразовый код создан. Отправьте команду вашему боту MAX.');
      } catch (error) { if (generation === this.generation) this.fail(error); }
      finally { if (generation === this.generation) this.setBusy(false); }
    }
    async copyCode() {
      if (!this.state.code || this.state.expires <= Date.now()) { this.clearCode(); return; }
      const sequence = ++this.copySequence;
      try {
        if (!this.win.navigator.clipboard || typeof this.win.navigator.clipboard.writeText !== 'function') throw new Error('unavailable');
        await this.win.navigator.clipboard.writeText('/link ' + this.state.code);
        if (sequence !== this.copySequence || !this.state.code) return;
        this.text('copy', 'Скопировано'); this.text('announcement', 'Команда скопирована.');
      } catch (_) {
        if (sequence !== this.copySequence || !this.state.code) return;
        this.text('announcement', 'Выделите команду и скопируйте её вручную.'); this.text('copy', 'Выделите команду');
      }
    }
    async confirm(requestId) {
      if (this.busy || this.state.blocked || !this.state.pending.some(item => item.id === requestId && expiresAt(item.expires_at) > Date.now())) return;
      await this.mutate('confirm', { request_id: requestId });
    }
    async unlink() { if (!this.busy && !this.state.blocked && this.state.active) await this.mutate('unlink'); }
    async mutate(operation, payload = {}) {
      this.generation++; const generation = this.generation; this.setBusy(true); this.clearError();
      try {
        await this.request(operation, payload);
        if (generation !== this.generation) return;
        this.clearCode();
        const status = await this.request('status');
        if (generation !== this.generation) return;
        this.applyStatus(status);
        if (operation === 'unlink') this.text('announcement', 'Доступ бота к вашему аккаунту отключён.');
      } catch (error) { if (generation === this.generation) this.fail(error); }
      finally { if (generation === this.generation) this.setBusy(false); }
    }
  }
  return { LinkError, expiresAt, timeLabel, errorMessage, createRequest, LinkPage, boot(win) { new LinkPage(win).init(); } };
});
