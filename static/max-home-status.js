(function (global) {
  'use strict';
  class MaxHomeStatus {
    constructor(win) {
      this.win = win;
      this.links = Array.from(win.document.querySelectorAll('[data-max-home-link]'));
      this.pending = false;
      this.timer = null;
      this.lastCheck = 0;
    }
    render(active) {
      this.links.forEach(link => {
        const label = link.querySelector('[data-max-home-label]');
        if (label) label.textContent = active === true ? 'Подключено' : active === false ? 'Подключить MAX' : 'MAX';
        link.dataset.maxState = active === true ? 'connected' : active === false ? 'disconnected' : 'unknown';
        link.setAttribute('aria-label', active === true ? 'MAX подключён. Управление подключением' : 'Открыть подключение MAX');
        link.title = active === true ? 'MAX подключён к вашему аккаунту' : active === false ? 'Подключить MAX к вашему аккаунту' : 'Открыть состояние подключения MAX';
      });
    }
    async refresh(force = false) {
      if (this.pending || this.win.document.hidden || !this.links.length) return;
      if (!force && Date.now() - this.lastCheck < 5000) return;
      this.pending = true;
      const controller = new AbortController();
      const timeout = this.win.setTimeout(() => controller.abort(), 8000);
      try {
        const response = await this.win.fetch('/max-bot/api/status', {
          credentials: 'same-origin', cache: 'no-store', redirect: 'error',
          headers: {Accept: 'application/json'}, signal: controller.signal
        });
        if (!response.ok || response.redirected) throw new Error('status_unavailable');
        const state = await response.json();
        if (typeof state.active !== 'boolean' || typeof state.enabled !== 'boolean') throw new Error('invalid_status');
        this.render(state.enabled && state.active);
      } catch (_) { this.render(null); }
      finally {
        this.win.clearTimeout(timeout);
        this.pending = false;
        this.lastCheck = Date.now();
      }
    }
    init() {
      if (!this.links.length) return;
      this.render(null);
      this.refresh(true);
      this.win.addEventListener('focus', () => this.refresh());
      this.win.addEventListener('pageshow', () => this.refresh(true));
      this.win.document.addEventListener('visibilitychange', () => this.refresh());
      this.timer = this.win.setInterval(() => this.refresh(), 30000);
      this.win.addEventListener('pagehide', () => { this.win.clearInterval(this.timer); this.timer = null; });
      this.win.addEventListener('pageshow', () => {
        if (this.timer === null) this.timer = this.win.setInterval(() => this.refresh(), 30000);
      });
    }
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = {MaxHomeStatus};
  if (global.document) {
    const boot = () => new MaxHomeStatus(global).init();
    if (global.document.readyState === 'loading') global.document.addEventListener('DOMContentLoaded', boot, {once: true}); else boot();
  }
})(typeof window !== 'undefined' ? window : globalThis);
