/* This opt-in guide explains linking. It never creates a code or confirms an account. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else api.boot(root);
})(typeof window !== 'undefined' ? window : null, function () {
  'use strict';
  const STEPS = Object.freeze([
    { target: 'max-guide-code', title: 'Получите одноразовый код', text: 'В карточке «Подключение аккаунта» нажмите «Получить код». Нейрона подготовит команду для вашего аккаунта.', hint: 'Код действует 5 минут. Если срок истёк, можно получить новый. Логин и пароль отправлять боту не нужно.' },
    { target: 'max-guide-send', title: 'Отправьте команду в MAX', text: 'Нажмите «Скопировать», откройте личный чат с ботом вашей организации в MAX и отправьте команду /link вместе с кодом.', hint: 'Отправляйте команду только своему боту, не в общий чат. Если бот ещё не добавлен, его название подскажет администратор.' },
    { target: 'max-guide-confirm', title: 'Подтвердите свой аккаунт', text: 'Вернитесь на эту страницу. После отправки кода здесь появится запрос. Сверьте аккаунт и нажмите «Это мой аккаунт — подключить».', hint: 'Запрос проверяется автоматически каждые 10 секунд. Если запрос не ваш, не подтверждайте его — получите новый код.' },
    { target: 'max-guide-search', title: 'Откройте свои функции', text: 'Отправьте /menu в личный чат с ботом. Для презентаций выберите соответствующую кнопку и включите личную подписку. Для поиска остатков напишите название позиции.', hint: 'Бот показывает только доступные вашему аккаунту функции. Подписку можно выключить в том же меню. Команда /help покажет подсказки.' },
  ]);
  const clamp = (n, min, max) => Math.max(min, Math.min(n, Math.max(min, max)));
  function placement(rect, card, viewport) {
    const margin = 16, gap = 30, width = Math.min(card.width, viewport.width - margin * 2), height = card.height;
    const fitY = value => clamp(value, margin, viewport.height - height - margin);
    const fitX = value => clamp(value, margin, viewport.width - width - margin);
    if (rect.left - width - gap >= margin) return { x: rect.left - width - gap, y: fitY(rect.top), side: 'left' };
    if (rect.right + width + gap <= viewport.width - margin) return { x: rect.right + gap, y: fitY(rect.top), side: 'right' };
    if (rect.bottom + height + gap <= viewport.height - margin) return { x: fitX(rect.left), y: rect.bottom + gap, side: 'bottom' };
    if (rect.top - height - gap >= margin) return { x: fitX(rect.left), y: rect.top - height - gap, side: 'top' };
    return { x: fitX((viewport.width - width) / 2), y: Math.max(margin, viewport.height - height - margin), side: 'bottom' };
  }
  function visible(element) { return !!element && !element.hidden && element.getClientRects().length > 0; }
  class Tour {
    constructor(win) {
      this.win = win; this.doc = win.document; this.index = 0; this.open = false; this.previous = null; this.frame = null; this.pending = null; this.cancelled = false;
      this.layer = this.el('layer'); this.card = this.el('card');
    }
    el(name) { return this.doc.getElementById('max-tour-' + name); }
    target() {
      const controls = ['max-bot-create', 'max-bot-code-panel', 'max-bot-pending'];
      const control = this.doc.getElementById(controls[this.index] || 'max-guide-search');
      return visible(control) && !control.disabled ? control : this.doc.getElementById(STEPS[this.index].target);
    }
    otherDialog() { return Array.from(this.doc.querySelectorAll('dialog[open], [role="dialog"][aria-modal="true"], .modal.show, .modal.open')).some(node => node !== this.card && visible(node)); }
    init() {
      for (const button of this.doc.querySelectorAll('[data-max-tour-start]')) button.addEventListener('click', () => this.start(button));
      this.el('close').addEventListener('click', () => this.close());
      this.el('prev').addEventListener('click', () => this.show(this.index - 1));
      this.el('next').addEventListener('click', () => this.index === STEPS.length - 1 ? this.finish() : this.show(this.index + 1));
      this.doc.addEventListener('keydown', event => {
        if (!this.open || this.otherDialog()) return;
        if (event.key === 'Escape') { event.preventDefault(); this.close(); }
        else if (this.card.contains(this.doc.activeElement) && ['ArrowLeft', 'ArrowRight'].includes(event.key)) { event.preventDefault(); this.show(this.index + (event.key === 'ArrowRight' ? 1 : -1)); }
      });
      const reposition = () => this.schedulePosition();
      this.win.addEventListener('resize', reposition); this.win.addEventListener('scroll', reposition, { passive: true });
      this.win.visualViewport?.addEventListener('resize', reposition);
      if (this.win.MutationObserver) new this.win.MutationObserver(reposition).observe(this.doc.getElementById('max-bot-main'), { subtree: true, childList: true, attributes: true, attributeFilter: ['hidden'] });
      this.win.addEventListener('pagehide', () => this.close(false));
      this.requested = new URLSearchParams(this.win.location.search).get('tour') === '1';
      this.doc.addEventListener('visibilitychange', () => { if (this.requested && !this.open && !this.cancelled && !this.doc.hidden && this.pending === null) this.requestStart(); });
      if (this.requested) this.requestStart();
      return this;
    }
    requestStart() {
      this.pending = null;
      if (this.cancelled || this.doc.hidden) return;
      if (this.otherDialog()) { this.pending = this.win.setTimeout(() => this.requestStart(), 600); return; }
      this.pending = null; this.start(this.doc.querySelector('[data-max-tour-start]'));
    }
    start(launcher) {
      if (this.otherDialog()) return;
      this.previous = launcher || this.doc.activeElement; this.open = true; this.layer.hidden = false; this.doc.body.classList.add('max-tour-open'); this.show(0); this.card.focus({ preventScroll: true });
    }
    show(index) {
      if (!this.open) return;
      this.index = clamp(index, 0, STEPS.length - 1);
      const step = STEPS[this.index];
      this.el('title').textContent = step.title; this.el('text').textContent = step.text; this.el('hint').textContent = step.hint;
      if (this.index === 1) {
        const destination = this.doc.getElementById('max-bot-name');
        if (destination && destination.textContent.startsWith('Ваш бот в MAX: ')) this.el('hint').textContent = destination.textContent + '. Отправляйте команду в личный чат, не в группу.';
      }
      this.el('count').textContent = 'Шаг ' + (this.index + 1) + ' из ' + STEPS.length;
      this.el('announcement').textContent = 'Шаг ' + (this.index + 1) + '. ' + step.title;
      this.el('prev').disabled = this.index === 0; this.el('next').textContent = this.index === STEPS.length - 1 ? 'Понятно, начать' : 'Далее →';
      this.card.querySelectorAll('.max-tour-progress span').forEach((node, i) => node.classList.toggle('is-done', i <= this.index));
      const target = this.target();
      if (!visible(target)) { this.close(); return; }
      this.activeTarget = target.id;
      // On narrow screens leave enough room below the highlighted instruction for the guide.
      const topbar = this.doc.querySelector('.topbar');
      const headerBottom = topbar ? Math.max(16, Math.min(topbar.getBoundingClientRect().bottom + 16, 180)) : 24;
      const offset = this.win.innerWidth < 700 ? headerBottom : Math.max(100, (this.win.innerHeight - target.getBoundingClientRect().height) / 2);
      this.win.scrollTo({ top: Math.max(0, this.win.scrollY + target.getBoundingClientRect().top - offset), behavior: 'instant' });
      this.position(); this.schedulePosition();
    }
    schedulePosition() {
      if (!this.open || this.frame !== null) return;
      this.frame = this.win.requestAnimationFrame(() => { this.frame = null; this.position(); });
    }
    position() {
      if (!this.open) return;
      const target = this.target();
      if (!visible(target)) { this.close(); return; }
      if (target.id !== this.activeTarget) { this.show(this.index); return; }
      const bounds = target.getBoundingClientRect(), pad = 10;
      const rect = { left: Math.max(8, bounds.left - pad), top: Math.max(8, bounds.top - pad), right: Math.min(this.win.innerWidth - 8, bounds.right + pad), bottom: Math.min(this.win.innerHeight - 8, bounds.bottom + pad) };
      const spot = this.el('spotlight');
      const inView = bounds.bottom > 0 && bounds.top < this.win.innerHeight;
      spot.hidden = !inView; this.el('arrow').hidden = !inView;
      Object.assign(spot.style, { left: rect.left + 'px', top: rect.top + 'px', width: Math.max(0, rect.right - rect.left) + 'px', height: Math.max(0, rect.bottom - rect.top) + 'px' });
      const size = this.card.getBoundingClientRect(), viewport = { width: this.win.innerWidth, height: this.win.visualViewport?.height || this.win.innerHeight };
      const pos = placement(rect, size, viewport); Object.assign(this.card.style, { left: pos.x + 'px', top: pos.y + 'px' });
      let from, to;
      if (pos.side === 'left') { from = [pos.x + size.width + 5, pos.y + Math.min(60, size.height / 2)]; to = [rect.left - 4, (rect.top + rect.bottom) / 2]; }
      else if (pos.side === 'right') { from = [pos.x - 5, pos.y + Math.min(60, size.height / 2)]; to = [rect.right + 4, (rect.top + rect.bottom) / 2]; }
      else if (pos.side === 'top') { from = [clamp((rect.left + rect.right) / 2, pos.x + 25, pos.x + size.width - 25), pos.y + size.height + 5]; to = [(rect.left + rect.right) / 2, rect.top - 4]; }
      else { from = [clamp((rect.left + rect.right) / 2, pos.x + 25, pos.x + size.width - 25), pos.y - 5]; to = [(rect.left + rect.right) / 2, rect.bottom + 4]; }
      const overlaps = pos.x < rect.right && pos.x + size.width > rect.left && pos.y < rect.bottom && pos.y + size.height > rect.top;
      this.el('arrow').hidden = !inView || overlaps;
      this.el('arrow-path').setAttribute('d', 'M ' + from.join(' ') + ' Q ' + from[0] + ' ' + to[1] + ' ' + to.join(' '));
    }
    close(restoreFocus = true) {
      this.cancelled = true; this.open = false; this.layer.hidden = true; this.doc.body.classList.remove('max-tour-open');
      if (this.pending !== null) this.win.clearTimeout(this.pending); this.pending = null;
      if (this.frame !== null) this.win.cancelAnimationFrame(this.frame); this.frame = null;
      if (restoreFocus && this.previous?.isConnected) { this.previous.scrollIntoView({ block: 'center', behavior: 'instant' }); this.previous.focus({ preventScroll: true }); }
    }
    finish() {
      this.close(false);
      const create = this.doc.getElementById('max-bot-create');
      const destination = visible(create) && !create.disabled ? create : this.doc.getElementById('max-bot-heading');
      destination?.scrollIntoView({ block: 'center', behavior: 'instant' }); destination?.focus({ preventScroll: true });
    }
  }
  return { STEPS, placement, Tour, boot(win) { if (win?.document.getElementById('max-tour-layer')) return new Tour(win).init(); } };
});
