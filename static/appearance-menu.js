/* One appearance controller for the home page and all modules. */
(() => {
  'use strict';
  const root = document.documentElement;
  const modes = new Set(['light', 'dark', 'kreml', 'auto-moscow']);
  const labels = {light:'Светлая', dark:'Тёмная', kreml:'Московская', 'auto-moscow':'Подмосковное солнце'};
  const storage = {
    get(key, fallback) { try { return localStorage.getItem(key) || fallback; } catch (_) { return fallback; } },
    set(key, value) { try { localStorage.setItem(key, value); } catch (_) {} }
  };
  let mode = storage.get('neurona-theme-mode', 'auto-moscow');
  if (!modes.has(mode)) mode = 'auto-moscow';
  let glass = storage.get('neurona-glass', 'off') === 'on';
  let dropdown, trigger, menu;

  function isDaytimeInMoscow(now = new Date()) {
    // The setting uses Moscow time even if the workstation is in another timezone.
    const moscow = new Date(now.getTime() + 3 * 3600000);
    const day = Math.floor((Date.UTC(moscow.getUTCFullYear(), moscow.getUTCMonth(), moscow.getUTCDate()) - Date.UTC(moscow.getUTCFullYear(), 0, 0)) / 86400000);
    const decl = .409 * Math.sin(2 * Math.PI * (day - 80) / 365), lat = 55.7558 * Math.PI / 180;
    const cosH = (Math.sin(-.83 * Math.PI / 180) - Math.sin(lat) * Math.sin(decl)) / (Math.cos(lat) * Math.cos(decl));
    const h = cosH >= 1 ? 0 : cosH <= -1 ? Math.PI : Math.acos(cosH);
    const sunrise = (12 - h * 12 / Math.PI - 37.6173 / 15 + 3 + 24) % 24;
    const sunset = (12 + h * 12 / Math.PI - 37.6173 / 15 + 3 + 24) % 24;
    const hour = moscow.getUTCHours() + moscow.getUTCMinutes() / 60;
    return hour >= sunrise && hour < sunset;
  }

  function apply() {
    const theme = mode === 'auto-moscow' ? (isDaytimeInMoscow() ? 'light' : 'dark') : mode;
    if (root.dataset.theme !== theme) root.dataset.theme = theme;
    if (root.dataset.glass !== (glass ? 'on' : 'off')) root.dataset.glass = glass ? 'on' : 'off';
    if (!dropdown) return;
    dropdown.querySelector('.theme-toggle-text').textContent = labels[mode];
    menu.querySelectorAll('[data-theme-mode]').forEach(button => {
      const active = button.dataset.themeMode === mode;
      button.classList.toggle('active', active);
      button.setAttribute('aria-checked', String(active));
      button.tabIndex = active ? 0 : -1;
    });
    const switcher = menu.querySelector('#glassMenuToggle');
    switcher.classList.toggle('active', glass);
    switcher.setAttribute('aria-checked', String(glass));
  }
  apply();

  const markup = '<div class="theme-dropdown appearance-dropdown neurona-appearance">\n                    <button id="themeToggle" class="theme-toggle appearance-trigger" type="button"\n                            aria-haspopup="true" aria-expanded="false" aria-label="Настроить оформление">\n                        <span class="appearance-trigger-icon" aria-hidden="true">◐</span>\n                        <span class="theme-toggle-copy">\n                            <small>Оформление</small>\n                            <span class="theme-toggle-text">Подмосковное солнце</span>\n                        </span>\n                        <span class="theme-toggle-arrow" aria-hidden="true">⌄</span>\n                    </button>\n                    <div id="themeMenu" class="theme-menu appearance-menu" role="dialog" aria-label="Настройки оформления" hidden>\n                        <div class="appearance-menu-head">\n                            <div><strong>Оформление</strong><span>Настройте вид рабочего пространства</span></div>\n                        </div>\n                        <div class="appearance-section-label">Цветовая тема</div>\n                        <div class="appearance-theme-grid" role="radiogroup" aria-label="Цветовая тема">\n                            <button class="theme-menu-item" type="button" role="radio" data-theme-mode="light">\n                                <span class="theme-menu-icon">☀️</span><span class="theme-menu-label">Светлая</span><span class="theme-menu-check">✓</span>\n                            </button>\n                            <button class="theme-menu-item" type="button" role="radio" data-theme-mode="dark">\n                                <span class="theme-menu-icon">🌙</span><span class="theme-menu-label">Тёмная</span><span class="theme-menu-check">✓</span>\n                            </button>\n                            <button class="theme-menu-item" type="button" role="radio" data-theme-mode="kreml">\n                                <span class="theme-menu-icon">🏛️</span><span class="theme-menu-label">Московская</span><span class="theme-menu-check">✓</span>\n                            </button>\n                            <button class="theme-menu-item theme-menu-auto" type="button" role="radio" data-theme-mode="auto-moscow">\n                                <span class="theme-menu-icon">🛰️</span><span class="theme-menu-label"><strong>Подмосковное солнце</strong><small>Автоматически по времени суток</small></span><span class="theme-menu-check">✓</span>\n                            </button>\n                        </div>\n                        <div class="appearance-divider"></div>\n                        <button id="glassMenuToggle" class="appearance-setting" type="button" role="switch" aria-checked="false">\n                            <span class="appearance-setting-icon" aria-hidden="true">◇</span>\n                            <span class="appearance-setting-copy"><strong>Жидкое стекло</strong><small>Прозрачность и мягкое размытие карточек</small></span>\n                            <span class="appearance-switch" aria-hidden="true"><span></span></span>\n                        </button>\n                    </div>\n                </div>';
  function floatingPanel(panel, anchor, preferredWidth = 368) {
    const native = typeof panel.showPopover === 'function';
    panel.setAttribute('data-floating-panel', '');
    if (native) panel.setAttribute('popover', 'manual');
    let portal, originalParent, originalNext;
    function position() {
      if (panel.hidden) return;
      const viewport = window.visualViewport;
      const width = viewport?.width || root.clientWidth;
      const height = viewport?.height || window.innerHeight;
      const left = viewport?.offsetLeft || 0, top = viewport?.offsetTop || 0;
      const margin = 12;
      const targetWidth = Math.min(width - margin * 2, preferredWidth);
      const rect = anchor.getBoundingClientRect();
      panel.style.setProperty('--menu-width', `${targetWidth}px`);
      const naturalHeight = panel.scrollHeight;
      const below = top + height - rect.bottom - margin - 8;
      const above = rect.top - top - margin - 8;
      const room = Math.max(below, above);
      const availableHeight = Math.max(1, height - margin * 2);
      const panelHeight = Math.min(naturalHeight, availableHeight, room >= 220 ? room : availableHeight);
      let y = below >= panelHeight ? rect.bottom + 8 : above >= panelHeight ? rect.top - panelHeight - 8 : top + margin;
      y = Math.max(top + margin, Math.min(y, top + height - panelHeight - margin));
      panel.style.setProperty('--menu-x', `${Math.max(left + margin, Math.min(rect.right - targetWidth, left + width - targetWidth - margin))}px`);
      panel.style.setProperty('--menu-y', `${y}px`);
      panel.style.setProperty('--menu-height', `${Math.min(availableHeight, Math.max(80, panelHeight))}px`);
    }
    function watch(on) {
      const method = on ? 'addEventListener' : 'removeEventListener';
      window[method]('resize', position);
      window[method]('scroll', position, true);
      window.visualViewport?.[method]('resize', position);
      window.visualViewport?.[method]('scroll', position);
    }
    return {
      open() {
        // Older browsers have no top layer. Temporarily move the panel outside
        // blurred headers so fixed positioning remains relative to the viewport.
        if (!native && !portal) {
          originalParent = panel.parentNode;
          originalNext = panel.nextSibling;
          portal = document.createElement('div');
          portal.className = 'neurona-appearance floating-panel-portal';
          document.body.appendChild(portal);
          portal.appendChild(panel);
        }
        panel.hidden = false;
        if (native && !panel.matches(':popover-open')) panel.showPopover();
        position(); watch(true);
      },
      close() {
        if (native && panel.matches(':popover-open')) panel.hidePopover();
        panel.hidden = true; watch(false);
        if (portal) {
          originalParent.insertBefore(panel, originalNext);
          portal.remove();
          portal = null;
        }
      }
    };
  }
  let floating;
  function close(restoreFocus = false) {
    if (!menu) return;
    floating.close();
    dropdown.classList.remove('open');
    trigger.setAttribute('aria-expanded', 'false');
    if (restoreFocus) trigger.focus();
  }
  function open(focus = false) {
    floating.open();
    dropdown.classList.add('open');
    trigger.setAttribute('aria-expanded', 'true');
    if (focus) menu.querySelector('[aria-checked="true"][role="radio"]')?.focus();
  }
  function mount(host) {
    if (!host || dropdown) return;
    // Remove only the old appearance controls; preserve module actions and navigation.
    document.querySelectorAll('#glassToggle').forEach(element => element.remove());
    const old = host.querySelector('.theme-dropdown') || host.querySelector('#themeToggle, #themeToggleBtn, .theme-toggle-report');
    const template = document.createElement('template');
    template.innerHTML = markup;
    dropdown = template.content.firstElementChild;
    if (old) old.replaceWith(dropdown); else host.prepend(dropdown);
    trigger = dropdown.querySelector('#themeToggle');
    menu = dropdown.querySelector('#themeMenu');
    floating = floatingPanel(menu, trigger);
    trigger.setAttribute('aria-controls', 'themeMenu');
    trigger.addEventListener('click', event => menu.hidden ? open(event.detail === 0) : close());
    trigger.addEventListener('keydown', event => {
      if (event.key === 'ArrowDown') { event.preventDefault(); open(true); }
    });
    dropdown.querySelectorAll('[data-theme-mode]').forEach(button => {
      button.addEventListener('click', () => {
        mode = button.dataset.themeMode;
        storage.set('neurona-theme-mode', mode);
        apply(); close(true);
      });
      button.addEventListener('keydown', event => {
        const choices = [...menu.querySelectorAll('[data-theme-mode]')];
        const offset = ['ArrowRight','ArrowDown'].includes(event.key) ? 1 : ['ArrowLeft','ArrowUp'].includes(event.key) ? -1 : 0;
        if (!offset) return;
        event.preventDefault();
        choices[(choices.indexOf(button) + offset + choices.length) % choices.length].focus();
      });
    });
    dropdown.querySelector('#glassMenuToggle').addEventListener('click', () => {
      glass = !glass; storage.set('neurona-glass', glass ? 'on' : 'off'); apply();
    });
    const contains = target => dropdown.contains(target) || menu.contains(target);
    document.addEventListener('pointerdown', event => { if (!contains(event.target)) close(); });
    document.addEventListener('keydown', event => { if (event.key === 'Escape' && !menu.hidden) close(true); });
    const onFocusOut = event => { if (!contains(event.relatedTarget)) close(); };
    dropdown.addEventListener('focusout', onFocusOut);
    menu.addEventListener('focusout', onFocusOut);
    apply();
  }
  window.NeuronaAppearance = {mount, floatingPanel};
  function init() {
    const topbar = document.querySelector('.unified-shell-topbar, header.topbar, .topbar, .top');
    if (topbar) mount(topbar.querySelector('.top-actions, .topbar-right, .system-shell-actions, .actions, .nav'));
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, {once:true}); else init();
  window.addEventListener('storage', event => {
    if (event.key === 'neurona-theme-mode') mode = modes.has(event.newValue) ? event.newValue : 'auto-moscow';
    else if (event.key === 'neurona-glass') glass = event.newValue === 'on';
    else return;
    apply();
  });
  setInterval(() => { if (mode === 'auto-moscow') apply(); }, 60000);
})();
