/* Keep the four work areas mounted so switching sections never discards form input. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else if (root.document.getElementById('adminConsole')) api.createController(root).start();
})(typeof window === 'undefined' ? globalThis : window, function () {
  'use strict';
  const sections = {
    users: {title:'Пользователи и права', description:'Управляйте доступом к сервису и смотрите активность каждого пользователя.'},
    statistics: {title:'Статистика', description:'Какие блоки востребованы, а какие пока остаются без посещений.'},
    feedback: {title:'Обратная связь', description:'Результаты опроса об удобстве и оформлении сервиса, оценки и предложения пользователей.'},
    integrations: {title:'Логины и пароли', description:'Настройки доступа к внешним системам и источникам данных.'}
  };
  function createController(env) {
    const doc = env.document;
    const shell = doc.getElementById('adminConsole');
    const panels = [...shell.querySelectorAll('[data-admin-panel]')];
    const links = [...shell.querySelectorAll('[data-admin-link]')];
    const fallback = Object.hasOwn(sections, shell.dataset.initialSection) ? shell.dataset.initialSection : 'users';
    let started = false;
    function requested() {
      const hash = env.location.hash.slice(1);
      return Object.hasOwn(sections, hash) ? hash : fallback;
    }
    function show(section, focus = false) {
      if (!Object.hasOwn(sections, section)) section = fallback;
      for (const panel of panels) panel.hidden = panel.dataset.adminPanel !== section;
      for (const link of links) {
        if (link.dataset.adminLink === section) link.setAttribute('aria-current','page');
        else link.removeAttribute('aria-current');
      }
      doc.body.dataset.adminSection = section;
      const title = doc.getElementById('adminTitle');
      title.textContent = sections[section].title;
      doc.getElementById('adminDescription').textContent = sections[section].description;
      doc.title = `${sections[section].title} · Админ-панель — Нейрона`;
      doc.dispatchEvent(new env.CustomEvent('admin:sectionchange', {detail:{section}}));
      if (focus) {
        title.focus({preventScroll:true});
        doc.getElementById('adminMain').scrollIntoView({block:'start',behavior:'instant'});
      }
    }
    function navigate(event) {
      // Preserve normal open-in-new-tab behaviour and browser link menus.
      if (event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
      const section = event.currentTarget.dataset.adminLink;
      event.preventDefault();
      if (doc.body.dataset.adminSection === section) return;
      env.history.pushState(null, '', '#' + section);
      show(section, true);
    }
    function restore() {
      const hash = env.location.hash.slice(1);
      // In-page accessibility anchors must not select another work area.
      if (hash && !Object.hasOwn(sections, hash) && doc.getElementById(hash)) return;
      show(requested(), true);
    }
    function unload(event) {
      if (['userForm','integrationForm'].some(id => doc.getElementById(id)?.dataset.dirty === 'true')) {
        event.preventDefault();
        event.returnValue = '';
      }
    }
    function start() {
      if (started) return;
      started = true;
      show(requested());
      links.forEach(link => link.addEventListener('click', navigate));
      env.addEventListener('popstate', restore);
      env.addEventListener('hashchange', restore);
      env.addEventListener('beforeunload', unload);
    }
    return {start, show};
  }
  return {createController, sections};
});
