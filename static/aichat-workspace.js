/* Workspace presentation only. Data comes from the existing ACL-filtered
 * dialog/job controllers; this module never fetches, runs or persists tasks. */
(function(root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.NeuronaWorkspace = api;
})(typeof window === 'undefined' ? globalThis : window, function() {
  'use strict';
  const SOURCES = new Set(['/mingkh', '/mingkh/water-map', '/mingkh/treatment', '/mingkh/collective',
    '/edds', '/ecur', '/cds', '/water-dashboard', '/water-rm', '/mgkh-rm', '/zips', '/cameras',
    '/overdue', '/watercontrol', '/utnkr', '/edo', '/omsu-rating', '/summarizer']);
  const LABELS = {queued:'Ожидает выполнения', running:'Выполняется', succeeded:'Завершено',
    failed:'Не удалось завершить', interrupted:'Выполнение прервано', timed_out:'Время истекло',
    partial:'Данные подтверждены частично', done:'Завершено', error:'Ошибка', success:'Завершено',
    ok:'Подтверждено', missing:'Нет данных', pending:'Ожидает подтверждения'};
  const TOOLS = {'query:mingkh-complaints':'Жалобы МинЖКХ', 'query:collective':'Коллективные обращения',
    'query:omsu-rating':'Рейтинг ОМСУ', 'transform:summarizer':'Сокращение текста',
    'query:water-dashboard':'Сводный дашборд', 'analytics:mingkh-map':'Карта МинЖКХ',
    'analytics:edds-water':'ЕДДС — жалобы по воде', 'query:overdue':'Просроченные задачи', 'query:ecur':'Жалобы ЕЦУР'};
  const TOPICS = {all:'Все темы', water:'Водоснабжение', sewerage:'Водоотведение',
    water_quality:'Качество воды', hot_water:'Горячая вода', heating:'Отопление',
    capital_repair:'Капитальный ремонт', garbage:'Вывоз отходов', rusty_water:'Ржавая вода',
    no_water:'Отсутствие воды', pressure:'Давление воды', wells:'Колодцы',
    sewage_discharge:'Сброс сточных вод', boilers:'Котельные', heat_networks:'Тепловые сети',
    network_modernization:'Модернизация сетей', ecology:'Экология', yards:'Дворы', public_areas:'Общественные территории'};
  const SERVICES = {all:'Все услуги', hvs:'ХВС', gvs:'ГВС', vo:'Водоотведение'};
  const list = value => Array.isArray(value) ? value : [];
  const obj = value => value && typeof value === 'object' && !Array.isArray(value) ? value : {};
  const str = value => typeof value === 'string' ? value : '';
  const title = value => str(value).trim() === 'Общий диалог' ? 'Диалог' : str(value).trim();
  function sourcePath(value) {
    // Only canonical local block pages. Never resolve attacker-controlled URLs.
    const path = str(value).trim().replace(/\/$/, '');
    return SOURCES.has(path) ? path : '';
  }
  function sources(value) {
    const used = new Set();
    return list(value).flatMap(item => {
      const path = sourcePath(obj(item).source_url);
      if (!path || used.has(path)) return [];
      used.add(path);
      return [{path, title:str(item.title) || 'Открыть источник'}];
    });
  }
  function period(value) {
    const p = obj(value);
    if (str(p.label)) return p.label;
    const from = str(p.from) || str(p.start), to = str(p.to) || str(p.end);
    return from && to ? from + ' — ' + to : from || to;
  }
  function params(value, result) {
    const p = obj(value), r = obj(result), filters = obj(r.filters);
    const rows = [];
    const add = (label, text) => { if (text) rows.push([label, text]); };
    add('Период', str(r.period_label) || period(r.requested_period) || period(p.requested_period) ||
      period(p) || period(p.period));
    add('Муниципалитет', str(p.municipality) || str(filters.municipality) || str(r.municipality) ||
      list(p.municipalities).filter(item => typeof item === 'string').join(', '));
    add('Тема', str(r.topic_label) || str(filters.topic_label) || str(p.topic_label) || TOPICS[p.topic] || '');
    add('Услуга', str(r.service_label) || str(p.service_label) || SERVICES[p.service] || '');
    return rows;
  }
  function time(value) {
    if (typeof value === 'number') return value < 1e12 ? value * 1000 : value;
    return Date.parse(value) || 0;
  }
  function create(env, options) {
    options = options || {};
    const doc = options.document || env.document;
    const get = id => doc.getElementById(id);
    const nodes = {};
    ['acContext', 'workspaceTaskSummary', 'workspaceContextContent', 'workspaceContextEmpty', 'workspaceSources',
      'contextToggle', 'contextClose', 'taskCount', 'acSidebar', 'historyToggle', 'historyClose', 'workspaceBackdrop',
      'reportOvl', 'reportTitle', 'reportClose', 'reportCopy', 'reportContent', 'workspaceActivity', 'workspaceActivityText', 'conversationTitle']
      .forEach(id => { nodes[id] = get(id); });
    let selected = null, selectedTitle = '', dialog = null, jobs = [], jobState = {}, generation = 0;
    let contextOpen = false, drawer = '', returnFocus = null, userClosed = false, autoOpened = false;
    let report = null, reportVersion = 0, destroyed = false;
    const listeners = [], inertNodes = new Map();
    const historyQuery = env.matchMedia?.('(max-width: 759px)');
    const contextQuery = env.matchMedia?.('(max-width: 1199px)');
    const narrow = name => name === 'history' ? (historyQuery ? historyQuery.matches : env.innerWidth < 760) :
      (contextQuery ? contextQuery.matches : env.innerWidth < 1200);
    function on(node, event, fn) {
      if (!node?.addEventListener) return;
      node.addEventListener(event, fn);listeners.push(() => node.removeEventListener(event, fn));
    }
    function el(tag, className, text) {
      const node = doc.createElement(tag);
      if (className) node.className = className;
      if (text != null) node.textContent = text;
      return node;
    }
    function modalOpen() { return Array.from(doc.querySelectorAll('.ovl')).some(node => !node.hidden); }
    function restoreInert() {
      inertNodes.forEach((value, node) => { node.inert = value; });inertNodes.clear();
    }
    function syncInert() {
      restoreInert();
      if (!drawer || modalOpen()) return;
      let branch = drawer === 'history' ? nodes.acSidebar : nodes.acContext;
      while (branch?.parentElement) {
        Array.from(branch.parentElement.children).forEach(node => {
          if (node === branch || node === nodes.workspaceBackdrop || node.classList.contains('ovl')) return;
          inertNodes.set(node, !!node.inert);node.inert = true;
        });
        branch = branch.parentElement;
        if (branch === doc.body) break;
      }
    }
    function visible(node) {
      if (!node || node.disabled || node.closest('[hidden], [inert]') || node.getAttribute('aria-hidden') === 'true') return false;
      const style = env.getComputedStyle?.(node);
      if (style && (style.display === 'none' || style.visibility === 'hidden' || style.visibility === 'collapse')) return false;
      return !node.getClientRects || node.getClientRects().length > 0;
    }
    function focusables(panel) {
      return Array.from(panel?.querySelectorAll('button, a[href], input, textarea, select, [tabindex]') || [])
        .filter(node => visible(node) && node.getAttribute('tabindex') !== '-1');
    }
    function focusFirst(name) {
      const panel = name === 'history' ? nodes.acSidebar : nodes.acContext;
      const preferred = name === 'history' ? get('historySearch') : nodes.contextClose;
      const target = (visible(preferred) ? preferred : null) || focusables(panel)[0] || panel;
      if (target) { if (target === panel) target.setAttribute('tabindex', '-1');target.focus({preventScroll:true}); }
    }
    function syncLayout() {
      const showContext = drawer === 'context' || (!narrow('context') && contextOpen);
      if (nodes.acContext) nodes.acContext.hidden = !showContext;
      if (nodes.acSidebar) nodes.acSidebar.hidden = narrow('history') && drawer !== 'history';
      doc.body.classList.toggle('ac-context-visible', showContext && !narrow('context'));
      if (drawer) doc.body.setAttribute('data-workspace-panel', drawer);
      else doc.body.removeAttribute('data-workspace-panel');
      if (nodes.workspaceBackdrop) nodes.workspaceBackdrop.hidden = !drawer;
      nodes.contextToggle?.setAttribute('aria-expanded', String(showContext));
      nodes.historyToggle?.setAttribute('aria-expanded', String(drawer === 'history'));
      [['history',nodes.acSidebar],['context',nodes.acContext]].forEach(([name, panel]) => {
        if (!panel) return;
        if (drawer === name) { panel.setAttribute('role','dialog');panel.setAttribute('aria-modal','true'); }
        else { panel.removeAttribute('role');panel.removeAttribute('aria-modal'); }
      });
      syncInert();
    }
    function closePanel(name, restore) {
      const closing = name || drawer || 'context';
      if (closing === 'context') { contextOpen = false;userClosed = true; }
      if (drawer === closing) drawer = '';
      syncLayout();
      if (restore !== false && !modalOpen()) {
        const fallback = closing === 'history' ? nodes.historyToggle : nodes.contextToggle;
        const target = returnFocus?.isConnected && visible(returnFocus) ? returnFocus : fallback;
        if (visible(target)) target.focus({preventScroll:true});
      }
      returnFocus = null;
    }
    function openPanel(name, automatic) {
      if (destroyed || !['history','context'].includes(name)) return;
      if (automatic && (narrow(name) || userClosed || modalOpen())) return;
      if (!automatic) returnFocus = doc.activeElement;
      if (drawer && drawer !== name) drawer = '';
      if (name === 'context') contextOpen = true;
      if (narrow(name)) drawer = name;
      syncLayout();
      if (!automatic) focusFirst(name);
    }
    function toggle(name) {
      if (drawer === name || (name === 'context' && contextOpen && !narrow(name))) closePanel(name);
      else openPanel(name);
    }
    function resized() {
      const previousFocus = doc.activeElement;
      if (drawer && !narrow(drawer)) { drawer = '';returnFocus = null; }
      // A docked panel never turns into an unsolicited modal after resize.
      if (narrow('context') && drawer !== 'context') contextOpen = false;
      syncLayout();
      if (!visible(previousFocus) && !modalOpen()) {
        const target = nodes.acSidebar?.contains(previousFocus) ? nodes.historyToggle :
          nodes.acContext?.contains(previousFocus) ? nodes.contextToggle : null;
        if (visible(target)) target.focus({preventScroll:true});
      }
    }
    function viewportChanged() {
      const height = env.visualViewport?.height || env.innerHeight;
      if (Number.isFinite(height)) doc.body.style?.setProperty('--ac-viewport-height',Math.round(height) + 'px');
      doc.body.classList.toggle('ac-keyboard-open', narrow('history') && env.innerHeight - height > 140);
    }
    function keydown(event) {
      if (modalOpen()) { syncInert();return; }
      const name = drawer || (contextOpen && nodes.acContext?.contains(doc.activeElement) ? 'context' : '');
      if (!name) return;
      if (event.key === 'Escape') { event.preventDefault();closePanel(name);return; }
      if (event.key !== 'Tab' || !drawer) return;
      const panel = drawer === 'history' ? nodes.acSidebar : nodes.acContext;
      const items = focusables(panel), first = items[0], last = items[items.length - 1];
      if (!first) { event.preventDefault();panel?.focus({preventScroll:true});return; }
      if ((event.shiftKey && (doc.activeElement === first || !panel.contains(doc.activeElement))) ||
          (!event.shiftKey && (doc.activeElement === last || !panel.contains(doc.activeElement)))) {
        event.preventDefault();(event.shiftKey ? last : first).focus({preventScroll:true});
      }
    }
    function clearReport() {
      report = null;reportVersion++;
      if (nodes.reportOvl && !nodes.reportOvl.hidden) {
        if (options.closeModal) options.closeModal('#reportOvl');else nodes.reportOvl.hidden = true;
      }
      nodes.reportContent?.replaceChildren();
      if (nodes.reportTitle) nodes.reportTitle.textContent = 'Отчёт';
      if (nodes.reportCopy) { nodes.reportCopy.textContent = 'Скопировать текст';nodes.reportCopy.disabled = false; }
      syncInert();
    }
    function sameMessage(a, b) {
      return !b.restricted && a.role === b.role && a.content === b.content &&
        (!a.event_id || a.event_id === b.event_id);
    }
    function authorized(message) {
      return !message.restricted && (!dialog || list(dialog.messages).some(item => sameMessage(message, item)));
    }
    function openReport(message, version) {
      if (destroyed || version !== generation || !authorized(message) || !nodes.reportContent || !nodes.reportOvl) return;
      if (drawer) closePanel(drawer, false);
      report = message;reportVersion++;
      if (nodes.reportTitle) nodes.reportTitle.textContent = selectedTitle ? 'Отчёт · ' + selectedTitle : 'Отчёт';
      if (nodes.reportCopy) nodes.reportCopy.textContent = 'Скопировать текст';
      const body = el('div','bubble');
      if (typeof options.format === 'function') body.innerHTML = options.format(message.content);
      else body.textContent = message.content;
      nodes.reportContent.replaceChildren(body);
      if (options.openModal) options.openModal('#reportOvl', '#reportClose');
      else { nodes.reportOvl.hidden = false;nodes.reportClose?.focus(); }
      syncInert();
    }
    async function copyReport() {
      if (!report || !authorized(report)) { clearReport();return; }
      const version = reportVersion;
      const text = nodes.reportContent.innerText || nodes.reportContent.textContent || report.content;
      try {
        if (!env.navigator?.clipboard?.writeText) throw new Error('Clipboard unavailable');
        await env.navigator.clipboard.writeText(text);
        if (version === reportVersion && nodes.reportCopy) nodes.reportCopy.textContent = 'Текст скопирован';
      } catch (_) {
        if (version === reportVersion && nodes.reportCopy) nodes.reportCopy.textContent = 'Не удалось скопировать. Повторить';
      }
    }
    function decorateMessage(element, message) {
      if (destroyed || !element || !message || message.role !== 'assistant' || message.restricted || !str(message.content).trim()) return;
      const content = str(message.content);
      const table = content.split(/\r?\n/).some(line => /^\s*\|?\s*:?-{3,}:?\s*\|\s*:?-{3,}:?/.test(line));
      const reportData = sources(message.report_sources).length || Object.keys(obj(message.artifact)).length ||
        Object.keys(obj(message.report_scope)).length || list(obj(message.agent_plan).steps).length;
      if (!table && content.length < 800 && !reportData) return;
      const body = element.querySelector('.ac-message-body') || element;
      if (body.querySelector('.ac-report-open')) return;
      const actions = el('div','ac-report-actions'), button = el('button','ac-report-open','Открыть отчёт');
      const version = generation;
      button.type = 'button';button.setAttribute('aria-haspopup','dialog');button.setAttribute('aria-controls','reportOvl');
      button.addEventListener('click', () => openReport(message, version));
      actions.appendChild(button);body.appendChild(actions);
    }
    function hasMetadata(value) {
      const v = obj(value);
      return !v.restricted && (list(obj(v.agent_plan).steps).length || list(obj(obj(v.action).plan).steps).length ||
        Object.keys(obj(v.report_scope)).length || sources(v.report_sources).length || Object.keys(obj(v.agent_pending)).length);
    }
    function currentTask() {
      let candidate = null;
      list(dialog?.messages).forEach((message, index) => {
        if (message.role === 'assistant' && hasMetadata(message)) {
          candidate = {value:message, stamp:time(message.ts), index, job:null};
        }
      });
      const ordered = jobs.filter(job => job.dialog_id === selected && !job.restricted)
        .slice().sort((a,b) => time(b.created_at) - time(a.created_at));
      const latest = ordered[0];
      if (latest && (!candidate || latest.job_id === candidate.value.event_id || time(latest.created_at) >= candidate.stamp)) {
        candidate = {value:obj(latest.result), job:latest, stamp:time(latest.created_at)};
      }
      return candidate;
    }
    function appendParams(host, rows) {
      if (!rows.length) return;
      const dl = el('dl','ac-context-params');
      rows.forEach(([label,value]) => { dl.appendChild(el('dt','ac-context-label',label));dl.appendChild(el('dd','ac-context-value',value)); });
      host.appendChild(dl);
    }
    function render() {
      const focusedSource = nodes.workspaceSources?.contains(doc.activeElement) ? doc.activeElement.getAttribute('href') : null;
      const task = currentTask(), value = obj(task?.value), job = task?.job;
      const plan = Object.keys(obj(value.agent_plan)).length ? obj(value.agent_plan) : obj(obj(value.action).plan);
      const pending = obj(value.agent_pending), scope = obj(value.report_scope), action = obj(value.action);
      const steps = list(plan.steps), progress = list(job?.steps), hasTask = !!task;
      const active = jobs.filter(item => ['queued','running'].includes(item.status));
      if (nodes.taskCount) { nodes.taskCount.textContent = String(active.length);nodes.taskCount.hidden = !active.length; }
      if (nodes.workspaceContextEmpty) nodes.workspaceContextEmpty.hidden = hasTask;
      if (nodes.workspaceContextContent) nodes.workspaceContextContent.hidden = !hasTask;
      if (nodes.workspaceTaskSummary) {
        nodes.workspaceTaskSummary.hidden = !hasTask;
        nodes.workspaceTaskSummary.replaceChildren();
        if (hasTask) {
          nodes.workspaceTaskSummary.appendChild(el('strong','ac-context-value',
            str(action.label) || (steps.length ? 'План действий' : hasMetadata(value) ? 'Контекст отчёта' : 'Задача чата')));
          const status = action.kind && !action.confirmed_job_id ? 'Ожидает подтверждения' : LABELS[job?.status || plan.status];
          if (status) nodes.workspaceTaskSummary.appendChild(el('p','ac-context-status',status));
          if (str(job?.stage)) nodes.workspaceTaskSummary.appendChild(el('p','ac-context-stage',job.stage));
        }
      }
      if (nodes.workspaceContextContent) {
        nodes.workspaceContextContent.replaceChildren();
        appendParams(nodes.workspaceContextContent, params(scope));
        if (Object.keys(pending).length) {
          nodes.workspaceContextContent.appendChild(el('p','ac-context-stage','Уточнение параметров'));
          appendParams(nodes.workspaceContextContent, params(pending));
          const periods = list(pending.periods).map(period).filter(Boolean);
          if (periods.length) appendParams(nodes.workspaceContextContent, [['Период',periods.join('; ')]]);
        }
        if (steps.length) {
          const ol = el('ol','ac-plan-steps');
          steps.forEach((step,index) => {
            const item = obj(step), result = obj(item.result), source = obj(result.source);
            const row = el('li','ac-plan-step');
            // Do not label unstarted steps successful when their parent job ended.
            const status = str(item.status) || (action.kind ? 'pending' : '');
            row.setAttribute('data-status', LABELS[status] ? status : 'unknown');
            const sourceTitle = sources(value.report_sources)[index]?.title;
            row.appendChild(el('strong','ac-plan-step-title',str(item.label) || str(source.title) || TOOLS[item.tool] || sourceTitle || 'Шаг ' + (index + 1)));
            if (LABELS[status]) row.appendChild(el('span','ac-plan-step-status',LABELS[status]));
            appendParams(row, params(item.params,result));
            ol.appendChild(row);
          });
          nodes.workspaceContextContent.appendChild(ol);
        } else if (progress.length) {
          const ol = el('ol','ac-plan-steps');
          progress.forEach(step => {
            if (!str(step.label)) return;
            const row = el('li','ac-plan-step');row.setAttribute('data-status', LABELS[step.status] ? step.status : 'unknown');
            row.appendChild(el('strong','ac-plan-step-title',step.label));
            if (LABELS[step.status]) row.appendChild(el('span','ac-plan-step-status',LABELS[step.status]));
            ol.appendChild(row);
          });
          nodes.workspaceContextContent.appendChild(ol);
        }
      }
      if (nodes.workspaceSources) {
        nodes.workspaceSources.replaceChildren();
        const links = sources(value.report_sources);
        if (links.length) nodes.workspaceSources.appendChild(el('h3','ac-context-label','Источники'));
        links.forEach(source => {
          const link = el('a','ac-context-source',source.title);link.setAttribute('href',source.path);nodes.workspaceSources.appendChild(link);
        });
        nodes.workspaceSources.hidden = !links.length;
        if (focusedSource && !modalOpen()) {
          const target = Array.from(nodes.workspaceSources.querySelectorAll('a')).find(link => link.getAttribute('href') === focusedSource);
          if (visible(target)) target.focus({preventScroll:true});
        }
      }
      const awaiting = jobs.some(item => obj(item.result).action && !obj(obj(item.result).action).confirmed_job_id);
      const pendingItems = list(jobState.pending), failed = job && ['failed','interrupted','timed_out'].includes(job.status);
      const showActivity = active.length || awaiting || pendingItems.length || jobState.connection || failed;
      if (nodes.workspaceActivity) nodes.workspaceActivity.hidden = !showActivity;
      if (nodes.workspaceActivityText) nodes.workspaceActivityText.textContent =
        jobState.connection ? 'Связь прервана — открыть задачи' : pendingItems.some(item => item.error) ? 'Отправка требует внимания' :
        awaiting ? 'План ждёт подтверждения' : active.length ? (str(active[0].stage) || 'Задача выполняется') + ' — открыть задачи' :
        pendingItems.length ? 'Отправляем запрос…' : failed ? LABELS[job.status] + ' — открыть задачи' : '';
      if (hasMetadata(value) && !autoOpened) {
        autoOpened = true;openPanel('context',true);
      }
    }
    function select(id, name) {
      generation++;selected = id || null;selectedTitle = title(name);dialog = null;jobs = [];jobState = {};
      if (nodes.conversationTitle) nodes.conversationTitle.textContent = selectedTitle || (selected ? 'Диалог' : 'Новый диалог');
      clearReport();contextOpen = false;drawer = '';returnFocus = null;userClosed = false;autoOpened = false;
      syncLayout();render();
    }
    function setDialog(value) {
      if (!value) { select(null);return; }
      if (value.id && value.id !== selected) return;
      dialog = value;selectedTitle = title(value.title) || selectedTitle;
      if (nodes.conversationTitle) nodes.conversationTitle.textContent = selectedTitle || 'Диалог';
      if (report && !authorized(report)) clearReport();
      // A filtered message is authoritative even if an earlier job poll included its result.
      const restricted = new Set(list(value.messages).filter(item => item.restricted).map(item => item.event_id).filter(Boolean));
      jobs = jobs.filter(job => !restricted.has(job.job_id));
      render();
      if (!currentTask() && contextOpen) { contextOpen = false;drawer = drawer === 'context' ? '' : drawer;syncLayout(); }
    }
    function setJobs(state) {
      jobState = obj(state);
      const restricted = new Set(list(dialog?.messages).filter(item => item.restricted).map(item => item.event_id).filter(Boolean));
      jobs = list(jobState.jobs).filter(job => job && !job.restricted && !restricted.has(job.job_id));
      render();
    }
    on(nodes.contextToggle,'click',() => toggle('context'));
    on(nodes.contextClose,'click',() => closePanel('context'));
    on(nodes.historyToggle,'click',() => toggle('history'));
    on(nodes.historyClose,'click',() => closePanel('history'));
    on(nodes.workspaceActivity,'click',() => openPanel('context'));
    on(nodes.workspaceBackdrop,'click',() => closePanel());
    on(nodes.reportClose,'click',clearReport);
    on(nodes.reportCopy,'click',copyReport);
    on(doc,'keydown',keydown);
    on(doc,'focusin',() => {
      if (!drawer || modalOpen()) return;
      const panel = drawer === 'history' ? nodes.acSidebar : nodes.acContext;
      if (!panel?.contains(doc.activeElement)) focusFirst(drawer);
    });
    on(env,'resize',resized);
    on(env,'resize',viewportChanged);
    on(env.visualViewport,'resize',viewportChanged);
    [historyQuery,contextQuery].forEach(query => on(query,'change',resized));
    const observer = env.MutationObserver ? new env.MutationObserver(syncInert) : null;
    observer?.observe(doc.body,{subtree:true,attributes:true,attributeFilter:['hidden']});
    syncLayout();viewportChanged();render();
    return {select, setDialog, setJobs, decorateMessage, openPanel, closePanel, reset:() => select(null),
      destroy:() => { destroyed = true;clearReport();drawer = '';contextOpen = false;syncLayout();observer?.disconnect();listeners.forEach(remove => remove()); }};
  }
  return {create, sourcePath, sources};
});
