(function (root) {
  'use strict';

  var STEP_KINDS = ['routing', 'skill', 'tool', 'model'];
  var STEP_STATUSES = ['running', 'done', 'error'];
  var STATUS_LABELS = {running: 'В работе', done: 'Готово', error: 'Не выполнено', unknown: 'Статус не получен'};
  var KIND_LABELS = {routing: 'Выбор способа', skill: 'Навык', tool: 'Инструмент', model: 'Ответ'};
  var MAX_RECORD = 2 * 1024 * 1024;

  function text(value, limit) { return typeof value === 'string' ? value.trim().slice(0, limit || 300) : ''; }
  function russianTitle(value, fallback) { var title = text(value, 160); return /[А-Яа-яЁё]/.test(title) ? title : fallback; }
  function copy(value) { return JSON.parse(JSON.stringify(value)); }
  function escapeHtml(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, function (char) {
      return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[char];
    });
  }
  function toolTitle(name) {
    var titles = {calculate: 'Вычисление', table_profile: 'Анализ таблицы', attachment_text: 'Чтение вложения',
      read_resource: 'Справка навыка', platform_report: 'Данные платформы'};
    return Object.prototype.hasOwnProperty.call(titles, name) ? titles[name] : 'Инструмент навыка';
  }
  function cleanStep(event) {
    if (!event || typeof event !== 'object' || Array.isArray(event) ||
        !text(event.id, 120) || !text(event.label, 400) ||
        STEP_KINDS.indexOf(event.kind) === -1 || STEP_STATUSES.indexOf(event.status) === -1) return null;
    var step = {id: text(event.id, 120), kind: event.kind, label: text(event.label, 400), status: event.status};
    if (text(event.skill_id, 120)) step.skill_id = text(event.skill_id, 120);
    return step;
  }
  function uniqueSteps(events) {
    var steps = [];
    (Array.isArray(events) ? events.slice(-120) : []).forEach(function (event) {
      var step = cleanStep(event);
      if (!step) return;
      var index = steps.findIndex(function (item) { return item.id === step.id; });
      if (index === -1) steps.push(step); else steps[index] = step;
    });
    return steps;
  }
  function cleanRun(run) {
    run = run && typeof run === 'object' && !Array.isArray(run) ? run : {};
    return {
      skills: (Array.isArray(run.skills) ? run.skills : []).slice(0, 30).filter(function (item) {
        return item && text(item.id, 120);
      }).map(function (item) { return {id: text(item.id, 120), title: russianTitle(item.title, 'Навык')}; }),
      tools: (Array.isArray(run.tools) ? run.tools : []).slice(0, 60).filter(function (item) {
        return item && text(item.name, 120) && STEP_STATUSES.indexOf(item.status) !== -1;
      }).map(function (item) { return {name: text(item.name, 120), title: russianTitle(item.title, toolTitle(text(item.name, 120))), status: item.status}; }),
      steps: uniqueSteps(run.steps),
      warnings: (Array.isArray(run.warnings) ? run.warnings : []).slice(0, 10).map(function (warning) {
        return text(warning, 600);
      }).filter(Boolean)
    };
  }
  function createTrace() {
    var steps = [];
    return {
      update: function (event) {
        var step = cleanStep(event);
        if (!step) return;
        var index = steps.findIndex(function (item) { return item.id === step.id; });
        if (index === -1) { if (steps.length < 120) steps.push(step); } else steps[index] = step;
      },
      snapshot: function () { return {skills: [], tools: [], steps: copy(steps), warnings: []}; },
      finish: function (resultRun, failed) {
        var run = cleanRun(resultRun);
        if (!run.steps.length) run.steps = copy(steps);
        if (failed) {
          run.steps.forEach(function (step) { if (step.status === 'running') step.status = 'error'; });
          run.tools.forEach(function (tool) { if (tool.status === 'running') tool.status = 'error'; });
        }
        return run;
      }
    };
  }
  function stepHTML(step, final) {
    var status = final && step.status === 'running' ? 'unknown' : step.status;
    return '<li class="ac-skill-step is-' + status + '"><span class="ac-skill-step-mark" aria-hidden="true">' +
      (status === 'done' ? '✓' : status === 'error' ? '!' : '') + '</span><span class="ac-skill-step-content">' +
      '<span class="ac-skill-step-kind">' + KIND_LABELS[step.kind] + '</span><span>' + escapeHtml(step.label) +
      '</span></span><span class="ac-skill-step-status">' + STATUS_LABELS[status] + '</span></li>';
  }
  function activityHTML(run) {
    var steps = cleanRun(run).steps;
    var active = steps.find(function (step) { return step.status === 'running'; });
    return '<div class="ac-skill-activity" role="status" aria-live="polite" aria-atomic="false">' +
      '<div class="ac-skill-activity-heading"><span class="ac-skill-activity-signal' + (active ? ' is-running' : '') +
      '" aria-hidden="true"></span><strong>' + (active ? 'Нейрона работает над задачей' : steps.length ? 'Ожидаем продолжение ответа' : 'Ожидаем ответ') +
      '</strong></div>' + (steps.length ? '<ol class="ac-skill-steps">' + steps.slice(-6).map(function (step) { return stepHTML(step, false); }).join('') +
      '</ol>' : '<p class="ac-skill-wait-note">Здесь появятся этапы выполнения.</p>') + '</div>';
  }
  function traceHTML(value) {
    var run = cleanRun(value);
    if (!run.skills.length && !run.tools.length && !run.steps.length && !run.warnings.length) return '';
    var badges = run.skills.slice(0, 3).map(function (skill) {
      return '<span class="ac-skill-badge">' + escapeHtml(skill.title) + '</span>';
    }).join('');
    if (run.skills.length > 3) badges += '<span class="ac-skill-badge">Ещё ' + (run.skills.length - 3) + '</span>';
    var tools = run.tools.length ? '<div class="ac-skill-trace-tools"><span>Инструменты</span>' + run.tools.map(function (tool) {
      var status = tool.status === 'running' ? 'unknown' : tool.status;
      return '<span class="ac-skill-tool-result is-' + status + '">' + escapeHtml(tool.title) + ' · ' + STATUS_LABELS[status].toLowerCase() + '</span>';
    }).join('') + '</div>' : '';
    return '<details class="ac-skill-trace"><summary><span class="ac-skill-trace-label">' +
      (run.skills.length || run.tools.length ? 'Навыки и инструменты' : 'Ход ответа') + '</span>' + badges +
      '<span class="ac-skill-trace-chevron" aria-hidden="true">⌄</span></summary><div class="ac-skill-trace-body">' + tools +
      (run.steps.length ? '<ol class="ac-skill-steps">' + run.steps.map(function (step) { return stepHTML(step, true); }).join('') + '</ol>' : '') +
      run.warnings.map(function (warning) { return '<p class="ac-skill-warning">' + escapeHtml(warning) + '</p>'; }).join('') +
      '</div></details>';
  }

  function protocolError(message) {
    var error = new Error(message || 'Не удалось полностью прочитать ответ. Обновите историю чата перед повторной отправкой.');
    error.isSkillStreamError = true;
    return error;
  }
  function validResult(value) {
    return value && typeof value === 'object' && !Array.isArray(value) &&
      (typeof value.answer === 'string' || typeof value.error === 'string');
  }
  async function readResponse(response, onStep) {
    var loginRedirect = false;
    try { loginRedirect = response.redirected && new URL(response.url, 'http://localhost').pathname === '/login'; } catch (_) {}
    if (response.status === 401 || loginRedirect) throw new Error('Сессия завершилась. Обновите страницу и войдите снова.');
    if (!response.ok) {
      var problem;
      try { problem = await response.json(); } catch (_) {}
      throw new Error(problem && (typeof problem.detail === 'string' ? problem.detail : typeof problem.error === 'string' ? problem.error : '') ||
        'Не удалось выполнить запрос. Проверьте сообщение и повторите попытку.');
    }
    var contentType = (response.headers.get('content-type') || '').toLowerCase();
    if (contentType.indexOf('application/json') !== -1 && contentType.indexOf('ndjson') === -1) {
      var json;
      try { json = await response.json(); } catch (_) { throw protocolError(); }
      if (!validResult(json)) throw protocolError();
      return json;
    }
    if (contentType.indexOf('ndjson') === -1 || !response.body || typeof response.body.getReader !== 'function') {
      throw protocolError('Сервер не передал ожидаемый ответ. Обновите историю чата перед повторной отправкой.');
    }
    var reader = response.body.getReader(), decoder = new TextDecoder('utf-8', {fatal: true});
    var buffer = '', result = null;
    function consume(line) {
      if (!line.trim()) return;
      if (line.length > MAX_RECORD) throw protocolError('Ответ слишком большой для отображения. Обновите историю чата.');
      var event;
      try { event = JSON.parse(line); } catch (_) { throw protocolError(); }
      if (!event || typeof event !== 'object' || Array.isArray(event)) throw protocolError();
      if (event.type === 'ping') return;
      if (result) throw protocolError();
      if (event.type === 'step') {
        var step = cleanStep(event);
        if (!step) throw protocolError();
        if (onStep) onStep(step);
      } else if (event.type === 'result' && validResult(event)) result = event;
      else throw protocolError();
    }
    try {
      while (true) {
        var chunk = await reader.read();
        if (chunk.done) break;
        buffer += decoder.decode(chunk.value, {stream: true});
        var newline;
        while ((newline = buffer.indexOf('\n')) !== -1) {
          consume(buffer.slice(0, newline));
          buffer = buffer.slice(newline + 1);
        }
        if (buffer.length > MAX_RECORD) throw protocolError('Ответ слишком большой для отображения. Обновите историю чата.');
      }
      buffer += decoder.decode();
      if (buffer.trim()) consume(buffer);
      if (!result) throw protocolError('Ответ не завершён. Обновите историю чата, чтобы проверить, сохранился ли результат.');
      return result;
    } catch (failure) {
      try { await reader.cancel(); } catch (_) {}
      if (failure.isSkillStreamError) throw failure;
      throw protocolError('Получение ответа прервалось. Обновите историю чата, чтобы проверить результат.');
    } finally { reader.releaseLock(); }
  }

  function normalizeCatalog(data) {
    if (!data || data.scope !== 'account' || !Array.isArray(data.items)) throw new Error('Не удалось загрузить список навыков. Попробуйте ещё раз.');
    var ids = new Set();
    var items = data.items.map(function (item) {
      if (!item || !text(item.id, 120) || item.id !== text(item.id, 120) || ids.has(item.id) || typeof item.enabled !== 'boolean') {
        throw new Error('Сервер вернул неполный список навыков. Попробуйте ещё раз.');
      }
      ids.add(item.id);
      return {id: text(item.id, 120), name: text(item.name, 120), title: russianTitle(item.title, 'Навык'),
        description: text(item.description, 800), enabled: item.enabled,
        allowed_tools: (Array.isArray(item.allowed_tools) ? item.allowed_tools : []).slice(0, 20).map(function (tool) { return text(tool, 120); }).filter(Boolean)};
    });
    return {items: items, issues: Array.isArray(data.issues) ? data.issues.map(function () { return true; }) : [], scope: 'account'};
  }
  function createCatalogStore(request) {
    var catalog = null, loading = null, saving = false;
    return {
      get: function () { return catalog ? copy(catalog) : null; },
      load: function () {
        if (loading) return loading;
        loading = request('/aichat/api/skills').then(function (data) {
          var next = normalizeCatalog(data);
          catalog = next;
          return copy(next);
        }).finally(function () { loading = null; });
        return loading;
      },
      toggle: async function (id, enabled) {
        if (saving) throw new Error('Дождитесь сохранения предыдущего изменения.');
        if (!catalog || typeof enabled !== 'boolean' || !catalog.items.some(function (item) { return item.id === id; })) {
          throw new Error('Неизвестный навык. Обновите список.');
        }
        saving = true;
        try {
          var data = await request('/aichat/api/skills/' + encodeURIComponent(id), {
            method: 'PATCH', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({enabled: enabled})
          });
          var next = normalizeCatalog(data);
          if (!next.items.some(function (item) { return item.id === id && item.enabled === enabled; })) {
            throw new Error('Сервер не подтвердил изменение навыка. Попробуйте ещё раз.');
          }
          catalog = next;
          return copy(next);
        } finally { saving = false; }
      }
    };
  }
  function create(options) {
    var doc = options.document || document;
    function el(id) { return doc.getElementById(id); }
    var store = createCatalogStore(options.request), loading = false, saving = false;
    function setError(message) {
      el('skillsError').textContent = message || '';
      el('skillsError').hidden = !message;
    }
    function syncBusy() {
      var locked = saving || !!options.isBusy();
      el('skillsBtn').disabled = locked;
      el('skillsClose').disabled = saving;
      el('skillsRetry').disabled = loading || locked;
      el('skillsCatalog').querySelectorAll('input').forEach(function (input) { input.disabled = loading || locked; });
      el('skillsPanel').setAttribute('aria-busy', String(loading || saving));
    }
    function render() {
      var catalog = store.get();
      if (!catalog) return;
      var enabled = catalog.items.filter(function (item) { return item.enabled; }).length;
      el('skillsEntryHint').textContent = 'Включено: ' + enabled + ' из ' + catalog.items.length;
      el('skillsCount').textContent = enabled ? 'Включено навыков: ' + enabled : 'Все навыки выключены';
      el('skillsEmpty').hidden = !!catalog.items.length;
      el('skillsIssues').hidden = !catalog.issues.length;
      el('skillsIssues').textContent = catalog.issues.length ? 'Часть навыков временно недоступна. Остальные можно использовать.' : '';
      el('skillsCatalog').innerHTML = catalog.items.map(function (skill, index) {
        return '<article class="ac-skill-card' + (skill.enabled ? ' is-enabled' : '') + '"><div class="ac-skill-card-heading">' +
          '<span class="ac-skill-card-icon" aria-hidden="true">' + escapeHtml(skill.title.slice(0, 1)) + '</span>' +
          '<h4 id="skillTitle' + index + '">' + escapeHtml(skill.title) + '</h4>' +
          '<label class="ac-skill-switch"><input type="checkbox" role="switch" data-skill-id="' + escapeHtml(skill.id) + '"' +
          (skill.enabled ? ' checked' : '') + ' aria-labelledby="skillTitle' + index + '" aria-describedby="skillDescription' + index + '">' +
          '<span aria-hidden="true"></span></label></div><p id="skillDescription' + index + '">' + escapeHtml(skill.description) + '</p>' +
          '<div class="ac-skill-card-tools">' + (skill.allowed_tools.length ? skill.allowed_tools.map(function (name) {
            return '<span>' + escapeHtml(toolTitle(name)) + '</span>';
          }).join('') : '<span>Работа с текстом</span>') + '</div></article>';
      }).join('');
      syncBusy();
    }
    async function load() {
      if (loading || saving) return;
      loading = true; setError('');
      el('skillsLoading').hidden = false;
      el('skillsRetry').hidden = true;
      syncBusy();
      try { await store.load(); render(); }
      catch (failure) { setError(failure.message); el('skillsRetry').hidden = false; }
      finally { loading = false; el('skillsLoading').hidden = true; syncBusy(); }
    }
    async function open() {
      if (options.isBusy() || saving) return;
      if (el('skillsOvl').hidden) options.openModal('#skillsOvl', '#skillsClose');
      el('skillsBtn').setAttribute('aria-expanded', 'true');
      await load();
    }
    async function toggle(event) {
      var input = event.target;
      if (!input.matches('input[data-skill-id]')) return;
      var catalog = store.get(), id = input.dataset.skillId;
      var skill = catalog && catalog.items.find(function (item) { return item.id === id; });
      if (!skill) return;
      var enabled = input.checked;
      input.checked = skill.enabled;
      if (saving || loading || options.isBusy()) return;
      saving = true; setError('');
      if (options.setBusy) options.setBusy(true);
      el('skillsSaveStatus').textContent = 'Сохраняем настройку…';
      syncBusy();
      try {
        await store.toggle(id, enabled);
        render();
        el('skillsSaveStatus').textContent = 'Настройка сохранена для вашей учётной записи.';
      } catch (failure) {
        setError(failure.message);
        el('skillsSaveStatus').textContent = 'Настройка не изменилась.';
      } finally {
        saving = false;
        if (options.setBusy) options.setBusy(false);
        syncBusy();
        var restored = Array.from(el('skillsCatalog').querySelectorAll('input')).find(function (item) { return item.dataset.skillId === id; });
        if (restored && !el('skillsOvl').hidden) restored.focus();
      }
    }
    el('skillsBtn').addEventListener('click', open);
    el('skillsRetry').addEventListener('click', load);
    el('skillsClose').addEventListener('click', function () { if (!saving) options.closeModal('#skillsOvl'); });
    el('skillsCatalog').addEventListener('change', toggle);
    return {open: open, syncBusy: syncBusy, isSaving: function () { return saving; }};
  }

  root.NeuronaSkills = {create: create, createCatalogStore: createCatalogStore, readResponse: readResponse,
    createTrace: createTrace, activityHTML: activityHTML, traceHTML: traceHTML};
})(typeof window !== 'undefined' ? window : globalThis);
