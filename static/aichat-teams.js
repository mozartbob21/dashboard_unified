(function (root) {
  'use strict';

  function copyTeam(team) {
    return team ? {preset: team.preset, roles: team.roles.slice()} : null;
  }
  function escapeHtml(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, function (char) {
      return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[char];
    });
  }
  function roleCount(count) {
    var rest = count % 100, last = count % 10;
    return count + ' ' + (rest >= 11 && rest <= 14 ? 'ролей' : last === 1 ? 'роль' : last >= 2 && last <= 4 ? 'роли' : 'ролей');
  }
  function messageBadge(team) {
    if (!team || !Array.isArray(team.roles) || !team.roles.length) return '';
    return '<span class="ac-team-answer-badge">ИИ-команда · ' + roleCount(team.roles.length) + '</span>';
  }

  // Commit the visible state only after the server accepts the change.
  function createStore(request) {
    var current = null;
    return {
      get: function () { return copyTeam(current); },
      restore: function (team) { current = copyTeam(team); },
      apply: async function (dialogId, selection) {
        var next = copyTeam(selection);
        if (dialogId) {
          var response = await request('/aichat/api/dialogs/' + encodeURIComponent(dialogId) + '/team', {
            method: 'PATCH', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({team: next})
          });
          if (!Object.prototype.hasOwnProperty.call(response, 'team')) {
            throw new Error('Сервер не подтвердил состав команды. Повторите попытку.');
          }
          next = copyTeam(response.team);
        }
        current = next;
        return copyTeam(current);
      }
    };
  }

  function create(options) {
    var doc = options.document || document;
    function el(id) { return doc.getElementById(id); }
    var store = createStore(options.request), catalog = null, catalogPromise = null, draft = null, saving = false;
    var icon = '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3"/><path d="M3 20v-2a6 6 0 0 1 12 0v2m0-16a3 3 0 0 1 0 6m3 4a5 5 0 0 1 3 4v2"/></svg>';

    function names(team) {
      if (!catalog) return roleCount(team.roles.length);
      return team.roles.map(function (id) {
        var found = catalog && catalog.roles.find(function (role) { return role.id === id; });
        return found ? found.name : id;
      }).join(', ');
    }
    function presetName(team) {
      var preset = catalog && catalog.presets.find(function (item) { return item.id === team.preset; });
      return preset ? preset.name : 'ИИ-команда';
    }
    function renderActive() {
      var team = store.get();
      el('teamBtn').classList.toggle('is-active', !!team);
      el('teamEntryHint').textContent = team ? roleCount(team.roles.length) + ' · изменить состав' : 'Несколько взглядов на задачу';
      el('teamStrip').hidden = !team;
      if (team) {
        el('teamActiveName').textContent = presetName(team);
        el('teamActiveRoles').textContent = names(team);
        el('teamStrip').title = names(team);
      }
      if (options.onChange) options.onChange(team);
    }
    function error(message) {
      el('teamError').textContent = message || '';
      el('teamError').hidden = !message;
    }
    function renderDraft() {
      if (!catalog || !draft) return;
      var selected = draft.roles;
      el('teamPresets').innerHTML = catalog.presets.map(function (preset) {
        return '<button type="button" class="ac-team-preset' + (draft.preset === preset.id ? ' selected' : '') +
          '" data-preset="' + escapeHtml(preset.id) + '" aria-pressed="' + (draft.preset === preset.id) + '">' +
          '<strong>' + escapeHtml(preset.name) + '</strong><span>' + escapeHtml(preset.description) + '</span></button>';
      }).join('');
      el('teamRoles').innerHTML = catalog.roles.map(function (role, index) {
        var checked = selected.indexOf(role.id) !== -1;
        return '<label class="ac-team-role' + (checked ? ' selected' : '') + '">' +
          '<input type="checkbox" value="' + escapeHtml(role.id) + '"' + (checked ? ' checked' : '') +
          ' aria-describedby="teamRoleDescription' + index + '">' +
          '<span class="ac-team-role-mark" aria-hidden="true">' + escapeHtml(role.name.slice(0, 1)) + '</span>' +
          '<span class="ac-team-role-text"><strong>' + escapeHtml(role.name) + '</strong>' +
          '<span id="teamRoleDescription' + index + '">' + escapeHtml(role.description) + '</span></span>' +
          '<span class="ac-team-role-check" aria-hidden="true">✓</span></label>';
      }).join('');
      el('teamSelectionCount').textContent = roleCount(selected.length);
      el('teamSelectionNames').textContent = selected.length ? names(draft) : 'Добавьте роли из списка';
      el('teamApply').disabled = saving || !selected.length;
      syncBusy();
    }
    function syncBusy() {
      var locked = !!options.isBusy();
      ['teamBtn', 'teamEdit', 'teamReset'].forEach(function (id) { el(id).disabled = locked; });
      el('teamOvl').querySelectorAll('button,input').forEach(function (control) { control.disabled = saving; });
      el('teamApply').disabled = saving || !catalog || !draft || !draft.roles.length;
      el('teamApply').textContent = saving ? 'Сохраняем…' : 'Применить команду';
      el('teamPanel').setAttribute('aria-busy', String(saving));
    }
    function loadCatalog() {
      if (catalog) return Promise.resolve(catalog);
      if (!catalogPromise) {
        catalogPromise = options.request('/aichat/api/teams').then(function (data) {
          if (!Array.isArray(data.roles) || !Array.isArray(data.presets) || !data.roles.length || !data.presets.length) {
            throw new Error('Список ролей пока недоступен. Повторите попытку.');
          }
          catalog = data;
          return data;
        }).finally(function () { catalogPromise = null; });
      }
      return catalogPromise;
    }
    async function open() {
      if (options.isBusy()) return;
      error('');
      el('teamRetry').hidden = true;
      if (el('teamOvl').hidden) options.openModal('#teamOvl', '#teamClose');
      else el('teamClose').focus();
      el('teamBtn').setAttribute('aria-expanded', 'true');
      if (!catalog) {
        el('teamLoading').hidden = false;
        el('teamLayout').hidden = true;
        el('teamApply').disabled = true;
        try {
          await loadCatalog();
        } catch (failure) {
          error(failure.message);
          el('teamRetry').hidden = false;
          return;
        } finally { el('teamLoading').hidden = true; }
      }
      var initial = catalog.presets.find(function (preset) { return preset.id === 'council'; }) || catalog.presets[0];
      draft = store.get() || {preset: initial.id, roles: initial.roles.slice()};
      el('teamLayout').hidden = false;
      renderDraft();
      renderActive();
    }
    async function apply(selection, closeAfter) {
      if (options.isBusy() || saving) return;
      error(''); saving = true; options.setBusy(true); syncBusy();
      var applied = false;
      try {
        await store.apply(options.getDialogId(), selection);
        renderActive();
        applied = true;
      } catch (failure) {
        if (el('teamOvl').hidden) {
          options.showError(failure.message);
        } else error(failure.message);
      } finally {
        saving = false; options.setBusy(false); syncBusy();
        if (applied && closeAfter) options.closeModal('#teamOvl');
      }
    }
    el('teamPresets').addEventListener('click', function (event) {
      var button = event.target.closest('[data-preset]');
      if (!button || saving) return;
      var preset = catalog.presets.find(function (item) { return item.id === button.dataset.preset; });
      if (!preset) return;
      draft = {preset: preset.id, roles: preset.id === 'custom' ? draft.roles.slice() : preset.roles.slice()};
      error(''); renderDraft();
      el('teamPresets').querySelector('[data-preset="' + preset.id + '"]').focus();
    });
    el('teamRoles').addEventListener('change', function (event) {
      if (saving || !event.target.matches('input[type="checkbox"]')) return;
      var roleId = event.target.value;
      if (!catalog.roles.some(function (role) { return role.id === roleId; })) return;
      var selected = draft.roles.filter(function (id) { return id !== roleId; });
      if (event.target.checked) selected.push(roleId);
      draft = {preset: 'custom', roles: selected};
      error(''); renderDraft();
      Array.from(el('teamRoles').querySelectorAll('input')).find(function (input) { return input.value === roleId; }).focus();
    });
    el('teamBtn').addEventListener('click', open);
    el('teamEdit').addEventListener('click', open);
    el('teamRetry').addEventListener('click', open);
    el('teamReset').addEventListener('click', function () { apply(null, false); });
    el('teamStandard').addEventListener('click', function () { apply(null, true); });
    el('teamClose').addEventListener('click', function () { if (!saving) options.closeModal('#teamOvl'); });
    el('teamApply').addEventListener('click', function () { if (draft && draft.roles.length) apply(draft, true); });
    el('teamEntryIcon').innerHTML = icon;
    renderActive();
    loadCatalog().then(renderActive).catch(function () { /* The chooser provides retry and an error state. */ });
    return {
      get: store.get,
      restore: function (team) { store.restore(team); renderActive(); },
      syncBusy: syncBusy,
      isSaving: function () { return saving; },
      open: open
    };
  }

  root.NeuronaTeams = {create: create, createStore: createStore, messageBadge: messageBadge, roleCount: roleCount};
})(typeof window !== 'undefined' ? window : globalThis);
