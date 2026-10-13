(function (global) {
  'use strict';
  const IDS = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6'];
  const API = '/api/ux-survey';
  const cleanDraft = (value) => ({
    ratings: Object.fromEntries(IDS.map(id => [id, Number.isInteger(value?.ratings?.[id]) && value.ratings[id] >= 0 && value.ratings[id] <= 5 ? value.ratings[id] : null])),
    comment: typeof value?.comment === 'string' ? value.comment : ''
  });
  const missingRatings = value => IDS.filter(id => !Number.isInteger(value.ratings[id]) || value.ratings[id] < 0 || value.ratings[id] > 5);
  const commentLength = value => Array.from(value || '').length;

  class SurveyAPI {
    constructor(fetcher) { this.fetcher = fetcher; }
    async request(path, method = 'GET', payload, options = {}) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 15000);
      try {
        const response = await this.fetcher(API + path, {
          method, credentials: 'same-origin', cache: 'no-store', signal: controller.signal,
          headers: {'Accept': 'application/json', ...(payload !== undefined ? {'Content-Type': 'application/json'} : {})},
          ...(payload !== undefined ? {body: JSON.stringify(payload)} : {}),
          ...(options.keepalive ? {keepalive: true} : {})
        });
        const data = await response.json().catch(() => null);
        if (!response.ok || !data || response.redirected) {
          const error = new Error(response.status === 401 || response.redirected
            ? 'Сессия завершилась. Войдите в сервис в другой вкладке и повторите сохранение. Введённые ответы остаются здесь.'
            : response.status === 409 ? 'Черновик изменился в другой вкладке.'
            : 'Не удалось сохранить ответы. Введённые данные остаются на этой странице. Проверьте соединение и повторите попытку.');
          error.status = response.status;
          error.data = data || {};
          error.code = data?.code || data?.detail?.code;
          error.state = data?.state || data?.detail?.state;
          throw error;
        }
        return data;
      } catch (error) {
        if (error.status) throw error;
        const failed = new Error('Нет ответа от сервера. Введённые данные остаются на этой странице. Проверьте соединение и повторите попытку.');
        failed.network = true;
        throw failed;
      } finally { clearTimeout(timer); }
    }
  }

  // Server is authoritative. A revision guards in-flight autosaves and two-tab edits.
  class SurveySession {
    constructor(api, state, notify = () => {}) {
      this.api = api;
      this.notify = notify;
      this.data = cleanDraft(state.draft);
      this.version = state.draft_version;
      this.completed = Boolean(state.completed);
      this.revision = 0;
      this.savedRevision = 0;
      this.inflight = null;
      this.conflict = null;
      this.submitting = false;
      this.timer = null;
    }
    get dirty() { return !this.completed && this.revision !== this.savedRevision; }
    change(data) {
      if (this.completed || this.submitting) return;
      this.data = cleanDraft(data);
      this.revision += 1;
      this.notify('changed');
      this.schedule();
    }
    schedule() {
      clearTimeout(this.timer);
      if (commentLength(this.data.comment) > 2000) {
        this.notify('invalid_comment');
        return;
      }
      if (!this.conflict && !this.completed && !this.submitting) {
        this.timer = setTimeout(() => this.save().catch(() => {}), 650);
      }
    }
    complete() {
      this.completed = true;
      clearTimeout(this.timer);
      this.conflict = null;
      this.notify('completed');
    }
    async save(options = {}) {
      clearTimeout(this.timer);
      if (!this.dirty) return;
      if (commentLength(this.data.comment) > 2000) {
        this.notify('invalid_comment');
        const error = new Error('Сократите комментарий до 2 000 символов. Введённый текст сохранён в форме.');
        error.validation = true;
        throw error;
      }
      if (this.conflict) throw new Error('draft_conflict');
      if (this.inflight) {
        await this.inflight;
        return this.save(options);
      }
      const revision = this.revision;
      const payload = {...cleanDraft(this.data), draft_version: this.version};
      this.notify('saving');
      this.inflight = this.api.request('/draft', 'PUT', payload, options).then(state => {
        if (state.completed) { this.complete(); return; }
        this.version = state.draft_version;
        this.savedRevision = revision;
        this.notify(this.dirty ? 'changed' : 'saved');
      }).catch(error => {
        if (error.code === 'already_submitted' || error.state?.completed) { this.complete(); return; }
        if (error.code === 'draft_conflict' && error.state) {
          this.conflict = error.state;
          this.notify('conflict', error);
        } else this.notify('save_error', error);
        throw error;
      });
      try { await this.inflight; }
      finally { this.inflight = null; }
      if (this.dirty && !this.submitting) this.schedule();
    }
    useLatest() {
      if (!this.conflict) return;
      this.data = cleanDraft(this.conflict.draft);
      this.version = this.conflict.draft_version;
      this.revision += 1;
      this.savedRevision = this.revision;
      this.conflict = null;
      this.notify('restored');
    }
    async keepMine() {
      if (!this.conflict) return this.save();
      this.version = this.conflict.draft_version;
      this.conflict = null;
      return this.save();
    }
    async submit() {
      if (this.completed || this.submitting) return;
      if (missingRatings(this.data).length || commentLength(this.data.comment) > 2000) {
        const error = new Error('Проверьте заполнение анкеты.'); error.validation = true; throw error;
      }
      if (this.conflict) throw new Error('draft_conflict');
      this.submitting = true;
      clearTimeout(this.timer);
      this.notify('submitting');
      try {
        // Let an existing save finish; a network failure must not prevent a direct safe retry of submit.
        if (this.inflight) await this.inflight.catch(() => {});
        if (this.completed) return;
        if (this.conflict) throw new Error('draft_conflict');
        const result = await this.api.request('/submit', 'POST', cleanDraft(this.data));
        if (!result.completed) throw new Error('Сервер не подтвердил сохранение. Повторите отправку.');
        this.complete();
      } catch (error) {
        if (error.code === 'already_submitted' || error.state?.completed) { this.complete(); return; }
        if (!this.conflict) this.notify('submit_error', error);
        throw error;
      } finally { this.submitting = false; this.notify('submit_finished'); }
    }
  }

  function element(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }
  function star() {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', '0 0 24 24'); svg.setAttribute('aria-hidden', 'true');
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', 'm12 3 2.78 5.63 6.22.91-4.5 4.38 1.06 6.2L12 17.19l-5.56 2.93 1.06-6.2L3 9.54l6.22-.91Z');
    svg.appendChild(path); return svg;
  }

  async function mountForm(api) {
    const form = document.getElementById('uxSurveyForm');
    const fields = document.getElementById('uxSurveyFields');
    const questionHost = document.getElementById('uxSurveyQuestions');
    const comment = document.getElementById('uxSurveyComment');
    const saveState = document.getElementById('uxSurveySaveState');
    const alert = document.getElementById('uxSurveyAlert');
    const alertText = document.getElementById('uxSurveyAlertText');
    const retry = document.getElementById('uxSurveyRetry');
    const latest = document.getElementById('uxSurveyUseLatest');
    const submit = document.getElementById('uxSurveySubmit');
    let session = null;
    let retryAction = null;
    let loadPending = false;

    function status(text, tone = '') { saveState.textContent = text; saveState.dataset.tone = tone; }
    function clearAlert() { alert.hidden = true; retry.hidden = true; latest.hidden = true; retryAction = null; }
    function showAlert(text, action, label = 'Повторить', allowLatest = false) {
      alert.hidden = false; alert.dataset.tone = 'error'; alertText.textContent = text;
      retry.hidden = !action; retry.textContent = label; retryAction = action;
      latest.hidden = !allowLatest;
    }
    function readForm() {
      return {ratings: Object.fromEntries(IDS.map(id => {
        const checked = form.querySelector('input[name="' + id + '"]:checked');
        return [id, checked ? Number(checked.value) : null];
      })), comment: comment.value};
    }
    function updateVisuals(data) {
      const missing = missingRatings(data);
      document.getElementById('uxSurveyProgress').textContent = 'Оценено ' + (6 - missing.length) + ' из 6';
      const length = commentLength(data.comment);
      const tooLong = length > 2000;
      const counter = document.getElementById('uxSurveyCommentCount');
      counter.textContent = length.toLocaleString('ru-RU') + ' / 2 000';
      counter.dataset.invalid = String(tooLong);
      comment.setAttribute('aria-invalid', String(tooLong));
      comment.setCustomValidity(tooLong ? 'Сократите комментарий до 2 000 символов.' : '');
      submit.disabled = tooLong;
      IDS.forEach(id => {
        const group = document.getElementById('uxSurvey-' + id);
        const score = data.ratings[id];
        group.querySelectorAll('label').forEach(label => {
          const input = label.querySelector('input');
          label.classList.toggle('is-filled', score !== null && Number(input.value) > 0 && Number(input.value) <= score);
        });
        document.getElementById('uxSurvey-' + id + '-value').textContent = score === null ? 'Не выбрано' : 'Оценка: ' + score + ' из 5';
        if (score !== null) { group.classList.remove('is-missing'); group.removeAttribute('aria-invalid'); }
      });
    }
    function applyForm(data) {
      IDS.forEach(id => form.querySelectorAll('input[name="' + id + '"]').forEach(input => { input.checked = data.ratings[id] === Number(input.value); }));
      comment.value = data.comment;
      updateVisuals(data);
    }
    function finish() {
      clearAlert(); form.hidden = true; fields.disabled = true;
      document.getElementById('uxSurveyComplete').hidden = false;
      document.getElementById('uxSurveyThanks').focus({preventScroll: true});
    }
    function notify(event, error) {
      if (event === 'completed') { finish(); return; }
      if (event === 'submitting') { fields.disabled = true; submit.textContent = 'Отправляем…'; status('Отправляем ответы…'); clearAlert(); return; }
      if (event === 'submit_finished') { if (!session.completed) fields.disabled = false; submit.textContent = 'Отправить ответы'; return; }
      if (event === 'changed') {
        if (commentLength(session.data.comment) > 2000) { notify('invalid_comment'); return; }
        if (session.conflict) { notify('conflict'); return; }
        status('Есть несохранённые изменения');
      }
      if (event === 'invalid_comment') {
        status('Комментарий пока не сохранён', 'error');
        showAlert('В комментарии ' + commentLength(session.data.comment).toLocaleString('ru-RU') + ' символов. Сократите его до 2 000, чтобы сохранить черновик и отправить ответы. Введённый текст не удалён.');
      }
      if (event === 'saving') status('Сохраняем черновик…');
      if (event === 'saved') { status('Черновик сохранён на сервере', 'success'); clearAlert(); }
      if (event === 'restored') { applyForm(session.data); status('Открыт сохранённый черновик', 'success'); clearAlert(); }
      if (event === 'save_error') {
        status('Черновик пока не сохранён', 'error');
        showAlert(error.message, () => session.save());
      }
      if (event === 'submit_error') {
        status('Отправка не подтверждена', 'error');
        showAlert(error.message, () => session.submit(), 'Повторить отправку');
      }
      if (event === 'conflict') {
        status('Изменения из двух вкладок', 'error');
        showAlert('В другой вкладке или на другом устройстве сохранён более новый черновик. Ваши ответы остаются в форме. Можно сохранить их вместо другого черновика или открыть сохранённый вариант.', () => session.keepMine(), 'Сохранить мой черновик', true);
      }
    }
    function renderQuestions(questions) {
      questionHost.replaceChildren();
      questions.forEach((question, index) => {
        const item = element('section', 'ux-survey-question'); item.id = 'uxSurvey-' + question.id;
        const title = element('h2'); title.id = item.id + '-title';
        title.append(element('span', 'ux-survey-question-number', (index + 1) + '.'), document.createTextNode(question.label));
        const scale = element('p', 'ux-survey-scale', '0 — ' + question.low + ', 5 — ' + question.high + '.'); scale.id = item.id + '-scale';
        const group = element('div', 'ux-survey-rating'); group.setAttribute('role', 'radiogroup');
        group.setAttribute('aria-labelledby', title.id); group.setAttribute('aria-describedby', scale.id); group.setAttribute('aria-required', 'true');
        for (let value = 0; value <= 5; value += 1) {
          if (value === 1) { const divider = element('span', 'ux-survey-rating-separator'); divider.setAttribute('aria-hidden', 'true'); group.append(divider); }
          const label = element('label', 'ux-survey-choice' + (value ? ' ux-survey-choice-star' : ''));
          const input = element('input'); input.type = 'radio'; input.name = question.id; input.value = String(value); input.required = true;
          input.setAttribute('aria-label', value + ' из 5' + (value === 0 ? ' — ' + question.low : value === 5 ? ' — ' + question.high : ''));
          const visual = element('span', 'ux-survey-choice-visual'); visual.setAttribute('aria-hidden', 'true');
          if (value) visual.append(star()); else visual.textContent = '0';
          label.append(input, visual); group.append(label);
        }
        const note = element('p', 'ux-survey-rating-note', 'Не выбрано'); note.id = item.id + '-value';
        item.append(title, scale, group, note); questionHost.append(item);
      });
    }
    async function load() {
      if (loadPending) return;
      loadPending = true; fields.disabled = true; clearAlert(); status('Загружаем черновик…');
      try {
        // Opening the form defers reminders; only submitting completes the survey.
        const state = await api.request('/open', 'POST');
        if (!Array.isArray(state.survey?.questions) || state.survey.questions.length !== 6 || !IDS.every(id => state.survey.questions.some(q => q.id === id))) {
          throw new Error('Не удалось загрузить вопросы. Повторите попытку.');
        }
        renderQuestions(state.survey.questions);
        session = new SurveySession(api, state, notify);
        applyForm(session.data);
        if (state.completed) { finish(); return; }
        fields.disabled = false;
        status(state.has_draft ? 'Сохранённый черновик восстановлен' : 'Выберите оценки — черновик сохранится автоматически', state.has_draft ? 'success' : '');
      } catch (error) {
        status('Не удалось загрузить опрос', 'error');
        showAlert(error.message, load);
      } finally { loadPending = false; }
    }
    retry.addEventListener('click', async () => {
      const action = retryAction; if (!action) return;
      retry.disabled = true;
      try { await action(); } catch (_) {} finally { retry.disabled = false; }
    });
    latest.addEventListener('click', () => session?.useLatest());
    form.addEventListener('input', () => {
      if (!session || session.submitting) return;
      const data = readForm(); updateVisuals(data); session.change(data);
    });
    form.addEventListener('submit', async event => {
      event.preventDefault(); if (!session || session.submitting || session.completed) return;
      const data = readForm();
      const missing = missingRatings(data);
      if (missing.length) {
        missing.forEach(id => {
          const card = document.getElementById('uxSurvey-' + id); card.classList.add('is-missing'); card.setAttribute('aria-invalid', 'true');
        });
        showAlert('Оцените все шесть пунктов. Оценка «0» — полноценный ответ; пустое поле ответом не считается.');
        form.querySelector('input[name="' + missing[0] + '"]').focus(); return;
      }
      if (commentLength(data.comment) > 2000) { showAlert('Сократите комментарий до 2 000 символов.'); comment.focus(); return; }
      if (session.conflict) { retry.focus(); return; }
      try { await session.submit(); } catch (_) {}
    });
    document.getElementById('uxSurveyReturn').addEventListener('click', async event => {
      if (!session?.dirty) return;
      event.preventDefault();
      try { await session.save(); if (!session.dirty) global.location.assign('/'); } catch (_) {}
    });
    global.addEventListener('beforeunload', event => {
      if (session?.dirty) { event.preventDefault(); event.returnValue = ''; }
    });
    global.addEventListener('pagehide', () => {
      if (session?.dirty && !session.conflict) session.save({keepalive: true}).catch(() => {});
    });
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden' && session?.dirty && !session.conflict) session.save({keepalive: true}).catch(() => {});
    });
    global.addEventListener('online', () => {
      if (session?.dirty && !session.conflict && !session.submitting) session.save().catch(() => {});
    });
    await load();
  }

  function safeToInvite(document, lastActivity, edited, now = Date.now()) {
    if (edited || document.visibilityState !== 'visible' || !document.hasFocus() || now - lastActivity < 15000) return false;
    const active = document.activeElement;
    if (active?.matches('input, textarea, select, [contenteditable="true"]') || active?.closest('form')) return false;
    return !Array.from(document.querySelectorAll('dialog[open], [role="dialog"]:not([hidden]), [aria-busy="true"], [aria-expanded="true"], .modal.show, .modal.open, .loading:not([hidden]), form button[type="submit"]:disabled, form input[type="submit"]:disabled'))
      .some(node => node.getClientRects().length && global.getComputedStyle(node).visibility !== 'hidden');
  }
  async function mountInvitation(api) {
    let state;
    try { state = await api.request('/state'); } catch (_) { return; }
    document.querySelectorAll('[data-ux-survey-link]').forEach(link => {
      link.textContent = state.completed ? 'Обратная связь · ответы отправлены' : state.has_draft ? 'Продолжить опрос' : 'Оценить удобство сервиса';
    });
    if (state.completed || !state.eligible || typeof HTMLDialogElement === 'undefined') return;
    let lastActivity = Date.now();
    let edited = false;
    let claimed = false;
    const recordActivity = () => { lastActivity = Date.now(); };
    ['pointerdown', 'keydown', 'scroll'].forEach(name => document.addEventListener(name, recordActivity, {passive: true}));
    document.addEventListener('input', event => { if (!event.target.closest('.ux-survey')) edited = true; }, {once: true});
    const interval = setInterval(async () => {
      if (claimed || !safeToInvite(document, lastActivity, edited)) return;
      claimed = true; clearInterval(interval);
      let claim;
      try { claim = await api.request('/invitation/claim', 'POST'); } catch (_) { return; }
      if (!claim.show || claim.state?.completed || !safeToInvite(document, lastActivity, edited)) return;
      showInvitation(api, claim.token, Boolean(claim.state?.has_draft));
    }, 3000);
    global.addEventListener('pagehide', () => clearInterval(interval), {once: true});
  }
  function showInvitation(api, token, hasDraft) {
    const dialog = element('dialog', 'ux-survey ux-survey-dialog');
    dialog.setAttribute('aria-labelledby', 'uxSurveyInviteTitle'); dialog.setAttribute('aria-describedby', 'uxSurveyInviteDescription');
    const close = element('button', 'ux-survey-close', '×'); close.type = 'button'; close.setAttribute('aria-label', 'Закрыть и напомнить через 24 часа');
    const kicker = element('div', 'ux-survey-kicker', 'Обратная связь');
    const title = element('h2', '', 'Помогите сделать сервис удобнее'); title.id = 'uxSurveyInviteTitle';
    const description = element('p', '', 'Оцените удобство и оформление сервиса. Ваши ответы помогут нам определить, что улучшить. Результаты доступны только администраторам.'); description.id = 'uxSurveyInviteDescription';
    const actions = element('div', 'ux-survey-actions');
    const start = element('a', 'ux-survey-button ux-survey-button-primary', 'Пройти опрос'); start.href = '/feedback/survey';
    const later = element('button', 'ux-survey-button', 'Напомнить позже'); later.type = 'button'; later.autofocus = true;
    actions.append(start, later); dialog.append(close, kicker, title, description);
    if (hasDraft) dialog.append(element('p', '', 'У вас есть сохранённый черновик. Можно продолжить с того же места.'));
    dialog.append(actions); document.body.append(dialog);
    let dismissed = false;
    async function defer() {
      if (dismissed) return; dismissed = true; dialog.close(); dialog.remove();
      // Claim already reserves 24 hours on the server even if this request loses connectivity.
      for (let attempt = 0; attempt < 3; attempt += 1) {
        try { await api.request('/later', 'POST', undefined, {keepalive: true}); return; }
        catch (_) { if (attempt < 2) await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1))); }
      }
    }
    close.addEventListener('click', defer); later.addEventListener('click', defer);
    dialog.addEventListener('cancel', event => { event.preventDefault(); defer(); });
    dialog.addEventListener('close', () => { if (!dismissed) defer(); });
    start.addEventListener('click', () => { dismissed = true; dialog.close(); dialog.remove(); });
    try { dialog.showModal(); }
    catch (_) { dialog.remove(); return; }
    // Count exposure only after a dialog was actually presented. Retrying the same token is safe.
    (async () => {
      for (let attempt = 0; attempt < 3; attempt += 1) {
        try { await api.request('/invitation/shown', 'POST', {token}, {keepalive: true}); return; }
        catch (_) { if (attempt < 2) await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1))); }
      }
    })();
  }

  const exported = {cleanDraft, missingRatings, commentLength, SurveyAPI, SurveySession, safeToInvite};
  if (typeof module !== 'undefined' && module.exports) module.exports = exported;
  if (typeof document !== 'undefined') {
    const boot = () => {
      const api = new SurveyAPI(global.fetch.bind(global));
      if (document.getElementById('uxSurveyPage')) mountForm(api);
      else if (global.location.pathname === '/' || global.location.pathname === '/users') mountInvitation(api);
    };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot, {once: true}); else boot();
  }
})(typeof window !== 'undefined' ? window : globalThis);
