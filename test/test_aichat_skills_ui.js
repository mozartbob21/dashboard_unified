// Offline only: node --test test/test_aichat_skills_ui.js
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');
const projectRoot = process.env.AICHAT_PROJECT_ROOT || join(__dirname, '..');
const source = readFileSync(join(projectRoot, 'static/aichat-skills.js'), 'utf8');
const template = readFileSync(join(projectRoot, 'templates/aichat.html'), 'utf8');
const context = vm.createContext({TextDecoder, URL});
vm.runInContext(source, context);
const skills = context.NeuronaSkills;
const plain = value => JSON.parse(JSON.stringify(value));
const step = (status = 'running', id = 'calc') => ({type: 'step', id, kind: 'tool', label: 'Считаю 18% от суммы', status});
const result = {type: 'result', answer: 'Результат: 44 100 ₽', dialog_id: 'saved', team: null, ok: true};
const ndjson = events => events.map(value => JSON.stringify(value)).join('\n');
const bytes = value => new TextEncoder().encode(value);

function responseFromChunks(chunks, failAt = -1) {
  let cursor = 0, cancelled = 0, released = 0;
  return {
    ok: true, status: 200, headers: {get: () => 'application/x-ndjson; charset=utf-8'},
    body: {getReader: () => ({
      async read() {
        if (cursor === failAt) throw new Error('Mock connection lost');
        return cursor < chunks.length ? {value: chunks[cursor++], done: false} : {done: true};
      },
      async cancel() { cancelled++; },
      releaseLock() { released++; }
    })},
    stats: () => ({cancelled, released})
  };
}

test('NDJSON preserves UTF-8 across byte boundaries and ignores heartbeat events', async () => {
  const payload = bytes(ndjson([step(), {type: 'ping'}, step('done'), result]));
  const response = responseFromChunks(Array.from(payload, byte => new Uint8Array([byte])));
  const seen = [];
  const answer = await skills.readResponse(response, event => seen.push(event));
  assert.deepEqual(plain(answer), result);
  assert.deepEqual(seen.map(event => event.status), ['running', 'done']);
  assert.equal(seen[0].label, 'Считаю 18% от суммы');
  assert.deepEqual(response.stats(), {cancelled: 0, released: 1});
});

test('step callbacks run while the final response is still pending', async () => {
  let finish;
  const finalChunk = new Promise(resolve => { finish = resolve; });
  let reads = 0;
  const seen = [];
  const response = responseFromChunks([]);
  response.body.getReader = () => ({
    async read() {
      if (reads++ === 0) return {value: bytes(JSON.stringify(step()) + '\r\n\n'), done: false};
      if (reads === 2) return finalChunk;
      return {done: true};
    },
    async cancel() {}, releaseLock() {}
  });
  const waiting = skills.readResponse(response, event => seen.push(event));
  await new Promise(setImmediate);
  assert.equal(seen.length, 1);
  assert.equal(seen[0].status, 'running');
  finish({value: bytes(JSON.stringify(result) + '\r\n'), done: false});
  assert.equal((await waiting).answer, result.answer);
});

test('a missing final result, malformed record, invalid status or duplicate result fails without retry', async () => {
  const invalid = [
    '', ndjson([step(), {type: 'ping'}]), JSON.stringify(step()) + '\n{broken',
    ndjson([{...step(), status: 'invented'}, result]), ndjson([{type: 'private_reasoning', content: 'hidden'}]),
    ndjson([result, result]), ndjson([result, step()])
  ];
  for (const input of invalid) {
    const response = responseFromChunks([bytes(input)]);
    await assert.rejects(skills.readResponse(response), /ответ|Ответ/);
    assert.deepEqual(response.stats(), {cancelled: 1, released: 1});
  }
});

test('truncated UTF-8 and connection failures terminate the stream reader', async () => {
  const invalidUtf8 = responseFromChunks([new Uint8Array([0xd0])]);
  await assert.rejects(skills.readResponse(invalidUtf8), /прервалось/);
  const disconnected = responseFromChunks([bytes(JSON.stringify(step()) + '\n')], 1);
  await assert.rejects(skills.readResponse(disconnected), /прервалось/);
  assert.deepEqual(disconnected.stats(), {cancelled: 1, released: 1});
});

test('normal JSON errors preserve server detail and sessions get a clear login message', async () => {
  await assert.rejects(skills.readResponse({status: 400, ok: false, json: async () => ({detail: 'Файл слишком большой'})}), /Файл слишком большой/);
  await assert.rejects(skills.readResponse({status: 403, ok: false, json: async () => ({error: 'Нет доступа'})}), /Нет доступа/);
  await assert.rejects(skills.readResponse({status: 401, ok: false}), /Сессия завершилась/);
  await assert.rejects(skills.readResponse({status: 200, ok: true, redirected: true, url: 'https://local.invalid/login'}), /Сессия завершилась/);
  const answer = await skills.readResponse({status: 200, ok: true, headers: {get: () => 'application/json'}, json: async () => ({answer: 'Обычный ответ'})});
  assert.equal(answer.answer, 'Обычный ответ');
});

test('trace is driven by public events, collapses updates and terminates running steps on failure', () => {
  const trace = skills.createTrace();
  trace.update({...step(), reasoning: 'PRIVATE_REASONING', arguments: {secret: 'PRIVATE_ARGUMENTS'}});
  trace.update(step('done'));
  trace.update({...step('running', 'model'), kind: 'model', label: 'Готовлю ответ'});
  assert.equal(trace.snapshot().steps.length, 2);
  assert.match(skills.activityHTML(trace.snapshot()), /is-running/);
  const failed = trace.finish(null, true);
  assert.deepEqual(plain(failed.steps.map(item => item.status)), ['done', 'error']);
  const html = skills.traceHTML(failed);
  assert.doesNotMatch(html, /is-running|PRIVATE_REASONING|PRIVATE_ARGUMENTS/);
  assert.match(html, /<details class="ac-skill-trace">/);
  assert.match(html, /Не выполнено/);
});

test('saved trace uses final metadata, escapes labels and leaves old messages unchanged', () => {
  const trace = skills.createTrace();
  trace.update(step());
  const final = trace.finish({
    skills: [{id: 'calculation', title: '<img src=x onerror=evil()>Навык'}],
    tools: [{name: 'calculate', title: '<script>evil()</script>Инструмент', status: 'done', arguments: 'SECRET'}],
    steps: [{...step('done'), label: '<svg onload=evil()>'}], warnings: ['<iframe src=x>'], reasoning: 'PRIVATE'
  }, false);
  const html = skills.traceHTML(final);
  assert.match(html, /&lt;img/);
  assert.doesNotMatch(html, /<(?:img|script|svg|iframe)\b|SECRET|PRIVATE|is-running/);
  assert.equal(skills.traceHTML(undefined), '');
  assert.equal(skills.traceHTML(null), '');
  assert.equal(skills.traceHTML({}), '');
  assert.doesNotMatch(skills.traceHTML({steps: [step()]}), /is-running/);
  assert.match(skills.traceHTML({steps: [step()]}), /Статус не получен/);
});

const catalog = enabled => ({scope: 'account', issues: [], items: [
  {id: 'calculation', name: 'calculation', title: 'Точные расчёты', description: 'Проверка арифметики', allowed_tools: ['calculate'], enabled},
  {id: 'tables', name: 'tables', title: 'Анализ таблиц', description: 'План и факт', allowed_tools: ['table_profile'], enabled: true}
]});

test('account preference stays unchanged until the server confirms a successful PATCH', async () => {
  let resolveSave, patch;
  const store = skills.createCatalogStore(async (url, options) => {
    if (!options) return catalog(false);
    patch = {url, options};
    return new Promise(resolve => { resolveSave = resolve; });
  });
  await store.load();
  const changing = store.toggle('calculation', true);
  assert.equal(store.get().items[0].enabled, false);
  assert.equal(patch.url, '/aichat/api/skills/calculation');
  assert.equal(patch.options.method, 'PATCH');
  assert.deepEqual(JSON.parse(patch.options.body), {enabled: true});
  await assert.rejects(store.toggle('tables', false), /Дождитесь/);
  resolveSave(catalog(true));
  await changing;
  assert.equal(store.get().items[0].enabled, true);
  const copied = store.get();
  copied.items[0].enabled = false;
  assert.equal(store.get().items[0].enabled, true);
});

test('failed or unconfirmed skill saves cannot overwrite account preferences', async () => {
  for (const save of [async () => { throw new Error('Mock offline'); }, async () => catalog(false), async () => ({scope: 'global', items: []})]) {
    const store = skills.createCatalogStore(async (url, options) => options ? save() : catalog(false));
    await store.load();
    await assert.rejects(store.toggle('calculation', true));
    assert.equal(store.get().items[0].enabled, false);
  }
});

test('unknown skill IDs and nonboolean toggles never reach the server', async () => {
  let patches = 0;
  const store = skills.createCatalogStore(async (url, options) => { if (options) patches++; return catalog(false); });
  await store.load();
  await assert.rejects(store.toggle('unknown', true), /Неизвестный навык/);
  await assert.rejects(store.toggle('calculation', 'true'), /Неизвестный навык/);
  assert.equal(patches, 0);
});

test('internal English identifiers never become skill or tool titles', async () => {
  const raw = catalog(false);
  delete raw.items[0].title;
  const store = skills.createCatalogStore(async () => raw);
  await store.load();
  assert.equal(store.get().items[0].title, 'Навык');
  const html = skills.traceHTML({
    skills: [{id: 'technical-skill-id', title: 'technical-skill-id'}],
    tools: [{name: 'unknown_internal_tool', title: 'unknown_internal_tool', status: 'done'}]
  });
  assert.match(html, /Навык/);
  assert.match(html, /Инструмент навыка/);
  assert.doesNotMatch(html, /technical-skill-id|unknown_internal_tool/);
});

// Minimal DOM adapter executes the real controller. Initial hidden state comes
// from the actual template, and rendered checkbox state comes from its HTML.
function templateDocument() {
  const nodes = new Map();
  const doc = {activeElement: null, getElementById: id => nodes.get(id)};
  for (const match of template.matchAll(/<([a-z][\w:-]*)\b([^>]*\bid="([^"]+)"[^>]*)>/gi)) {
    nodes.set(match[3], {
      id: match[3], hidden: /(?:^|\s)hidden(?:\s|=|$)/.test(match[2]), disabled: false,
      textContent: '', innerHTML: '', listeners: {}, attrs: {},
      addEventListener(name, callback) { this.listeners[name] = callback; },
      setAttribute(name, value) { this.attrs[name] = String(value); },
      focus() { doc.activeElement = this; }
    });
  }
  const list = nodes.get('skillsCatalog');
  let html = '', inputs = [];
  Object.defineProperty(list, 'innerHTML', {
    get: () => html,
    set: value => {
      html = value;
      inputs = Array.from(value.matchAll(/<input\b([^>]*data-skill-id="([^"]+)"[^>]*)>/g), match => ({
        dataset: {skillId: match[2]}, checked: /\schecked(?:\s|$)/.test(match[1]), disabled: false,
        matches: selector => selector === 'input[data-skill-id]',
        focus() { doc.activeElement = this; }
      }));
    }
  });
  list.querySelectorAll = selector => { assert.equal(selector, 'input'); return inputs; };
  return doc;
}

function controllerHarness(request) {
  const doc = templateDocument();
  let busy = false;
  const ui = skills.create({document: doc, request, isBusy: () => busy, setBusy: value => { busy = value; },
    openModal: (selector, focus) => { doc.getElementById(selector.slice(1)).hidden = false; doc.getElementById(focus.slice(1)).focus(); },
    closeModal: selector => { doc.getElementById(selector.slice(1)).hidden = true; }
  });
  return {ui, doc, busy: () => busy};
}

test('catalog loading, retry and personal scope render without affecting team mode', async () => {
  let attempts = 0;
  const {ui, doc} = controllerHarness(async () => {
    if (++attempts === 1) throw new Error('Список временно недоступен');
    return catalog(false);
  });
  await ui.open();
  assert.equal(doc.getElementById('skillsLoading').hidden, true);
  assert.equal(doc.getElementById('skillsError').hidden, false);
  assert.equal(doc.getElementById('skillsRetry').hidden, false);
  assert.equal(doc.getElementById('teamOvl').hidden, true);
  await doc.getElementById('skillsRetry').listeners.click();
  assert.equal(doc.getElementById('skillsError').hidden, true);
  assert.equal(doc.getElementById('skillsRetry').hidden, true);
  assert.equal(doc.getElementById('skillsLoading').hidden, true);
  assert.equal(doc.getElementById('skillsCount').textContent, 'Включено навыков: 1');
  assert.match(doc.getElementById('skillsCatalog').innerHTML, /Вычисление/);
  assert.equal(doc.getElementById('skillsCatalog').querySelectorAll('input').length, 2);
});

test('controller reverses a native optimistic toggle, locks while saving and restores focus on failure', async () => {
  let rejectSave;
  const {ui, doc, busy} = controllerHarness(async (url, options) => {
    if (!options) return catalog(false);
    return new Promise((resolve, reject) => { rejectSave = reject; });
  });
  await ui.open();
  const list = doc.getElementById('skillsCatalog');
  const input = list.querySelectorAll('input')[0];
  input.checked = true;
  const saving = list.listeners.change({target: input});
  assert.equal(input.checked, false);
  assert.equal(input.disabled, true);
  assert.equal(busy(), true);
  assert.equal(ui.isSaving(), true);
  rejectSave(new Error('Не удалось сохранить'));
  await saving;
  assert.equal(input.checked, false);
  assert.equal(input.disabled, false);
  assert.equal(busy(), false);
  assert.equal(ui.isSaving(), false);
  assert.equal(doc.activeElement, input);
  assert.equal(doc.getElementById('skillsError').textContent, 'Не удалось сохранить');
});

function sendHarness(response) {
  const calls = [], assistants = [], activity = [], busyValues = [];
  const fields = new Map();
  const ctx = vm.createContext({
    window: {NeuronaSkills: skills}, busy: false, currentDid: null, pending: [],
    ta: {value: 'Посчитай 18% от 245 000', style: {}}, teamUI: {get: () => null, restore() {}},
    $: selector => selector === '#welcome' ? null : {hidden: true, textContent: ''},
    setBusy(value) { ctx.busy = value; busyValues.push(value); },
    appendUser() {}, renderPending() {}, loadDialogs() {},
    showSkillActivity(run) { activity.push(plain(run)); }, clearSkillActivity() {},
    appendAssist(answer, team, run) { assistants.push({answer, team, run: plain(run)}); },
    fetch: async (url, options) => { calls.push({url, options}); return response; },
    FormData: class {append(key, value) { fields.set(key, value); }},
    TypeError, URL
  });
  const start = template.indexOf('async function send(textOverride){');
  const end = template.indexOf('\nfunction setBusy(', start);
  assert.ok(start !== -1 && end > start);
  vm.runInContext(template.slice(start, end), ctx);
  return {send: ctx.send, ctx, calls, assistants, activity, busyValues, fields};
}

test('real template send performs one streaming POST and renders the persisted trace', async () => {
  const skillRun = {skills: [{id: 'calculation', title: 'Точные расчёты'}], tools: [{name: 'calculate', title: 'Вычисление', status: 'done'}], steps: [step('done')], warnings: []};
  const run = sendHarness(responseFromChunks([bytes(ndjson([step(), step('done'), {...result, skill_run: skillRun}]))]));
  await run.send();
  assert.equal(run.calls.length, 1);
  assert.equal(run.calls[0].url, '/aichat/api/send');
  assert.equal(run.calls[0].options.method, 'POST');
  assert.equal(run.calls[0].options.headers.Accept, 'application/x-ndjson');
  assert.equal(run.fields.get('skills_mode'), 'auto');
  assert.equal(run.ctx.currentDid, 'saved');
  assert.equal(run.assistants[0].answer, result.answer);
  assert.equal(run.assistants[0].run.skills[0].title, 'Точные расчёты');
  assert.equal(run.activity.length, 3);
  assert.deepEqual(run.busyValues, [true, false]);
});

test('real template send never resubmits a broken stream and ends its running activity', async () => {
  const run = sendHarness(responseFromChunks([bytes(JSON.stringify(step()) + '\n{broken')]));
  await run.send();
  assert.equal(run.calls.length, 1);
  assert.equal(run.assistants.length, 1);
  assert.match(run.assistants[0].answer, /Не удалось получить ответ/);
  assert.equal(run.assistants[0].run.steps[0].status, 'error');
  assert.deepEqual(run.busyValues, [true, false]);
});
