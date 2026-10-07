// Run with: node --test test/test_aichat_teams_ui.js
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');
const projectRoot = process.env.AICHAT_PROJECT_ROOT || join(__dirname, '..');
const context = vm.createContext({});
vm.runInContext(readFileSync(join(projectRoot, 'static/aichat-teams.js'), 'utf8'), context);
const teams = context.NeuronaTeams;
const plain = value => JSON.parse(JSON.stringify(value));
const council = {preset: 'council', roles: ['advisor', 'researcher', 'critic']};
const custom = {preset: 'custom', roles: ['designer']};

test('choosing a team before the first message needs no dialog mutation', async () => {
  const store = teams.createStore(() => { throw new Error('Unexpected request'); });
  await store.apply(null, council);
  assert.deepEqual(plain(store.get()), council);
  await store.apply(null, null);
  assert.equal(store.get(), null);
});

test('an existing dialog keeps its visible team until PATCH succeeds', async () => {
  let resolveRequest;
  let captured;
  const store = teams.createStore((url, options) => {
    captured = {url, options};
    return new Promise(resolve => { resolveRequest = resolve; });
  });
  store.restore(council);
  const saved = store.apply('dialog/one', custom);
  assert.deepEqual(plain(store.get()), council);
  assert.equal(captured.url, '/aichat/api/dialogs/dialog%2Fone/team');
  assert.equal(captured.options.method, 'PATCH');
  assert.deepEqual(JSON.parse(captured.options.body), {team: custom});
  resolveRequest({team: custom});
  await saved;
  assert.deepEqual(plain(store.get()), custom);
});

test('failed save leaves the previous active team intact', async () => {
  const store = teams.createStore(async () => { throw new Error('Network unavailable'); });
  store.restore(council);
  await assert.rejects(store.apply('existing', custom), /Network unavailable/);
  assert.deepEqual(plain(store.get()), council);
  await assert.rejects(store.apply('existing', null), /Network unavailable/);
  assert.deepEqual(plain(store.get()), council);
});

test('unconfirmed server response cannot silently disable team mode', async () => {
  const store = teams.createStore(async () => ({ok: true}));
  store.restore(council);
  await assert.rejects(store.apply('existing', custom), /не подтвердил состав/);
  assert.deepEqual(plain(store.get()), council);
});

test('server canonical team and explicit ordinary mode are authoritative', async () => {
  let result = {preset: 'custom', roles: ['advisor', 'designer']};
  const store = teams.createStore(async () => ({team: result}));
  await store.apply('existing', {preset: 'custom', roles: ['designer', 'advisor']});
  assert.deepEqual(plain(store.get()), result);
  result = null;
  await store.apply('existing', null);
  assert.equal(store.get(), null);
});

test('dialog restoration and draft edits cannot mutate saved team by reference', async () => {
  const original = {preset: 'custom', roles: ['advisor'], prompt: 'client prompt'};
  const store = teams.createStore(async () => ({team: original}));
  store.restore(original);
  original.roles.push('critic');
  const draft = store.get();
  draft.roles.push('designer');
  draft.preset = 'creative';
  assert.deepEqual(plain(store.get()), {preset: 'custom', roles: ['advisor']});
  assert.equal(Object.hasOwn(store.get(), 'prompt'), false);
  store.restore(null);
  assert.equal(store.get(), null);
});

test('answer badges use saved message metadata and never role strings as markup', () => {
  const badge = teams.messageBadge(council);
  assert.match(badge, /ИИ-команда · 3 роли/);
  assert.equal(teams.messageBadge(null), '');
  assert.equal(teams.messageBadge(undefined), '');
  assert.equal(teams.messageBadge({roles: []}), '');
  const malicious = teams.messageBadge({preset: '<script>', roles: ['<img src=x onerror=alert(1)>']});
  assert.match(malicious, /1 роль/);
  assert.doesNotMatch(malicious, /<script|<img|onerror/);
});

test('role counts remain readable for the complete eleven-role catalog', () => {
  assert.equal(teams.roleCount(0), '0 ролей');
  assert.equal(teams.roleCount(1), '1 роль');
  assert.equal(teams.roleCount(2), '2 роли');
  assert.equal(teams.roleCount(5), '5 ролей');
  assert.equal(teams.roleCount(11), '11 ролей');
});

// The controller needs only a small part of the DOM for its loading lifecycle.
// Read initial states from the real template so losing a `hidden` attribute is
// observable; do not replace the template with a fixture that assumes the fix.
function templateDocument() {
  const source = readFileSync(join(projectRoot, 'templates/aichat.html'), 'utf8');
  const nodes = new Map();
  const doc = {activeElement: null, getElementById: id => nodes.get(id)};
  for (const match of source.matchAll(/<([a-z][\w:-]*)\b([^>]*\bid="([^"]+)"[^>]*)>/gi)) {
    const classes = new Set();
    const attributes = {};
    const node = {
      id: match[3], tagName: match[1].toUpperCase(), offset: match.index,
      hidden: /(?:^|\s)hidden(?:\s|=|$)/.test(match[2]),
      disabled: /(?:^|\s)disabled(?:\s|=|$)/.test(match[2]),
      textContent: '', innerHTML: '', listeners: {},
      classList: {toggle: (name, active) => active ? classes.add(name) : classes.delete(name)},
      setAttribute(name, value) { attributes[name] = String(value); },
      getAttribute(name) { return attributes[name]; },
      addEventListener(name, callback) { this.listeners[name] = callback; },
      focus() { doc.activeElement = this; }
    };
    nodes.set(node.id, node);
  }
  const overlay = nodes.get('teamOvl');
  const overlayEnd = source.indexOf('</section>', overlay.offset);
  overlay.querySelectorAll = selector => {
    assert.equal(selector, 'button,input');
    return Array.from(nodes.values()).filter(node =>
      node.offset > overlay.offset && node.offset < overlayEnd && /^(BUTTON|INPUT)$/.test(node.tagName));
  };
  return doc;
}

const smallCatalog = {
  roles: [
    {id: 'advisor', name: 'Советник', description: 'Цели и приоритеты'},
    {id: 'researcher', name: 'Исследователь', description: 'Факты и источники'},
    {id: 'critic', name: 'Критик', description: 'Риски и контраргументы'}
  ],
  presets: [{id: 'council', name: 'Совет экспертов', description: 'Проверить решение', roles: council.roles}]
};

function controllerHarness(fetchCatalog) {
  const doc = templateDocument();
  let requests = 0;
  const ui = teams.create({
    document: doc,
    request: async url => {
      assert.equal(url, '/aichat/api/teams');
      requests++;
      return fetchCatalog();
    },
    getDialogId: () => null,
    isBusy: () => false,
    setBusy: () => {},
    openModal: (selector, focusSelector) => {
      doc.getElementById(selector.slice(1)).hidden = false;
      doc.getElementById(focusSelector.slice(1)).focus();
    },
    closeModal: selector => { doc.getElementById(selector.slice(1)).hidden = true; },
    showError: message => assert.fail(message)
  });
  return {ui, doc, requestCount: () => requests};
}

test('opening after background catalog load shows the chooser without a stale loading row', async () => {
  const {ui, doc, requestCount} = controllerHarness(async () => smallCatalog);
  await new Promise(setImmediate);
  ui.restore(council);
  await ui.open();
  assert.equal(doc.getElementById('teamOvl').hidden, false);
  assert.equal(doc.getElementById('teamLayout').hidden, false);
  assert.equal(doc.getElementById('teamLoading').hidden, true);
  assert.equal(doc.getElementById('teamApply').disabled, false);
  assert.equal(doc.getElementById('teamActiveName').textContent, 'Совет экспертов');
  assert.equal(doc.getElementById('teamActiveRoles').textContent, 'Советник, Исследователь, Критик');
  assert.match(doc.getElementById('teamRoles').innerHTML, /type="checkbox"/);
  assert.equal(doc.getElementById('teamBtn').getAttribute('aria-expanded'), 'true');
  assert.equal(requestCount(), 1);
});

test('opening while catalog loads shows loading only until the shared request resolves', async () => {
  let resolveCatalog;
  const loading = new Promise(resolve => { resolveCatalog = resolve; });
  const {ui, doc, requestCount} = controllerHarness(() => loading);
  const opened = ui.open();
  assert.equal(doc.getElementById('teamLoading').hidden, false);
  assert.equal(doc.getElementById('teamLayout').hidden, true);
  assert.equal(doc.getElementById('teamApply').disabled, true);
  assert.equal(requestCount(), 1);
  resolveCatalog(smallCatalog);
  await opened;
  assert.equal(doc.getElementById('teamLoading').hidden, true);
  assert.equal(doc.getElementById('teamLayout').hidden, false);
  assert.equal(doc.getElementById('teamApply').disabled, false);
  assert.equal(requestCount(), 1);
});
