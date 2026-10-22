'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { LinkError, createRequest, LinkPage, expiresAt } = require('../static/max-bot-link.js');

class Element {
  constructor(tag = 'div') { this.tagName = tag; this.children = []; this.listeners = {}; this.hidden = false; this.disabled = false; this._text = ''; this.className = ''; this.classList = { toggle() {} }; }
  set textContent(value) { this._text = String(value); this.children = []; }
  get textContent() { return this._text + this.children.map(child => child.textContent).join(''); }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this._text = ''; this.children = children; }
  addEventListener(name, callback) { this.listeners[name] = callback; }
  querySelectorAll(tag) { return this.children.flatMap(child => [...(child.tagName === tag ? [child] : []), ...child.querySelectorAll(tag)]); }
}
const html = fs.readFileSync(path.join(__dirname, '../templates/max_bot.html'), 'utf8');
function environment() {
  const nodes = new Map([...html.matchAll(/id="(max-bot-[^"]+)"/g)].map(match => [match[1], new Element()]));
  const listeners = {}; const documentListeners = {}; const requests = [];
  const win = {
    document: { hidden: false, querySelector: () => ({ content: 'test-csrf' }), getElementById: id => { if (!nodes.has(id)) throw new Error('Missing ' + id); return nodes.get(id); }, createElement: tag => new Element(tag), addEventListener: (name, callback) => { documentListeners[name] = callback; } },
    navigator: {}, fetch: async (url, options) => { requests.push({ url, options }); return { ok: true, json: async () => ({ enabled: true, active: false, pending: [] }) }; },
    addEventListener: (name, callback) => { listeners[name] = callback; }, setInterval: () => 1, clearInterval() {},
  };
  return { win, nodes, requests, listeners, documentListeners, page: new LinkPage(win) };
}
const fresh = () => new Date(Date.now() + 60000).toISOString();
const pending = () => ({ id: 'request-1', label: 'Аккаунт MAX ••1234', expires_at: fresh() });
const status = (extra = {}) => ({ enabled: true, active: false, pending: [], ...extra });
const flush = () => new Promise(resolve => setImmediate(resolve));

test('request uses only fixed same-origin routes, POST CSRF and no-store credentials', async () => {
  const calls = []; const win = { fetch: async (url, options) => { calls.push({ url, options }); return { ok: true, json: async () => ({}) }; } };
  const request = createRequest(win, 'test-csrf');
  await request('link'); await request('confirm', { request_id: 'req' }); await request('status');
  assert.equal(calls[0].url, '/max-bot/api/link'); assert.equal(calls[0].options.headers['X-CSRF-Token'], 'test-csrf');
  assert.equal(calls[0].options.credentials, 'same-origin'); assert.equal(calls[0].options.cache, 'no-store'); assert.equal(calls[0].options.redirect, 'error');
  assert.equal(calls[1].options.body, '{"request_id":"req"}'); assert.equal(calls[2].options.method, 'GET'); assert.equal(calls[2].options.body, undefined);
  await assert.rejects(request('../admin'), /INVALID_OPERATION/); assert.equal(calls.length, 3);
});
test('missing CSRF cannot send mutations', async () => {
  let called = false; const request = createRequest({ fetch: async () => { called = true; } }, '');
  await assert.rejects(request('link'), error => error.status === 403); assert.equal(called, false);
});
test('network failure and invalid response are errors, not disconnected-success', async () => {
  await assert.rejects(createRequest({ fetch: async () => { throw new Error('connection'); } }, 'csrf')('status'), /UNAVAILABLE/);
  await assert.rejects(createRequest({ fetch: async () => ({ ok: true, status: 200, json: async () => { throw new Error('html'); } }) }, 'csrf')('status'), /UNAVAILABLE/);
});
test('timeout aborts the request', async () => {
  const request = createRequest({ fetch: (url, options) => new Promise((resolve, reject) => options.signal.addEventListener('abort', () => reject(new Error('aborted')))) }, 'csrf', 5);
  await assert.rejects(request('status'), /TIMEOUT/);
});
test('pending request is rendered as safe text, never confirmed automatically', () => {
  const { page } = environment(); let calls = 0; page.request = async () => { calls++; };
  page.applyStatus(status({ pending: [{ ...pending(), label: '<img src=x onerror=alert(1)>' }] }));
  assert.match(page.el('pending-list').textContent, /<img src=x onerror=alert\(1\)>/); assert.equal(calls, 0);
  assert.equal(page.el('pending-list').querySelectorAll('img').length, 0);
  assert.equal(page.el('pending-list').querySelectorAll('button').length, 1);
});
test('confirmation requires a current displayed request and explicit action', async () => {
  const { page } = environment(); const item = pending(); page.applyStatus(status({ pending: [item] })); const calls = [];
  page.request = async (operation, body) => { calls.push({ operation, body }); return status({ active: true }); };
  await page.confirm('other-id'); assert.equal(calls.length, 0);
  await page.confirm(item.id); assert.deepEqual(calls[0], { operation: 'confirm', body: { request_id: item.id } }); assert.equal(calls[1].operation, 'status');
  assert.equal(page.state.active, true); assert.equal(page.el('pending').hidden, true);
});
test('double confirmation cannot send duplicate concurrent requests', async () => {
  const { page } = environment(); page.applyStatus(status({ pending: [pending()] })); let finish; let confirms = 0;
  page.request = async operation => { if (operation === 'confirm') { confirms++; await new Promise(resolve => { finish = resolve; }); } return status({ active: true }); };
  const first = page.confirm('request-1'); await page.confirm('request-1'); assert.equal(confirms, 1); finish(); await first;
});
test('one-time code is shown in memory, cleared on expiry, and never stored', async () => {
  const { page, win } = environment(); Object.defineProperty(win, 'localStorage', { get() { throw new Error('storage forbidden'); } });
  page.applyStatus(status()); page.request = async () => ({ code: 'ABCD-1234', expires_at: fresh() }); await page.createCode();
  assert.equal(page.el('code').textContent, '/link ABCD-1234'); assert.equal(page.state.waiting, true); assert.equal(page.el('code-panel').hidden, false);
  page.state.expires = Date.now() - 1; page.tick(); assert.equal(page.state.code, ''); assert.equal(page.el('code').textContent, ''); assert.equal(page.el('code-panel').hidden, true); assert.equal(page.el('expired').hidden, false);
});
test('invalid or expired server code is never displayed', async () => {
  for (const code of ['<script>bad</script>', 'x']) {
    const { page } = environment(); page.applyStatus(status()); page.request = async () => ({ code, expires_at: fresh() }); await page.createCode();
    assert.equal(page.state.code, ''); assert.equal(page.el('error').hidden, false);
  }
});
test('pending claim clears the reusable display of the one-time code', async () => {
  const { page } = environment(); page.applyStatus(status()); page.request = async () => ({ code: 'ABCD-1234', expires_at: fresh() }); await page.createCode();
  page.applyStatus(status({ pending: [pending()] })); assert.equal(page.state.code, ''); assert.equal(page.el('code').textContent, ''); assert.equal(page.state.waiting, true);
});
test('polling occurs only while waiting, visible and at least ten seconds apart', async () => {
  const { page, win } = environment(); let calls = 0; page.request = async () => { calls++; return status({ pending: [pending()] }); };
  page.applyStatus(status()); page.tick(); assert.equal(calls, 0);
  page.applyStatus(status({ pending: [pending()] })); page.lastPoll = Date.now() - 10001; win.document.hidden = true; page.tick(); assert.equal(calls, 0);
  win.document.hidden = false; page.tick(); await flush(); assert.equal(calls, 1);
  page.tick(); await flush(); assert.equal(calls, 1);
  page.applyStatus(status({ active: true })); page.lastPoll = Date.now() - 10001; page.tick(); assert.equal(calls, 1);
});
test('expired requests cannot be confirmed and disappear', async () => {
  const { page } = environment(); page.applyStatus(status({ pending: [{ ...pending(), expires_at: new Date(Date.now() - 1).toISOString() }] }));
  assert.equal(page.state.pending.length, 0); let calls = 0; page.request = async () => { calls++; }; await page.confirm('request-1'); assert.equal(calls, 0);
});
test('late status before a code mutation cannot overwrite new state', async () => {
  const { page } = environment(); page.applyStatus(status()); let finish;
  page.request = async operation => operation === 'status' ? new Promise(resolve => { finish = resolve; }) : { code: 'ABCD-1234', expires_at: fresh() };
  const old = page.loadStatus(); await page.createCode(); finish(status({ active: true })); await old;
  assert.equal(page.state.active, false); assert.equal(page.state.code, 'ABCD-1234');
});
test('session or permissions failure clears code and pending requests and blocks actions', async () => {
  const { page } = environment(); page.applyStatus(status()); page.state.code = 'ABCD-1234'; page.state.pending = [pending()];
  page.fail(new LinkError('HTTP_ERROR', 403)); assert.equal(page.state.code, ''); assert.equal(page.state.pending.length, 0); assert.equal(page.state.blocked, true);
  let calls = 0; page.request = async () => { calls++; }; await page.createCode(); await page.loadStatus(); assert.equal(calls, 0);
});
test('unlink waits for server status and never changes business data', async () => {
  const { page } = environment(); page.applyStatus(status({ active: true })); const calls = [];
  page.request = async operation => { calls.push(operation); return status(); }; await page.unlink();
  assert.deepEqual(calls, ['unlink', 'status']); assert.equal(page.state.active, false); assert.equal(page.el('active').hidden, true);
});
test('failed unlink keeps active status and exposes retry check', async () => {
  const { page } = environment(); page.applyStatus(status({ active: true })); page.request = async () => { throw new LinkError('UNAVAILABLE'); }; await page.unlink();
  assert.equal(page.state.active, true); assert.equal(page.el('error').hidden, false); assert.equal(page.el('retry').hidden, false);
});
test('copy requires explicit click and has fallback when Clipboard API is absent', async () => {
  const { page } = environment(); page.applyStatus(status()); page.request = async () => ({ code: 'ABCD-1234', expires_at: fresh() }); await page.createCode();
  await page.copyCode(); assert.equal(page.el('copy').textContent, 'Выделите команду'); assert.match(page.el('announcement').textContent, /вручную/);
});
test('page hide clears secret and a delayed response cannot restore it', async () => {
  const { page, listeners } = environment(); page.init(); await flush(); let finish;
  page.request = async () => new Promise(resolve => { finish = resolve; }); const request = page.createCode();
  listeners.pagehide(); finish({ code: 'ABCD-1234', expires_at: fresh() }); await request;
  assert.equal(page.state.code, ''); assert.equal(page.el('code').textContent, ''); assert.equal(page.busy, false);
});
test('template contains no password/token fields, storage or MiniApp bridge', () => {
  const js = fs.readFileSync(path.join(__dirname, '../static/max-bot-link.js'), 'utf8');
  assert.doesNotMatch(html, /type=["']password|initData|WebApp|st\.max\.ru|requestContact/);
  assert.doesNotMatch(js, /localStorage|sessionStorage|innerHTML|initData|console\./);
  assert.match(html, /max_bot_csrf/); assert.match(html, /\/zips/);
  assert.equal(expiresAt('invalid'), 0);
});

test('HTTP LAN policy failure explains HTTPS requirement without bypass instructions', async () => {
  const { page } = environment();
  page.request = createRequest({ fetch: async () => ({ ok: false, status: 403, json: async () => ({ detail: { code: 'MAX_BOT_HTTPS_REQUIRED' } }) }) }, 'csrf');
  await page.loadStatus();
  assert.match(page.el('error').textContent, /по HTTPS или локально/);
  assert.equal(page.state.blocked, true);
});
