const {test} = require('node:test');
const assert = require('node:assert/strict');
const waiting = require('../static/aichat-waiting.js');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');

test('the actual template keeps chat initialized when the optional script is absent or broken', () => {
  const template = readFileSync(join(__dirname,'../templates/aichat.html'), 'utf8');
  const initialization = template.slice(template.indexOf('var waitingFacts='), template.indexOf('var customCmds='));
  assert.ok(initialization.length > 0);
  for (const window of [{}, {NeuronaWaiting:{create(){throw Error('asset unavailable');}}}]) {
    const context = vm.createContext({window, $:()=>({})});
    vm.runInContext(initialization, context);
    assert.doesNotThrow(() => vm.runInContext('waitingFacts.start();waitingFacts.stop();waitingFacts.warm();', context));
  }
});

function harness(fetcher) {
  let now = 0, serial = 0;
  const timers = new Map(), calls = [], listeners = {};
  const nodes = Object.fromEntries(['topic','text','close'].map(key => [key, {
    textContent:'', addEventListener(name, fn) { this[name] = fn; }, removeEventListener(name) { delete this[name]; }
  }]));
  const host = {hidden:true, querySelector:selector => nodes[selector.match(/fact-(\w+)/)[1]]};
  const document = {hidden:false, addEventListener:(name, fn) => listeners[name] = fn,
    removeEventListener:name => delete listeners[name]};
  const env = {document, AbortController,
    setTimeout:(fn, ms) => {const id = ++serial;timers.set(id, {at:now+ms, fn});return id;},
    setInterval:(fn, ms) => {const id = ++serial;timers.set(id, {at:now+ms, fn, every:ms});return id;},
    clearTimeout:id => timers.delete(id), clearInterval:id => timers.delete(id),
    addEventListener:(name, fn) => listeners[name] = fn, removeEventListener:name => delete listeners[name],
    fetch:async (url, options) => {calls.push({url, options});return fetcher ? fetcher(url, options) : {
      ok:true, json:async () => ({items:[{topic:'ИИ', text:'Первый факт'}, {topic:'ЖКХ', text:'Второй факт'}]})
    };}
  };
  async function tick(ms) {
    const end = now + ms;
    while (true) {
      const entry = [...timers].filter(([, value]) => value.at <= end).sort((a,b) => a[1].at-b[1].at)[0];
      if (!entry) break;
      const [id, value] = entry;now = value.at;
      if (value.every) value.at += value.every;else timers.delete(id);
      value.fn();await settle();
    }
    now = end;await settle();
  }
  const controller = waiting.create(env, host);
  return {controller, host, nodes, document, listeners, calls, timers, tick};
}
async function settle() { for (let i=0;i<6;i++) await Promise.resolve(); }

test('facts wait a full minute, rotate and never send the conversation', async () => {
  const h = harness();await h.controller.warm();
  h.controller.start();await settle();
  await h.tick(59999);assert.equal(h.host.hidden, true);
  await h.tick(1);assert.equal(h.host.hidden, false);assert.equal(h.nodes.text.textContent, 'Первый факт');
  await h.tick(30000);assert.equal(h.nodes.text.textContent, 'Второй факт');
  assert.deepEqual(h.calls.map(call => call.url), ['/aichat/api/waiting-facts', '/aichat/api/waiting-facts?warm=false']);
  assert.ok(h.calls.every(call => !call.options.body && !call.options.method));
  h.controller.stop();assert.equal(h.host.hidden, true);assert.equal(h.timers.size, 0);
});

test('a quick answer never displays facts; repeated requests reset the clock', async () => {
  const h = harness();h.controller.start();await h.tick(40000);h.controller.stop();
  await h.tick(100000);assert.equal(h.host.hidden, true);
  h.controller.start();await h.tick(59999);assert.equal(h.host.hidden, true);
  await h.tick(1);assert.equal(h.host.hidden, false);h.controller.destroy();
  assert.equal(h.timers.size, 0);
});

test('dismissal lasts for this request and does not steal focus or rotate hidden text', async () => {
  const h = harness();h.controller.start();await h.tick(60000);
  h.nodes.close.click();await h.tick(120000);assert.equal(h.host.hidden, true);
  assert.equal(h.nodes.text.textContent, 'Первый факт');
  h.controller.start();await h.tick(60000);assert.equal(h.host.hidden, false);h.controller.stop();
});

test('network errors and an empty batch do not produce invented facts', async () => {
  let fail = true;
  const h = harness(async () => {if (fail) throw Error('offline');return {ok:true,json:async () => ({items:[]})};});
  h.controller.start();await h.tick(60000);assert.equal(h.host.hidden, true);
  fail = false;await h.tick(30000);assert.equal(h.host.hidden, true);
  h.controller.stop();assert.equal(h.timers.size, 0);
});

test('late responses after stop are ignored and request timeout is bounded', async () => {
  let resolve;
  const h = harness(() => new Promise(done => {resolve=done;}));
  h.controller.start();await h.tick(8000);
  assert.equal(h.calls[0].options.signal.aborted, true);
  h.controller.stop();resolve({ok:true,json:async () => ({items:[{topic:'ИИ',text:'Поздний факт'}]})});
  await settle();await h.tick(60000);assert.equal(h.host.hidden, true);assert.equal(h.nodes.text.textContent, '');
});

test('background tabs do not rotate; model content is inserted as plain text', async () => {
  const unsafe = '<img src=x onerror=alert(1)>';
  const h = harness(async () => ({ok:true,json:async () => ({items:[{topic:'ИИ',text:unsafe}]})}));
  h.controller.start();h.document.hidden = true;await h.tick(60000);assert.equal(h.host.hidden, true);
  h.document.hidden = false;h.listeners.visibilitychange();assert.equal(h.host.hidden, false);
  assert.equal(h.nodes.text.textContent, unsafe);assert.equal(h.nodes.text.innerHTML, undefined);
  h.listeners.pagehide();assert.equal(h.host.hidden, true);assert.equal(h.timers.size, 0);
});

test('a returned login page or invalid facts stay invisible', async () => {
  for (const response of [{ok:true,redirected:true}, {ok:false}, {ok:true,json:async () => ({items:[
    {topic:'Новость',text:'wrong'}, {topic:'ИИ',text:''}, {topic:'ЖКХ',text:'x'.repeat(321)}
  ]})}]) {
    const h = harness(async () => response);h.controller.start();await h.tick(60000);
    assert.equal(h.host.hidden, true);h.controller.stop();
  }
});
