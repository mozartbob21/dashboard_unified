const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const {ReleaseAPI, ReadSession, safeToShow} = require('../static/release-notes.js');

test('release requests stay same-origin and never cache private grants', async () => {
  const calls = [];
  const api = new ReleaseAPI(async (url, options) => {
    calls.push({url, options});
    return {ok:true, json:async () => ({releases:[], ok:true})};
  });
  assert.deepEqual(await api.list(), []);
  await api.list(true); await api.read('release/id');
  assert.deepEqual(calls.map(x => x.url), ['/api/releases/pending','/api/releases/history','/api/releases/release%2Fid/read']);
  assert.equal(calls[2].options.method, 'POST');
  assert.equal(calls[2].options.keepalive, true);
  for (const call of calls) { assert.equal(call.options.cache, 'no-store'); assert.equal(call.options.credentials, 'same-origin'); assert.equal(call.options.body, undefined); }
});

test('login redirect and malformed responses cannot claim completion', async () => {
  for (const response of [{ok:true,redirected:true}, {ok:false}, {ok:true,json:async()=>({})}]) {
    const api = new ReleaseAPI(async () => response);
    await assert.rejects(api.list()); await assert.rejects(api.read('r'));
  }
});

test('double acknowledgement shares a single request; failures retry only unsaved releases', async () => {
  let release, fail = true;
  const wait = new Promise(resolve => release = resolve);
  const calls = [];
  const session = new ReadSession({read:async id => { calls.push(id); await wait; if(id==='b' && fail) throw Error('offline'); }}, [{id:'a'},{id:'b'}]);
  const first = session.acknowledge();
  assert.equal(first, session.acknowledge()); release();
  await assert.rejects(first);
  assert.deepEqual([...session.saved], ['a']);
  fail = false; await session.acknowledge();
  assert.deepEqual(calls, ['a','b','b']);
  await session.acknowledge(); assert.equal(calls.length, 3);
});

test('automatic modal waits for idle, visible, unedited page without active form or other modal', () => {
  global.getComputedStyle = () => ({visibility:'visible'});
  const doc = {visibilityState:'visible',hasFocus:()=>true,activeElement:{closest:()=>null},querySelectorAll:()=>[]};
  assert.equal(safeToShow(doc,false,1000,6000),true);
  assert.equal(safeToShow(doc,false,1000,5999),false);
  assert.equal(safeToShow(doc,true,1000,9999),false);
  assert.equal(safeToShow({...doc,visibilityState:'hidden'},false,1000,9999),false);
  assert.equal(safeToShow({...doc,activeElement:{closest:()=>({})}},false,1000,9999),false);
  assert.equal(safeToShow({...doc,querySelectorAll:()=>[{getClientRects:()=>[1]}]},false,1000,9999),false);
});

test('frontend compiles, has no HTML interpolation sink, and is wired to authenticated shells', () => {
  const script = fs.readFileSync(require.resolve('../static/release-notes.js'),'utf8');
  new vm.Script(script);
  assert.doesNotMatch(script, /\.innerHTML\s*=|insertAdjacentHTML|localStorage|sessionStorage/);
  for (const name of ['home.html','base_module.html','zip_curator_original.html']) {
    const template = fs.readFileSync(require.resolve('../templates/'+name),'utf8');
    assert.match(template, /release-notes\.js/); assert.match(template, /release-notes\.css/);
  }
  const home = fs.readFileSync(require.resolve('../templates/home.html'),'utf8');
  assert.match(home, /data-release-notes-open/); assert.doesNotMatch(home, /id="whatsNewModal"/);
});
