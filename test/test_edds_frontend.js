// Run with: node --test test/test_edds_frontend.js
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../services/edds/dashboard.html'), 'utf8');
const parser = source.slice(source.indexOf('function parseDT('), source.indexOf('function build('));
const {parseDT, numOf} = vm.runInNewContext(parser + ';({parseDT,numOf})', {S: v => String(v ?? '').trim()});

test('all embedded scripts compile', () => {
  for(const match of source.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new vm.Script(match[1]);
});

test('ARM dates accept both year formats and reject impossible dates', () => {
  assert.equal(parseDT('24.09.26 07:30').getTime(), parseDT('24.09.2026 07:30:00').getTime());
  assert.equal(parseDT('24.09.2026 07:30').getFullYear(), 2026);
  for(const value of ['', '31.02.2026 07:00', '24.09.26 24:00', '24.09.2026 07:60'])
    assert.equal(parseDT(value), null);
});

test('missing counters remain unknown rather than becoming zero', () => {
  for(const value of ['', '  ', null, undefined, 'не указано']) assert.equal(numOf(value), null);
  assert.equal(numOf('0'), 0);
  assert.equal(numOf('12,5'), 12.5);
});

test('existing objects becoming placeholder points are saved without new codes', () => {
  const storage = new Map();
  const context = vm.createContext({localStorage: {setItem: (k,v) => storage.set(k,v)},
    isoDay: () => '2026-09-24', objRegLoad: () => [{code: 'VZ-0001', type: 'объект'}]});
  const code = source.slice(source.indexOf('function objRegCommit('), source.indexOf('/* ---------- листы Excel ---------- */'));
  vm.runInContext("let OBJREG = [{code:'VZ-0001',type:'объект'}]; const OBJ_KEY='test'; let OBJREG_VER=0;" +
    "function objRegLoad(){return OBJREG;} function objRegSave(){localStorage.setItem(OBJ_KEY,JSON.stringify(OBJREG));}" + code, context);
  vm.runInContext("objRegCommit({prefix:'VZ',objects:[{code:'VZ-0001',isNew:false,hub:true}]})", context);
  assert.equal(JSON.parse(storage.get('test'))[0].type, 'условная точка округа');
});
