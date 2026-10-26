'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const source = fs.readFileSync(require.resolve('../static/omsu-rating.js'), 'utf8');

test('job-level failure remains visible when last-good source rows have no error', async () => {
  const nodes = new Map();
  const make = () => ({textContent: '', hidden: false, value: '', dataset: {}, append() {}, replaceChildren() {}, setAttribute() {}, addEventListener() {}, parentElement: {setAttribute() {}}});
  const get = id => { if (!nodes.has(id)) nodes.set(id, make()); return nodes.get(id); };
  const state = {refresh: {status:'error',error:'Обновление было прервано. Предыдущие данные сохранены.'},history:[],municipalities:[],period_labels:['1 квартал'],indicators:{4:{name:'Показатель',note:'Методика',metric:'%',short:'Тест',periods:[[]],default_period:null,error:null}}};
  const document = {getElementById:get,querySelectorAll:()=>[],createElement:make,addEventListener(){}};
  const context = {document,window:{addEventListener(){}},AbortController,setTimeout,clearTimeout,Date,Map,Set,fetch:async()=>({ok:true,json:async()=>state})};
  vm.runInNewContext(source,context);
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(get('or-error').hidden,false);
  assert.equal(get('or-error').textContent,state.refresh.error);
  assert.equal(get('or-source-error').hidden,true);
  assert.equal(get('or-refresh').disabled,false);
});
