'use strict';
const assert = require('node:assert/strict');
const {test} = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function chart(bounds, periods) {
  const context = vm.createContext({document:{addEventListener(){}}, window:{addEventListener(){}}, console});
  for (const file of ['mingkh-app.js','mingkh-trend.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname,'../static',file),'utf8'), context);
  }
  const dates = [];
  const rows = Object.fromEntries(Object.entries(periods).map(([period, values])=>[
    period, values.map(([day, source=0])=>{
      let code = dates.indexOf(day);
      if (code < 0) { code = dates.length; dates.push(day); }
      return [0,source,0,0,0,0,0,code,0];
    })
  ]));
  context.data = {bounds,rows,dims:{date:dates}};
  return context;
}
const bounds = {curr:['2026-09-01','2026-09-03'],prev:['2026-08-29','2026-08-31'],appg:['2025-09-01','2025-09-03']};
const plain = value => JSON.parse(JSON.stringify(value));

test('year comparison fills empty days and preserves period totals',()=>{
  const c = chart(bounds,{curr:[['2026-09-01'],['2026-09-01'],['2026-09-03']],appg:[['2025-09-02']],prev:[]});
  const s = plain(c.trendSeries());
  assert.deepEqual(s.keys,['2026-09-01','2026-09-02','2026-09-03']);
  assert.deepEqual(s.cur,{'2026-09-01':2,'2026-09-03':1});
  assert.deepEqual(s.cmp,{'2026-09-02':1});
  assert.equal(s.curTotal,3);assert.equal(s.cmpTotal,1);
});
test('previous period aligns by its start and respects active filters',()=>{
  const c = chart(bounds,{curr:[['2026-09-01',1],['2026-09-02',0]],prev:[['2026-08-30',1],['2026-08-30',0]],appg:[]});
  c.trendCompare='prev'; c.sel.source=new Set([1]);
  const s = plain(c.trendSeries());
  assert.deepEqual(s.cur,{'2026-09-01':1});
  assert.deepEqual(s.cmp,{'2026-09-02':1});
  assert.equal(s.curTotal,1);assert.equal(s.cmpTotal,1);
});
test('monthly chart includes zero months and clips exact partial-month bounds',()=>{
  const b = {curr:['2026-07-15','2026-09-20'],prev:['2026-01-01','2026-03-31'],appg:['2025-07-15','2025-09-20']};
  const c = chart(b,{curr:[['2026-07-15'],['2026-09-20']],appg:[['2025-07-14'],['2025-07-15'],['2025-09-21']],prev:[]});
  const s = plain(c.trendSeries());
  assert.equal(s.grain,'month');assert.deepEqual(s.keys,['2026-07','2026-08','2026-09']);
  assert.deepEqual(s.cmp,{'2026-07':1});
  assert.equal(s.outside.appg,2);assert.equal(s.cmpTotal,3);
});
test('missing dates and longer custom comparison do not become fake dates or lower totals',()=>{
  const c = chart({...bounds,prev:['2026-08-29','2026-09-03']},{curr:[[''],['2026-09-02']],prev:[[''],['2026-08-29'],['2026-09-03']],appg:[]});
  c.trendCompare='prev';const s=plain(c.trendSeries());
  assert.deepEqual(s.missing,{curr:1,prev:1});
  assert.deepEqual(s.outside,{curr:0,prev:1});
  assert.equal(s.curTotal,2);assert.equal(s.cmpTotal,3);
  assert.deepEqual(s.cmp,{'2026-09-01':1});
});
test('leap day aligns to February 28 and date steps do not depend on timezone',()=>{
  const c = chart({curr:['2025-02-28','2025-03-01'],prev:['2025-02-26','2025-02-27'],appg:['2024-02-28','2024-03-01']},{curr:[],prev:[],appg:[['2024-02-29'],['2024-03-01']]});
  const s=plain(c.trendSeries());
  assert.deepEqual(s.cmp,{'2025-02-28':1,'2025-03-01':1});
  assert.equal(c.fmtDay(c.parseDay('2026-03-29')+c.DAY),'2026-03-30');
});
test('empty dataset produces a finite zero chart scale',()=>{
  const c=chart(bounds,{curr:[],prev:[],appg:[]});const s=plain(c.trendSeries());
  assert.equal(s.curTotal,0);assert.equal(s.cmpTotal,0);assert.equal(c.niceStep(0),1);
});
