/* Exercise the actual dashboard script offline. Element stand-ins model only
 * the DOM calls used for filters and tables; CSS/layout needs a real browser. */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const script = fs.readFileSync(path.resolve(__dirname,'../static/mingkh-treatment.js'),'utf8');

class El {
  constructor(tag) { this.tag=tag;this.children=[];this.events={};this.dataset={};this.value='';this.text=''; }
  append(...nodes) { this.children.push(...nodes); }
  replaceChildren(...nodes) { this.children=nodes;this.text=''; }
  add(el) { this.append(el); }
  set textContent(val) { this.text=String(val);this.children=[]; }
  get textContent() { return this.text+this.children.map(c=>c.textContent).join(' '); }
  addEventListener(ev, fn) { (this.events[ev] ||= []).push(fn); }
  fire(ev, event={}) { for (const fn of this.events[ev] || []) fn(event); }
  setAttribute(k,v) { this[k]=v; }
  removeAttribute(k) { delete this[k]; }
  appendNew(tag) { const e=new El(tag);this.append(e);return e; }
  insertCell() { return this.appendNew('td'); }
  insertRow() { return this.appendNew('tr'); }
  createTHead() { return this.appendNew('thead'); }
  createTBody() { return this.appendNew('tbody'); }
  querySelectorAll() {return this.children;}
  querySelector(query) {return this.children.find(el=>query.includes('"'+el.dataset.view+'"'));}
}

function harness() {
  const nodes = new Map();
  const get = id => {
    if (id === 'trImport') return null;
    if (!nodes.has(id)) nodes.set(id,new El(id));
    return nodes.get(id);
  };
  const form=get('trFilters'), tabs=get('trTabs');
  form.elements=Object.fromEntries('q omsu obj cur topic sg field start end linked'.split(' ').map(k=>[k,new El(k)]));
  form.elements.field.value='created';
  form.reset=()=> {
    form.elements.linked.checked=true;
    for (const key of ['start','end']) form.elements[key].value=form.elements[key].defaultValue;
  };
  for (const name of ['objects','appeals','rating','stats','map','operators','records','coordinates','edds']) {
    const el=tabs.appendNew('button');el.dataset.view=name;el.textContent=name;el.closest=()=>el;
  }
  const k={n:1,name:'КОС',omsu:'Дубна',place:'Дубна',operator:'МУП Водоканал',okey:'дубна',power:1};
  const c={id:10,omsu:'Соседний округ',cur:'МинЖКХ',sg:'closed',status:'Закрыта',topic:'kos',near:1,
    distance:.1,created:'2026-09-01',changed:'2026-11-01',mk:false};
  const data={from:'2026-01-01',as_of:'2026-11-03',built:'03.11.2026 12:38',kos:[k],topics:{kos:'Работа очистных'},
    complaints:[c,{...c,id:11,sg:'work',distance:5,mk:false},{...c,id:12,sg:'work',distance:5,mk:true}],
    radii:{kos:.5},records:[],payments:[{key:'муп водоканал | дубна',recv:0,accr:99}],edds:[]};
  const ctx={document:{getElementById:get,createElement:name=>new El(name)},Option:class {constructor(t,k){this.textContent=t;this.value=k;}},
    fetch:async ()=>({ok:true,json:async ()=>data}),setTimeout,Intl,window:{}};
  vm.runInNewContext(script,ctx);
  const select=(key,val)=>{form.elements[key].value=val;form.fire('input');};
  const view=name=>tabs.fire('click',{target:tabs.children.find(x=>x.dataset.view===name)});
  return {get,form,select,view};
}

test('original linkage radii, answered date, grouped operator payments and tables use actual filters',async()=>{
  const {get,form,select,view}=harness();
  await new Promise(resolve=>setImmediate(resolve));
  const count=()=>Number(get('trMetrics').children[0].children[0].textContent);
  assert.equal(count(),2); // Close object, and explicit mention within ten kilometres.
  select('omsu','Дубна'); assert.equal(count(),2); // Linked object's municipality takes priority.
  select('omsu',''); form.elements.linked.checked=false;form.fire('input');assert.equal(count(),3);
  select('field','answered');select('start','2026-11-01');assert.equal(count(),1);
  view('appeals');assert.match(get('trTable').textContent,/№10/);
  view('operators');assert.match(get('trTable').textContent,/99/);
  assert.ok(get('trTable').textContent.includes('0'));
  view('coordinates');assert.match(get('trTable').textContent,/Дубна/);
  view('edds');assert.match(get('trTable').textContent,/записей нет/);
});
