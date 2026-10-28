'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const source = fs.readFileSync(require.resolve('../static/omsu-rating.js'), 'utf8');

// A deliberately small DOM: assertions exercise user events and visible output.
// innerHTML is forbidden so a future rendering change cannot silently allow HTML from sources.
class Node {
  constructor(tag = 'div') { this.tagName = tag; this.children = []; this.attrs = {}; this.events = {}; this.dataset = {}; this.style = {}; this.value = ''; this.hidden = false; this.checked = false; this._text = ''; }
  set textContent(text) { this._text = String(text); this.children = []; }
  get textContent() { return this._text + this.children.map(child => child.textContent).join(''); }
  set innerHTML(_) { throw Error('Unsafe HTML rendering'); }
  append(...nodes) { nodes.forEach(node => { this.children.push(node); node.parentElement = this; }); }
  replaceChildren(...nodes) { this._text = ''; this.children = []; this.append(...nodes); }
  setAttribute(name, value) { this.attrs[name] = value; }
  addEventListener(name, handler) { this.events[name] = handler; }
  focus() { this.focused = true; }
  showModal() { this.open = true; }
  close() { this.open = false; }
}
const descendants = node => [node, ...node.children.flatMap(descendants)];
const flush = () => new Promise(resolve => setImmediate(resolve));
function row(omsu, place, value, overrides = {}) {
  return {omsu, place, place_label:String(place), value, group:'main', zone:'z2', bar_pct:50,
    no_docs:false, tag:null, previous_place:place, delta:0, place_tooltip:'Пояснение места', value_tooltip:'Формула значения',
    details:[{name:'РСО ' + omsu, value, amount:30, nvv:1000, fact:3, category:1, category_label:'НВВ < 300 млн руб.', plan:4}],
    rso_names:['РСО ' + omsu], extra:{fact:3,plan:[4],fact_tooltip:'Формула факта',plan_tooltip:'План по категории'}, ...overrides};
}
function state() {
  const rows = [row('АЛЬФА',1,120,{zone:'z1',place_label:'1–2',tag:'achieved',previous_place:3,delta:2,bar_pct:100}),
    row('БЕТА',1,110,{zone:'z1',place_label:'1–2',tag:'achieved',previous_place:null,delta:null}),
    row('ГАММА',3,40,{zone:'z2',tag:'not_achieved',previous_place:1,delta:-2}),
    row('ДЕЛЬТА',4,0,{zone:'z3',tag:'not_achieved'}), row('ВЛАСИХА',1,0,{group:'zato',zone:'z3',tag:'not_achieved'})];
  const indicators = {};
  for (const id of [4,5,6,7]) indicators[id] = {
    name:'Показатель ' + id, short:'Критерий ' + id, roman:{4:'IV',5:'V',6:'VI',7:'VII'}[id], metric:id === 5 ? 'Балл' : '%',
    note:id === 6 ? 'Условные места' : id === 7 ? 'Предварительная оценка' : 'Методика', better:id === 6 ? 'asc' : 'desc',
    periods:[[],[],[],[]],period_stats:[{},{},{},{}],default_period:null,error:null,
  };
  indicators[4].periods[0] = rows;
  indicators[4].periods[1] = rows;
  indicators[4].default_period = 1;
  indicators[4].period_stats[1] = {count:5,zato_count:1,average:54,achieved_count:2,has_previous:true,nonparticipants:[{omsu:'ЕГОРЬЕВСК',group:'main',na:true},{omsu:'ВОСХОД',group:'zato',na:true}]};
  indicators[4].period_stats[0] = {...indicators[4].period_stats[1],has_previous:false};
  return {refresh:{status:'idle'},history:[],municipalities:['АЛЬФА','БЕТА','ГАММА','ДЕЛЬТА','ВЛАСИХА','ЕГОРЬЕВСК','ВОСХОД'],
    period_labels:['1 квартал','1 полугодие','9 месяцев','11 месяцев'],indicators};
}
async function app(data, {preference={indicator:4,period:null,showAll:false}, offline=false, fail=false} = {}) {
  const nodes = new Map();
  const get = id => { if (!nodes.has(id)) nodes.set(id,new Node()); return nodes.get(id); };
  if (offline) get('or-snapshot-data').textContent = JSON.stringify(data);
  const storage = new Map(preference ? [['neurona:omsu-rating:v1',JSON.stringify(preference)]] : []);
  const calls = [], windowEvents = {}, timers = [], urls = [];
  const document = {
    getElementById:id => id === 'or-snapshot-data' && !offline ? null : get(id),
    documentElement:new Node('html'),createElement:tag => new Node(tag),addEventListener(){},
    querySelector:selector => {
      const match = selector.match(/^\[data-(sort|indicator)="([\w]+)"\]$/);
      return match ? [...nodes.values()].flatMap(descendants).find(node=>node.dataset[match[1]] === match[2]) : null;
    },
  };
  const context = {document,window:{addEventListener:(event,handler)=>{windowEvents[event]=handler;}},AbortController,
    setTimeout:handler=>{timers.push(handler);return timers.length;},clearTimeout(){},Date,Map,Set,Uint8Array,Blob,
    atob:value=>Buffer.from(value,'base64').toString('binary'),URL:class extends URL {static createObjectURL(blob){urls.push(blob);return 'blob:offline-'+urls.length;}static revokeObjectURL(){}},
    localStorage:{getItem:key=>storage.get(key),setItem:(key,value)=>storage.set(key,value)},
    fetch:async(path,options)=>{calls.push({path,options});if(fail)throw Error('Сеть недоступна');return {ok:true,json:async()=>data};}};
  vm.runInNewContext(source,context);
  await flush();
  const event = async(id,type,value) => {if (value !== undefined) get(id).value=value;await get(id).events[type]?.({target:get(id)});await flush();};
  const click = async(node) => {assert.ok(node,'control exists');await node.events.click?.();await flush();};
  return {get,event,click,document,storage,calls,timers,urls,windowEvents,
    displayed:()=>get('or-rows').children.filter(node=>node.dataset.omsu),
    sort:key=>document.querySelector('[data-sort="'+key+'"]'),tab:id=>document.querySelector('[data-indicator="'+id+'"]')};
}
function cell(tr,key) { return tr.children.find(node=>node.dataset.column===key); }

test('original colors, shared places, bars, source columns, flags and stats survive a municipality filter',async()=>{
  const ui=await app(state());
  assert.equal(cell(ui.displayed()[0],'place').textContent,'1–2');
  assert.match(cell(ui.displayed()[0],'omsu').textContent,/План выполнен/);
  assert.equal(cell(ui.displayed()[0],'value').children[1].children[0].style.width,'100%');
  assert.ok(ui.sort('extra'));assert.ok(ui.sort('plan'));assert.ok(ui.sort('rso'));assert.ok(ui.sort('delta'));
  await ui.event('or-search','input','гамма; власиха');
  assert.deepEqual(ui.displayed().map(r=>r.dataset.omsu),['ГАММА','ВЛАСИХА']);
  assert.match(cell(ui.displayed()[0],'omsu').children[0].className,/z2/);
  assert.match(cell(ui.displayed()[1],'omsu').children[0].className,/z3/);
  assert.match(ui.get('or-summary').textContent,/54,00/); // full period average, not the filtered zero/40 average
  assert.equal(cell(ui.displayed()[0],'place').textContent,'3');
});

test('all original sortable columns keep ZATO separate, null dynamics last and nonparticipants unranked',async()=>{
  const ui=await app(state());
  await ui.click(ui.sort('delta'));
  assert.deepEqual(ui.displayed().map(r=>r.dataset.omsu),['АЛЬФА','ДЕЛЬТА','ГАММА','БЕТА','ВЛАСИХА']);
  await ui.click(ui.sort('delta'));
  assert.deepEqual(ui.displayed().map(r=>r.dataset.omsu),['ГАММА','ДЕЛЬТА','АЛЬФА','БЕТА','ВЛАСИХА']);
  ui.get('or-all').checked=true;await ui.event('or-all','change');
  assert.deepEqual(ui.displayed().map(r=>r.dataset.omsu),['ГАММА','ДЕЛЬТА','АЛЬФА','БЕТА','ЕГОРЬЕВСК','ВЛАСИХА','ВОСХОД']);
  assert.equal(cell(ui.displayed()[4],'place').textContent,'—');
  assert.match(cell(ui.displayed()[4],'omsu').textContent,/Не участвует/);
  assert.equal(ui.displayed()[4].children.at(-1).textContent,'');
  assert.equal(JSON.parse(ui.storage.get('neurona:omsu-rating:v1')).showAll,true);
});

test('saved period/tab selection is validated, empty periods disabled and dynamics absent without comparison',async()=>{
  const ui=await app(state(),{preference:{indicator:4,period:0,showAll:false}});
  assert.equal(ui.sort('delta'),undefined);
  assert.equal(ui.get('or-period').children[0].selected,true);
  assert.equal(ui.get('or-period').children[2].disabled,true);
  assert.equal(ui.get('or-export').href,'/omsu-rating/export.xlsx?ind=4&period=0');
  await ui.event('or-period','change','1');
  assert.ok(ui.sort('delta'));
  assert.equal(JSON.parse(ui.storage.get('neurona:omsu-rating:v1')).period,1);
  const fallback=await app(state(),{preference:{indicator:4,period:3}});
  assert.equal(fallback.get('or-period').children[1].selected,true);
  await fallback.click(fallback.tab(6));
  assert.match(fallback.get('or-rows').textContent,/Пока нет данных/);
  assert.match(fallback.get('or-methodology').textContent,/Условные/);
  assert.ok(fallback.sort('lsno'));assert.ok(fallback.sort('lstot'));assert.equal(fallback.sort('rso'),undefined);
});

test('K7 detail keeps zero, deadlines, reasons and missing/duplicate source evidence as safe text',async()=>{
  const data=state(), raw='<img src=x onerror=alert(1)>';
  const r=row('АЛЬФА',1,0,{extra:{done:0,plan:17},details:[{name:raw,value:0,done:0,plan:17,found:true,duplicates:2,
    sheet_name:raw,sheet_omsu:'АЛЬФА',items:[{n:1,name:'Мероприятие',due:'27.02',raw:'0',ok:false,reason:'предыдущий шаг не выполнен'},
      {n:2,name:'Документы',due:'06.03',raw,ok:false,reason:'неясная запись'}]}]});
  data.indicators[7].periods[2]=[r];data.indicators[7].default_period=2;
  data.indicators[7].period_stats[2]={count:1,average:0,zato_count:0,has_previous:false,nonparticipants:[]};
  const ui=await app(data,{preference:{indicator:7,period:2}});
  assert.ok(ui.sort('extra'));
  assert.equal(cell(ui.displayed()[0],'extra').textContent,'0 / 17');
  await ui.click(ui.displayed()[0].children.at(-1).children[0]);
  assert.equal(ui.get('or-detail').open,true);
  const text=ui.get('or-detail-body').textContent;
  assert.match(text,/В источнике: 0/);assert.match(text,/27\.02/);assert.match(text,/Строк РСО в источнике: 2/);
  assert.ok(text.includes(raw));
  assert.equal(descendants(ui.get('or-detail-body')).some(n=>n.tagName==='img'),false);
  await ui.click(ui.get('or-detail-close'));assert.equal(ui.get('or-detail').open,false);
});

test('job-level failure remains visible when last-good source rows have no error',async()=>{
  const data=state();data.refresh={status:'error',error:'Обновление было прервано. Предыдущие данные сохранены.'};
  const ui=await app(data);
  assert.equal(ui.get('or-error').hidden,false);assert.equal(ui.get('or-error').textContent,data.refresh.error);
  assert.equal(ui.get('or-source-error').hidden,true);assert.equal(ui.get('or-refresh').disabled,false);
  assert.equal(ui.displayed().length,5);
});

test('a network failure keeps retry available and never manufactures a rating',async()=>{
  const ui=await app(state(),{fail:true});
  assert.match(ui.get('or-error').textContent,/Сеть недоступна/);
  assert.equal(ui.get('or-refresh').disabled,false);assert.equal(ui.displayed().length,0);
});

test('offline snapshot has all interactions and Excel while making no requests or polling',async()=>{
  const data=state();data.offline_exports={'4:1':Buffer.from('xlsx selected H1').toString('base64'),'4:0':Buffer.from('xlsx selected Q1').toString('base64')};
  const ui=await app(data,{offline:true});
  assert.equal(ui.calls.length,0);assert.equal(ui.timers.length,0);
  assert.equal(ui.get('or-refresh').hidden,true);assert.equal(ui.get('or-home').hidden,true);
  assert.equal(ui.get('or-export').hidden,false);assert.match(ui.get('or-export').href,/^blob:/);
  assert.match(ui.get('or-history').textContent,/Автономный снимок/);
  await ui.event('or-search','input','альфа');assert.equal(ui.displayed().length,1);
  await ui.event('or-period','change','0');
  assert.equal(await ui.urls.at(-1).text(),'xlsx selected Q1');
  assert.equal(ui.calls.length,0);assert.equal(ui.timers.length,0);
});


test('source links use fixed HTTPS hosts, never request the source automatically and reject unsafe addresses',async()=>{
  const data=state();data.indicators[4].source_url='https://mgkh.rm.mosreg.ru/issues?query_id=53';
  const ui=await app(data);
  assert.equal(ui.get('or-source-link').hidden,false);
  assert.equal(ui.get('or-source-link').target,'_blank');
  assert.equal(ui.get('or-source-link').rel,'noopener noreferrer');
  assert.ok(ui.calls.every(call=>call.path.startsWith('/omsu-rating/api/')));
  data.indicators[4].source_url='javascript:alert(1)';
  const unsafe=await app(data);assert.equal(unsafe.get('or-source-link').hidden,true);
  data.indicators[4].source_url='https://mgkh.rm.mosreg.ru.attacker.example/issues';
  const foreign=await app(data);assert.equal(foreign.get('or-source-link').hidden,true);
  data.indicators[4].source_url='https://docs.google.com/spreadsheets/d/fixed-sheet/edit';
  const offline=await app(data,{offline:true});
  assert.equal(offline.get('or-source-link').hidden,false);assert.equal(offline.calls.length,0);
});


test('a fresh profile opens indicator V like the original application, while an explicit saved tab wins',async()=>{
  const data=state();
  data.indicators[5].periods[1]=data.indicators[4].periods[1];
  data.indicators[5].period_stats[1]={...data.indicators[4].period_stats[1],achieved_count:null};
  data.indicators[5].default_period=1;
  const fresh=await app(data,{preference:null});
  assert.equal(fresh.tab(5).attrs['aria-pressed'],'true');
  assert.match(fresh.get('or-title').textContent,/^V\./);
  assert.equal(fresh.get('or-export').href,'/omsu-rating/export.xlsx?ind=5&period=1');
  const returning=await app(data,{preference:{indicator:4,period:0}});
  assert.equal(returning.tab(4).attrs['aria-pressed'],'true');
});
