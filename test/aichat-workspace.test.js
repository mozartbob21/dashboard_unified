// Run with node --test test/aichat-workspace.test.js. No browser packages needed.
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {readFileSync} = require('node:fs');
const vm = require('node:vm');
const workspace = require('../static/aichat-workspace.js');
const formatterContext = vm.createContext({});
vm.runInContext(readFileSync(require.resolve('../static/aichat-message-format.js'),'utf8'),formatterContext);
const format = formatterContext.NeuronaChatFormat.render;

class Events {
  constructor() { this.events = new Map(); }
  addEventListener(name, fn) { if (!this.events.has(name)) this.events.set(name,new Set());this.events.get(name).add(fn); }
  removeEventListener(name, fn) { this.events.get(name)?.delete(fn); }
  dispatch(name, values = {}) {
    const event = {target:this, defaultPrevented:false, preventDefault() { this.defaultPrevented = true; }, ...values};
    for (const fn of this.events.get(name) || []) fn(event);
    return event;
  }
}
class Element extends Events {
  constructor(doc, tag = 'div') {
    super();this.ownerDocument = doc;this.tagName = tag.toUpperCase();this.children = [];this.attributes = {};
    this.parentElement = null;this._text = '';this._html = '';this.disabled = false;this.style = {setProperty:(key,value) => { this.style[key] = value; }};
    this.classList = {contains:name => this.className.split(/\s+/).includes(name), toggle:(name,force) => {
      const classes = new Set(this.className.split(/\s+/).filter(Boolean));
      const next = force === undefined ? !classes.has(name) : force;
      if (next) classes.add(name);else classes.delete(name);
      this.className = Array.from(classes).join(' ');return next;
    }};
  }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  getAttribute(name) { return this.attributes[name] ?? null; }
  removeAttribute(name) { delete this.attributes[name]; }
  get className() { return this.getAttribute('class') || ''; }
  set className(value) { this.setAttribute('class',value); }
  get id() { return this.getAttribute('id') || ''; }
  set id(value) { this.setAttribute('id',value); }
  get hidden() { return this.getAttribute('hidden') !== null; }
  set hidden(value) {
    const previous = this.hidden;
    if (value) this.setAttribute('hidden','');else this.removeAttribute('hidden');
    if (previous !== !!value) this.ownerDocument?.observers.forEach(fn => fn());
  }
  get inert() { return this.getAttribute('inert') !== null; }
  set inert(value) { if (value) this.setAttribute('inert','');else this.removeAttribute('inert'); }
  get isConnected() { return !!this.ownerDocument?.body?.contains(this); }
  get textContent() { return this._text + this.children.map(child => child.textContent).join(''); }
  set textContent(value) { this.replaceChildren();this._text = String(value); }
  get innerHTML() { return this._html; }
  set innerHTML(value) {
    this.replaceChildren();this._html = String(value);
    this._text = this._html.replace(/<[^>]*>/g,'').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
  }
  get innerText() { return this.textContent; }
  appendChild(child) { child.remove();this.children.push(child);child.parentElement = this;return child; }
  replaceChildren(...children) {
    this.children.forEach(child => { child.parentElement = null; });this.children = [];this._text = '';this._html = '';
    children.forEach(child => this.appendChild(child));
  }
  remove() { if (this.parentElement) this.parentElement.children = this.parentElement.children.filter(child => child !== this);this.parentElement = null; }
  contains(child) { return child === this || this.children.some(item => item.contains(child)); }
  matches(selector) {
    return selector.split(',').some(raw => {
      const item = raw.trim();
      if (item.startsWith('#')) return this.id === item.slice(1);
      if (item.startsWith('.')) return this.classList.contains(item.slice(1));
      const attr = /^([a-z]+)?\[([\w-]+)(?:="([^"]*)")?\]$/.exec(item);
      if (attr) return (!attr[1] || this.tagName === attr[1].toUpperCase()) && this.getAttribute(attr[2]) !== null &&
        (attr[3] === undefined || this.getAttribute(attr[2]) === attr[3]);
      return this.tagName === item.toUpperCase();
    });
  }
  closest(selector) { return this.matches(selector) ? this : this.parentElement?.closest(selector) || null; }
  querySelectorAll(selector) {
    return this.children.flatMap(child => [...(child.matches(selector) ? [child] : []), ...child.querySelectorAll(selector)]);
  }
  querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
  focus() { if (this.closest('[hidden], [inert]')) return;this.ownerDocument.activeElement = this;this.ownerDocument.dispatch('focusin',{target:this}); }
  click() { return this.dispatch('click'); }
}
function fixture(width = 1440) {
  const doc = new Events();doc.observers = [];doc.createElement = tag => new Element(doc,tag);doc.body = doc.createElement('body');doc.activeElement = doc.body;
  doc.getElementById = id => doc.body.querySelector('#' + id);
  doc.querySelectorAll = selector => doc.body.querySelectorAll(selector);
  const env = new Events();env.document = doc;env.innerWidth = width;env.innerHeight = 900;
  const media = [];
  env.matchMedia = query => { const item = new Events();item.limit = parseInt(/\d+/.exec(query)[0],10);item.matches = env.innerWidth <= item.limit;media.push(item);return item; };
  env.visualViewport = new Events();env.visualViewport.height = 900;
  env.MutationObserver = class { constructor(fn) { this.fn = fn; } observe() { doc.observers.push(this.fn); } disconnect() { doc.observers = doc.observers.filter(fn => fn !== this.fn); } };
  const copied = [];
  env.navigator = {clipboard:{writeText:async value => { copied.push(value); }}};
  env.fetch = () => { throw new Error('Presentation must not fetch'); };
  const add = (id,tag = 'div',parent = doc.body) => { const node = doc.createElement(tag);node.id = id;parent.appendChild(node);return node; };
  const header = add('header');add('historyToggle','button',header);add('contextToggle','button',header);add('taskCount','span',header);add('conversationTitle','h1',header);
  const layout = add('layout'), sidebar = add('acSidebar','aside',layout), main = add('main','main',layout), context = add('acContext','aside',layout);
  add('historyClose','button',sidebar);add('historySearch','input',sidebar);add('historyLast','button',sidebar);
  add('contextClose','button',context);
  ['workspaceTaskSummary','workspaceContextContent','workspaceContextEmpty','workspaceSources','chatJobs'].forEach(id => add(id,'div',context));
  add('ta','textarea',main);const activity = add('workspaceActivity','button',main);add('workspaceActivityText','span',activity);
  const backdrop = add('workspaceBackdrop');backdrop.hidden = true;
  const overlay = add('reportOvl');overlay.className = 'ovl';overlay.hidden = true;
  ['reportTitle','reportContent'].forEach(id => add(id,'div',overlay));
  ['reportClose','reportCopy'].forEach(id => add(id,'button',overlay));
  const otherModal = add('otherModal');otherModal.className = 'ovl';otherModal.hidden = true;add('otherClose','button',otherModal);
  let modalReturn = null;
  const options = {format, openModal:(selector,focus) => {
    modalReturn = doc.activeElement;doc.body.querySelector(selector).hidden = false;doc.body.querySelector(focus)?.focus();
  }, closeModal:selector => { doc.body.querySelector(selector).hidden = true;modalReturn?.focus(); }};
  const api = workspace.create(env,options);
  const get = id => doc.getElementById(id);
  return {env,doc,api,get,copied,add,resize(next) {
    env.innerWidth = next;media.forEach(item => { item.matches = next <= item.limit;item.dispatch('change'); });env.dispatch('resize');
  }, message(value) {
    const element = add('message-' + main.children.length,'div',main);const body = add('','div',element);body.className = 'ac-message-body';
    api.decorateMessage(element,value);return element;
  }};
}
const planMessage = (overrides = {}) => ({role:'assistant',content:'Результат',event_id:'task-a',ts:10,
  agent_plan:{status:'succeeded',steps:[{tool:'query:mingkh-complaints',status:'succeeded',params:{municipality:'Чехов',start:'2026-01-01',end:'2026-01-31',topic:'water'},
    result:{filters:{topic_label:'Питьевое водоснабжение'},requested_period:{from:'2026-01-01',to:'2026-01-31'}}}]},
  report_sources:[{source_url:'/mingkh',title:'МинЖКХ'}],...overrides});

test('source paths are normalized/deduplicated without resolving arbitrary URLs', () => {
  assert.deepEqual(workspace.sources([
    {source_url:' /mingkh/ ',title:'Первый'}, {source_url:'/mingkh',title:'Дубль'},
    {source_url:'/omsu-rating',title:'Рейтинг'}, {source_url:'/summarizer',title:'Текст'},
    ...['https://example.org/mingkh','//evil.test/mingkh','javascript:alert(1)','/mingkh?x=1','/mingkh#x','/mingkh/../edo','/mingkh\\evil','/%65do','/admin'].map(source_url => ({source_url}))
  ]),[{path:'/mingkh',title:'Первый'},{path:'/omsu-rating',title:'Рейтинг'},{path:'/summarizer',title:'Текст'}]);
});
test('metadata stays text, uses real source labels/status, and only autoopens once without focus movement', () => {
  const f = fixture();f.get('ta').focus();f.api.select('a');
  const message = planMessage();message.agent_plan.steps[0].params.municipality = '<img src=x onerror=alert(1)>';
  message.report_sources[0].title = '<script>not executable</script>';
  f.api.setDialog({id:'a',title:'Общий диалог',messages:[message]});
  assert.equal(f.doc.activeElement,f.get('ta'));
  assert.equal(f.get('conversationTitle').textContent,'Диалог');
  assert.equal(f.get('acContext').hidden,false);
  assert.match(f.get('workspaceContextContent').textContent,/Питьевое водоснабжение/);
  assert.match(f.get('workspaceContextContent').textContent,/<img src=x/);
  assert.equal(f.get('workspaceContextContent').querySelector('img'),null);
  assert.equal(f.get('workspaceSources').querySelector('a').getAttribute('href'),'/mingkh');
  assert.equal(f.get('workspaceSources').querySelector('script'),null);
  f.get('contextClose').click();f.api.setDialog({id:'a',messages:[message]});
  assert.equal(f.get('acContext').hidden,true);
  assert.equal(f.get('workspaceContextContent').querySelector('.ac-plan-step').getAttribute('data-status'),'succeeded');
  assert.doesNotMatch(f.get('workspaceTaskSummary').textContent,/%/);
});
test('conversation title is current safe text and resets with selection', () => {
  const f=fixture();f.api.select('a','Отчёт по воде');assert.equal(f.get('conversationTitle').textContent,'Отчёт по воде');
  f.api.setDialog({id:'a',title:'<img src=x> Данные',messages:[]});
  assert.equal(f.get('conversationTitle').textContent,'<img src=x> Данные');assert.equal(f.get('conversationTitle').querySelector('img'),null);
  f.api.select('b','Общий диалог');assert.equal(f.get('conversationTitle').textContent,'Диалог');
  f.api.reset();assert.equal(f.get('conversationTitle').textContent,'Новый диалог');
});
test('latest actual job stage and terminal per-step outcomes replace earlier report metadata', () => {
  const f = fixture();f.api.select('a');f.api.setDialog({id:'a',messages:[planMessage()]});
  const job = {job_id:'new',dialog_id:'a',created_at:20,status:'running',stage:'Получаем данные источника',result:{agent_plan:{status:'running',steps:[
    {label:'Рейтинг',status:'succeeded',params:{period:2},result:{period_label:'Девять месяцев 2026 года'}},
    {label:'Сравнение',status:'running',params:{municipality:'Химки'}}
  ]}}};
  f.api.setJobs({jobs:[job],pending:[]});
  assert.match(f.get('workspaceTaskSummary').textContent,/Получаем данные источника/);
  assert.match(f.get('workspaceContextContent').textContent,/Девять месяцев 2026 года/);
  assert.doesNotMatch(f.get('workspaceContextContent').textContent,/Чехов|2026-01-01/);
  assert.equal(f.get('taskCount').textContent,'1');
  job.status = 'interrupted';job.result.agent_plan.steps[1].status = 'interrupted';f.api.setJobs({jobs:[job]});
  assert.deepEqual(f.get('workspaceContextContent').querySelectorAll('.ac-plan-step').map(node => node.getAttribute('data-status')),['succeeded','interrupted']);
  assert.match(f.get('workspaceActivityText').textContent,/Выполнение прервано/);
  assert.equal(f.get('workspaceSources').hidden,true);
});
test('switching, revocation and filtered messages clear old metadata/reports and invalidate old controls', () => {
  const f = fixture();f.api.select('a','Первый');const message = planMessage();
  f.api.setDialog({id:'a',messages:[message]});const oldButton = f.message(message).querySelector('button');oldButton.click();
  assert.equal(f.get('reportOvl').hidden,false);
  f.api.select('b','Другой');assert.equal(f.get('reportOvl').hidden,true);
  assert.equal(f.get('reportContent').textContent,'');assert.equal(f.get('workspaceSources').textContent,'');
  oldButton.click();assert.equal(f.get('reportOvl').hidden,true);
  f.api.setDialog({id:'a',messages:[message]});assert.equal(f.get('workspaceContextContent').textContent,'');
  f.api.select('a');f.api.setDialog({id:'a',messages:[message]});f.message(message).querySelector('button').click();
  f.api.setDialog({id:'a',messages:[{...message,restricted:true,content:'Нет доступа'}]});
  f.api.setJobs({jobs:[{job_id:'task-a',dialog_id:'a',created_at:20,result:message,status:'succeeded'}]});
  assert.equal(f.get('reportOvl').hidden,true);assert.equal(f.get('workspaceContextContent').textContent,'');
  assert.equal(f.message({...message,restricted:true}).querySelector('button'),null);
  f.api.setDialog(null);assert.equal(f.get('acContext').hidden,true);assert.equal(f.get('reportContent').textContent,'');
});
test('narrow drawers trap focus, restore it, release background, and defer to existing modals', () => {
  const f = fixture(390);assert.equal(f.get('acSidebar').hidden,true);
  f.get('historyToggle').focus();f.get('historyToggle').click();
  assert.equal(f.doc.body.getAttribute('data-workspace-panel'),'history');assert.equal(f.get('workspaceBackdrop').hidden,false);
  assert.equal(f.doc.activeElement,f.get('historySearch'));assert.equal(f.get('main').inert,true);
  f.get('historyLast').focus();const tab = f.doc.dispatch('keydown',{key:'Tab'});
  assert.equal(tab.defaultPrevented,true);assert.equal(f.doc.activeElement,f.get('historyClose'));
  const reverse = f.doc.dispatch('keydown',{key:'Tab',shiftKey:true});assert.equal(reverse.defaultPrevented,true);
  assert.equal(f.doc.activeElement,f.get('historyLast'));
  f.get('otherModal').hidden = false;assert.equal(f.get('main').inert,false);f.get('otherClose').focus();
  assert.equal(f.doc.dispatch('keydown',{key:'Escape'}).defaultPrevented,false);
  assert.equal(f.doc.body.getAttribute('data-workspace-panel'),'history');
  f.get('otherModal').hidden = true;f.doc.dispatch('keydown',{key:'Escape'});
  assert.equal(f.doc.body.getAttribute('data-workspace-panel'),null);assert.equal(f.get('main').inert,false);
  assert.equal(f.doc.activeElement,f.get('historyToggle'));
  f.get('contextToggle').focus();f.get('contextToggle').click();f.get('workspaceBackdrop').click();
  assert.equal(f.get('acContext').hidden,true);assert.equal(f.doc.activeElement,f.get('contextToggle'));
});
test('plans never autoopen on phones; viewport changes keep docks nonmodal and expose keyboard height', () => {
  const f = fixture(390);f.api.select('a');f.api.setDialog({id:'a',messages:[planMessage()]});
  assert.equal(f.get('acContext').hidden,true);assert.equal(f.get('workspaceBackdrop').hidden,true);
  f.get('contextToggle').click();f.resize(1440);
  assert.equal(f.get('acContext').hidden,false);assert.equal(f.get('workspaceBackdrop').hidden,true);
  assert.equal(f.get('main').inert,false);assert.equal(f.get('acContext').getAttribute('aria-modal'),null);
  f.resize(390);assert.equal(f.get('acContext').hidden,true);assert.equal(f.get('workspaceBackdrop').hidden,true);
  f.env.visualViewport.height = 520;f.env.visualViewport.dispatch('resize');
  assert.equal(f.doc.body.classList.contains('ac-keyboard-open'),true);
  assert.equal(f.doc.body.style['--ac-viewport-height'],'520px');
  f.api.destroy();assert.equal(f.doc.observers.length,0);assert.equal(f.env.events.get('resize').size,0);
});
test('CSS-hidden controls are excluded from trapping and source-link focus survives polling', () => {
  const f=fixture(390);f.env.getComputedStyle=node=>({display:node.id==='historyClose'?'none':'block',visibility:'visible'});
  f.get('historyToggle').click();f.get('historyLast').focus();f.doc.dispatch('keydown',{key:'Tab'});
  assert.equal(f.doc.activeElement,f.get('historySearch'));
  f.api.closePanel('history');f.resize(1440);f.api.select('a');const message=planMessage();f.api.setDialog({id:'a',messages:[message]});
  f.get('workspaceSources').querySelector('a').focus();f.api.setJobs({jobs:[],pending:[]});
  assert.equal(f.doc.activeElement.getAttribute('href'),'/mingkh');assert.equal(f.doc.activeElement.isConnected,true);
  f.get('contextClose').focus();f.resize(390);assert.equal(f.doc.activeElement,f.get('contextToggle'));
});
test('report renders safe tables, copies readable text with feedback, and closes when current content changes', async () => {
  const f = fixture();f.api.select('a','Общий диалог');
  const message = {role:'assistant',event_id:'answer',content:'# Отчёт\n\n| Округ | Число |\n| --- | --- |\n| Чехов | 12 |\n\n<img src=x onerror=alert(1)> [link](javascript:alert(1))'};
  f.api.setDialog({id:'a',title:'Общий диалог',messages:[message]});
  const element = f.message(message);f.api.decorateMessage(element,message);assert.equal(element.querySelectorAll('button').length,1);
  element.querySelector('button').click();
  assert.equal(f.get('reportTitle').textContent,'Отчёт · Диалог');
  assert.match(f.get('reportContent').querySelector('.bubble').innerHTML,/class="ac-table-wrap".*role="region"/);
  assert.doesNotMatch(f.get('reportContent').querySelector('.bubble').innerHTML,/<(?:img|script|a)\b/);
  f.get('reportCopy').click();await Promise.resolve();assert.match(f.get('reportCopy').textContent,/скопирован/);
  assert.match(f.copied[0],/Чехов/);assert.doesNotMatch(f.copied[0],/<table|<td/);
  f.env.navigator.clipboard.writeText = async () => { throw new Error('Permission denied'); };
  f.get('reportCopy').click();await Promise.resolve();assert.match(f.get('reportCopy').textContent,/Не удалось/);
  f.api.setDialog({id:'a',messages:[{...message,content:'Изменённый результат'}]});
  assert.equal(f.get('reportOvl').hidden,true);assert.equal(f.get('reportContent').textContent,'');
  element.querySelector('button').click();assert.equal(f.get('reportOvl').hidden,true);
});
test('wide-report controls appear for substantial output and report data, not short ordinary replies', () => {
  const f=fixture();f.api.select('a');
  for(const message of [
    {role:'assistant',content:'Здравствуйте! Чем помочь?'},
    {role:'assistant',content:'Уточните муниципалитет.',agent_pending:{municipality:''}},
    {role:'assistant',content:'Короткий ответ',report_sources:[{source_url:'https://evil.test'}]}
  ]) assert.equal(f.message(message).querySelector('.ac-report-open'),null);
  for(const message of [
    {role:'assistant',content:'А'.repeat(800)},
    {role:'assistant',content:'Готово',artifact:{id:'report-file'}},
    {role:'assistant',content:'Готово',report_scope:{municipality:'Чехов'}},
    {role:'assistant',content:'Готово',report_sources:[{source_url:'/mingkh',title:'МинЖКХ'}]},
    {role:'assistant',content:'Готово',agent_plan:{steps:[{status:'succeeded'}]}},
    {role:'assistant',content:'Округ | Число\n--- | ---\nЧехов | 12'}
  ]) assert.ok(f.message(message).querySelector('.ac-report-open'));
});
test('pending, connection loss and confirmation stay discoverable without issuing commands', () => {
  const f = fixture(390);f.api.select('a');
  f.api.setJobs({jobs:[{dialog_id:'a',job_id:'proposal',created_at:1,status:'succeeded',result:{action:{kind:'agent_plan',plan:{steps:[{tool:'refresh:cds',params:{}}]}}}}]});
  assert.equal(f.get('workspaceActivity').hidden,false);assert.match(f.get('workspaceActivityText').textContent,/ждёт подтверждения/);
  assert.equal(f.get('acContext').hidden,true);f.get('workspaceActivity').click();assert.equal(f.get('acContext').hidden,false);
  f.api.setJobs({jobs:[],pending:[{error:'Неизвестен результат отправки'}]});assert.match(f.get('workspaceActivityText').textContent,/внимания/);
  f.api.setJobs({jobs:[],connection:'offline'});assert.match(f.get('workspaceActivityText').textContent,/Связь прервана/);
  f.api.setJobs({jobs:[],pending:[]});assert.equal(f.get('workspaceActivity').hidden,true);
});
