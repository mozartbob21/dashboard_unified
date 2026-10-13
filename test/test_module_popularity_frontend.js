// Run: node --test test/test_module_popularity_frontend.js
const test = require('node:test');
const assert = require('node:assert/strict');
const {createController, validPayload, POLL_MS, REQUEST_TIMEOUT_MS} = require('../static/module-popularity.js');

class Element {
  constructor(tag='div') {
    this.tagName=tag;this.children=[];this.dataset={};this.attributes={};this.hidden=false;this.disabled=false;this._text='';this.handlers=new Map();
  }
  set textContent(text) {this._text=String(text);this.children=[];}
  get textContent() {return this._text + this.children.map(child=>child.textContent).join(' ');}
  set innerHTML(_) {throw new Error('Labels must never be interpreted as HTML');}
  append(...items) {this.children.push(...items);}
  replaceChildren(...items) {this._text='';this.children=items;}
  setAttribute(name,value) {this.attributes[name]=String(value);}
  removeAttribute(name) {delete this.attributes[name];}
  addEventListener(name,callback) {this.handlers.set(name,callback);}
  removeEventListener(name,callback) {if(this.handlers.get(name)===callback)this.handlers.delete(name);}
}
function payload(visits=2) {
  return {period:'all_time', tracking_since:'2026-10-01T07:05:38Z', updated_at:'2026-10-13T10:00:00Z', visit_gap_minutes:30,
    modules:[
      {rank:1,module_id:'edo',label:'Заполненность данных',visits,unique_users:2,page_views:3,actions:1,last_seen_at:'2026-10-13T09:30:00Z'},
      {rank:2,module_id:'telegram',label:'Telegram — выбранные чаты',visits:0,unique_users:0,page_views:0,actions:0,last_seen_at:null}
    ]};
}
function pending() {let resolve,reject;const promise=new Promise((a,b)=>{resolve=a;reject=b;});return {promise,resolve,reject};}
function response(body,status=200,options={}) {return {ok:status>=200&&status<300,status,json:async()=>body,...options};}
function fixture() {
  const nodes=new Map(),requests=[],intervals=new Map(),timeouts=new Map(),events=new Map(),documentEvents=new Map();let token=0;
  const document={hidden:false,getElementById(id){if(!nodes.has(id))nodes.set(id,new Element());return nodes.get(id);},createElement:tag=>new Element(tag),addEventListener:(name,cb)=>documentEvents.set(name,cb),removeEventListener:name=>documentEvents.delete(name)};
  const env={document,AbortController,fetch:(url,options)=>{const request=pending();requests.push({url,options,...request});return request.promise;},setInterval:(callback,ms)=>{const id=++token;intervals.set(id,{callback,ms});return id;},clearInterval:id=>intervals.delete(id),setTimeout:(callback,ms)=>{const id=++token;timeouts.set(id,{callback,ms});return id;},clearTimeout:id=>timeouts.delete(id),addEventListener:(name,cb)=>events.set(name,cb),removeEventListener:name=>events.delete(name)};
  return {controller:createController(env),requests,intervals,timeouts,events,documentEvents,document,node:id=>document.getElementById(id)};
}
async function initial(f,data=payload()) {const done=f.controller.start();f.requests[0].resolve(response(data));await done;}

test('all rows remain visible, including four zero counters and absent last activity',async()=>{
  const f=fixture();await initial(f);
  const rows=f.node('popularityRows').children;
  assert.equal(rows.length,2);assert.equal(rows[0].children.length,7);
  assert.equal(rows[0].children[1].tagName,'th');assert.equal(rows[0].children[1].attributes.scope,'row');
  assert.deepEqual(rows[1].children.slice(2,6).map(cell=>cell.textContent),['0','0','0','0']);
  assert.equal(rows[1].children[6].textContent,'Нет активности');
  assert.match(f.node('popularitySince').textContent,/01.10.2026, 10:05:38/);
  assert.match(f.node('popularityUpdated').textContent,/13:00:00/);
  assert.equal(f.node('popularityContent').hidden,false);
});

test('separate admin sessions poll the same shared totals every ten seconds without caching',async()=>{
  const a=fixture(),b=fixture();await initial(a);await initial(b);
  assert.equal(POLL_MS,10000);
  const loadA=[...a.intervals.values()][0].callback();const loadB=[...b.intervals.values()][0].callback();
  for(const f of [a,b])f.requests[1].resolve(response(payload(7)));
  await Promise.all([loadA,loadB]);
  for(const f of [a,b]){
    assert.equal(f.node('popularityRows').children[0].children[2].textContent,'7');
    assert.equal(f.requests[1].url,'/api/users/module-popularity');
    assert.equal(f.requests[1].options.cache,'no-store');assert.equal(f.requests[1].options.credentials,'same-origin');
  }
});

test('a failed refresh preserves totals and last update, and reconnect recovers',async()=>{
  const f=fixture();await initial(f);const row=f.node('popularityRows').children[0],stamp=f.node('popularityUpdated').textContent;
  const load=f.controller.refresh();f.requests[1].reject(new Error('Offline'));await load;
  assert.equal(f.node('popularityRows').children[0],row);assert.equal(f.node('popularityUpdated').textContent,stamp);
  assert.match(f.node('popularityStatus').textContent,/последние полученные данные/);assert.equal(f.node('popularityRefresh').disabled,false);
  const resumed=f.events.get('online')();f.requests[2].resolve(response(payload(9)));await resumed;
  assert.equal(f.node('popularityRows').children[0].children[2].textContent,'9');assert.equal(f.node('popularityStatus').dataset.state,'success');
});

test('revoked access or a login redirect removes restricted statistics, and re-login can recover',async()=>{
  for(const denied of [response({},401),response({},403),response({},200,{redirected:true})]){
    const f=fixture();await initial(f);const load=f.controller.refresh();f.requests[1].resolve(denied);await load;
    assert.equal(f.node('popularityContent').hidden,true);assert.equal(f.node('popularityRows').children.length,0);
    assert.equal(f.node('popularityAccessNotice').hidden,false);assert.equal(f.node('popularitySince').textContent,'');
    const retry=f.controller.refresh();f.requests[2].resolve(response(payload()));await retry;
    assert.equal(f.node('popularityContent').hidden,false);assert.equal(f.node('popularityAccessNotice').hidden,true);
  }
});

test('hidden tabs pause reads and visibility return resumes with no overlapping requests',async()=>{
  const f=fixture();const load=f.controller.start();await f.controller.refresh();assert.equal(f.requests.length,1);
  f.requests[0].resolve(response(payload()));await load;
  f.document.hidden=true;await [...f.intervals.values()][0].callback();assert.equal(f.requests.length,1);
  f.document.hidden=false;const resumed=f.documentEvents.get('visibilitychange')();
  f.requests[1].resolve(response(payload()));await resumed;
  assert.equal(f.requests.length,2);
  f.controller.stop();assert.equal(f.intervals.size,0);assert.equal(f.events.size,0);assert.equal(f.documentEvents.size,0);
});

test('a stalled request is aborted and the refresh control becomes available again',async()=>{
  const f=fixture();const load=f.controller.start();assert.equal(f.node('popularityRefresh').disabled,true);
  const timeout=[...f.timeouts.values()][0];assert.equal(timeout.ms,REQUEST_TIMEOUT_MS);timeout.callback();
  assert.equal(f.requests[0].options.signal.aborted,true);f.requests[0].reject(new Error('AbortError'));await load;
  assert.equal(f.node('popularityRefresh').disabled,false);assert.equal(f.timeouts.size,0);assert.equal(f.node('popularityStatus').dataset.state,'error');
});

test('unchanged totals preserve DOM nodes and labels are plain text',async()=>{
  const f=fixture(),data=payload();data.modules[0].label='<img src=x onerror=alert(1)>';
  await initial(f,data);const row=f.node('popularityRows').children[0];assert.equal(row.children[1].textContent,data.modules[0].label);
  const load=f.controller.refresh();f.requests[1].resolve(response({...data,updated_at:'2026-10-13T10:00:10Z'}));await load;
  assert.equal(f.node('popularityRows').children[0],row);assert.match(f.node('popularityUpdated').textContent,/13:00:10/);
});

test('malformed responses keep prior totals instead of fabricating an empty ranking',async()=>{
  const f=fixture();await initial(f);const load=f.controller.refresh();f.requests[1].resolve(response({}));await load;
  assert.equal(f.node('popularityRows').children.length,2);assert.equal(f.node('popularityStatus').dataset.state,'error');
  assert.equal(validPayload(payload()),true);
  for(const change of [data=>data.modules=[],data=>data.modules[1].module_id='edo',data=>data.modules[0].visits=-1,
    data=>data.modules[1].visits=100,data=>data.modules[1]=null,data=>data.updated_at='invalid']){
    const data=payload();change(data);assert.ok(!validPayload(data));
  }
});
