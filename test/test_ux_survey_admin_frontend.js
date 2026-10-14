// Run: node --test test/test_ux_survey_admin_frontend.js
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const sourcePath = process.env.SURVEY_ADMIN_SCRIPT || path.join(__dirname, '../static/ux-survey-admin.js');
const {createController, validPayload, POLL_MS, REQUEST_TIMEOUT_MS} = require(sourcePath);

class Element {
  constructor(tag='div') {
    this.tagName=tag;this.children=[];this.dataset={};this.style={};this.attributes={};this.hidden=false;this.disabled=false;this._text='';this.handlers=new Map();
  }
  set textContent(text) {this._text=String(text);this.children=[];}
  get textContent() {return this._text + this.children.map(child=>child.textContent).join(' ');}
  set innerHTML(_) {throw new Error('Untrusted survey text must never be interpreted as HTML');}
  append(...items) {this.children.push(...items);}
  replaceChildren(...items) {this._text='';this.children=items;}
  setAttribute(name,value) {this.attributes[name]=String(value);}
  removeAttribute(name) {delete this.attributes[name];}
  addEventListener(name,callback) {this.handlers.set(name,callback);}
  removeEventListener(name,callback) {if(this.handlers.get(name)===callback)this.handlers.delete(name);}
  closest(selector) {return selector==='[data-admin-panel]' ? this.panel || null : null;}
}
function payload(options={}) {
  const completed=options.completed_count??2;
  return {
    completed_count:completed,invited_count:4,completion_rate:50,
    questions:Array.from({length:6},(_,index)=>({id:`q${index+1}`,label:`Вопрос ${index+1}`,count:completed,average:completed?2.5:null,distribution:Array.from({length:6},(_,score)=>({score,count:completed && [0,5].includes(score)?1:0}))})),
    daily:completed?[{date:'2026-10-13',count:2}]:[],
    comments:completed?[{submitted_at:'2026-10-13T10:45:00Z',text:'Сделать поиск заметнее'}]:[], ...options
  };
}
function pending() {let resolve,reject;const promise=new Promise((a,b)=>{resolve=a;reject=b;});return {promise,resolve,reject};}
function response(body,status=200,options={}) {return {ok:status>=200&&status<300,status,json:async()=>body,...options};}
function fixture({section,withPanel=false}={}) {
  const nodes=new Map(), requests=[], intervals=new Map(), timeouts=new Map();let token=0;
  const events=new Map(),documentEvents=new Map();
  const document={body:new Element('body'),hidden:false,getElementById(id){if(!nodes.has(id))nodes.set(id,new Element());return nodes.get(id);},createElement:tag=>new Element(tag),addEventListener:(name,cb)=>documentEvents.set(name,cb),removeEventListener:(name)=>documentEvents.delete(name)};
  if(section!==undefined)document.body.dataset.adminSection=section;
  const panel=withPanel?new Element():null;
  if(panel){panel.dataset.adminPanel='feedback';panel.hidden=section!=='feedback';document.getElementById('feedbackAdmin').panel=panel;}
  const env={document,AbortController,fetch:(url,options)=>{const request=pending();requests.push({url,options,...request});return request.promise;},setInterval:(callback,ms)=>{const id=++token;intervals.set(id,{callback,ms});return id;},clearInterval:id=>intervals.delete(id),setTimeout:(callback,ms)=>{const id=++token;timeouts.set(id,{callback,ms});return id;},clearTimeout:id=>timeouts.delete(id),addEventListener:(name,cb)=>events.set(name,cb),removeEventListener:name=>events.delete(name)};
  const controller=createController(env);
  return {controller,requests,intervals,timeouts,events,documentEvents,document,panel,node:id=>document.getElementById(id),select(section){document.body.dataset.adminSection=section;if(panel)panel.hidden=section!=='feedback';return documentEvents.get('admin:sectionchange')?.({detail:{section}});}};
}
async function initial(f,data=payload()) {const done=f.controller.start();f.requests[0].resolve(response(data));await done;}

test('zero means zero, and its distribution is visible with one decimal mean',async()=>{
  const f=fixture();const data=payload();
  data.questions[0]={id:'q1',label:'Удобство',count:2,average:0,distribution:Array.from({length:6},(_,score)=>({score,count:score===0?2:0}))};
  await initial(f,data);
  assert.equal(f.node('feedbackCompleted').textContent,'2');
  assert.equal(f.node('feedbackConversion').textContent,'50,0%');
  const card=f.node('feedbackQuestions').children[0];
  assert.match(card.textContent,/0,0/);assert.match(card.textContent,/Ответов: 2/);
  const distribution=card.children[2];
  assert.equal(distribution.children.length,6);assert.equal(distribution.children[0].children[0].textContent,'0');
  assert.match(distribution.children[0].textContent,/2 · 100%/);
  assert.match(f.node('feedbackUpdated').textContent,/Последнее обновление/);
});

test('empty survey shows no invented averages or completion percentages',async()=>{
  const f=fixture();await initial(f,payload({completed_count:0,invited_count:0,completion_rate:null}));
  assert.equal(f.node('feedbackConversion').textContent,'—');
  assert.equal(f.node('feedbackRatingsEmpty').hidden,false);assert.equal(f.node('feedbackDailyEmpty').hidden,false);assert.equal(f.node('feedbackCommentsEmpty').hidden,false);
  assert.equal(f.node('feedbackQuestions').children[0].children[1].children[0]._text,'—');
  assert.equal(f.node('feedbackResults').hidden,false);
});

test('new submissions are fetched independently by two open admin sessions every ten seconds',async()=>{
  const a=fixture(),b=fixture();await initial(a);await initial(b);
  assert.equal(POLL_MS,10000);assert.equal([...a.intervals.values()][0].ms,10000);
  const changed=payload({completed_count:3,completion_rate:75});
  const ra=[...a.intervals.values()][0].callback();const rb=[...b.intervals.values()][0].callback();
  a.requests[1].resolve(response(changed));b.requests[1].resolve(response(changed));await Promise.all([ra,rb]);
  for(const f of [a,b]){assert.equal(f.node('feedbackCompleted').textContent,'3');assert.equal(f.node('feedbackConversion').textContent,'75,0%');assert.equal(f.requests[1].options.cache,'no-store');assert.equal(f.requests[1].options.credentials,'same-origin');}
});

test('a lost connection preserves results and last update time, then recovers automatically',async()=>{
  const f=fixture();await initial(f);const previousTime=f.node('feedbackUpdated').textContent;
  const failed=f.controller.refresh();f.requests[1].reject(new Error('Network error'));await failed;
  assert.equal(f.node('feedbackResults').hidden,false);assert.equal(f.node('feedbackCompleted').textContent,'2');assert.equal(f.node('feedbackUpdated').textContent,previousTime);
  assert.match(f.node('feedbackConnection').textContent,/последние полученные данные/);assert.equal(f.node('feedbackRefresh').disabled,false);
  const restored=f.events.get('online')();f.requests[2].resolve(response(payload({completed_count:3})));await new Promise(resolve=>setImmediate(resolve));
  assert.equal(f.node('feedbackCompleted').textContent,'3');assert.equal(f.node('feedbackConnection').dataset.state,'success');
});

test('comments including HTML, long text and line breaks remain plain text',async()=>{
  const f=fixture(),text='<img src=x onerror=alert(1)>\n<script>evil()</script> '+ 'аб'.repeat(950);
  await initial(f,payload({comments:[{submitted_at:'2026-10-13T10:00:00Z',text},{submitted_at:'2026-10-12T10:00:00Z',text:'  ' }]}));
  const items=f.node('feedbackComments').children;assert.equal(items.length,1);assert.equal(items[0].children[1].textContent,text);assert.equal(items[0].children[1].tagName,'p');
  assert.equal(items[0].children[0].tagName,'time');assert.match(items[0].children[0].textContent,/13:00:00/);
});

test('401, 403 and a redirected login page erase analytics from the DOM',async()=>{
  for(const rejected of [response({},401),response({},403),response({},200,{redirected:true})]){
    const f=fixture();await initial(f);const load=f.controller.refresh();f.requests[1].resolve(rejected);await load;
    assert.equal(f.node('feedbackResults').hidden,true);assert.equal(f.node('feedbackComments').children.length,0);assert.equal(f.node('feedbackQuestions').children.length,0);assert.equal(f.node('feedbackCompleted').textContent,'—');
    assert.equal(f.node('feedbackAccessNotice').hidden,false);assert.match(f.node('feedbackConnection').textContent,/только администраторам/);
  }
});

test('overlapping refreshes do not issue duplicate requests and stalled reads are aborted',async()=>{
  const f=fixture();const start=f.controller.start();assert.equal(f.node('feedbackRefresh').disabled,true);
  await f.controller.refresh();assert.equal(f.requests.length,1);
  const timeout=[...f.timeouts.values()][0];assert.equal(timeout.ms,REQUEST_TIMEOUT_MS);timeout.callback();assert.equal(f.requests[0].options.signal.aborted,true);
  f.requests[0].reject(new Error('AbortError'));await start;
  assert.equal(f.node('feedbackRefresh').disabled,false);assert.equal(f.timeouts.size,0);
});

test('unchanged polling does not replace selected/read DOM content',async()=>{
  const f=fixture();await initial(f);const comment=f.node('feedbackComments').children[0];
  const refresh=f.controller.refresh();f.requests[1].resolve(response({...payload(),generated_at:'different'}));await refresh;
  assert.equal(f.node('feedbackComments').children[0],comment);
});

test('hidden tabs pause requests, returning to the tab resumes, and stop cleans listeners',async()=>{
  const f=fixture();await initial(f);f.document.hidden=true;await f.controller.refresh();assert.equal(f.requests.length,1);
  f.document.hidden=false;f.documentEvents.get('visibilitychange')();assert.equal(f.requests.length,2);f.requests[1].resolve(response(payload()));await new Promise(resolve=>setImmediate(resolve));
  f.controller.stop();assert.equal(f.intervals.size,0);assert.equal(f.events.size,0);assert.equal(f.documentEvents.size,0);
});

test('malformed payloads cannot replace valid analytics with fabricated zeros',async()=>{
  const f=fixture();await initial(f);const read=f.controller.refresh();f.requests[1].resolve(response({}));await read;
  assert.equal(f.node('feedbackCompleted').textContent,'2');assert.equal(f.node('feedbackConnection').dataset.state,'error');
  assert.equal(validPayload(payload()),true);assert.equal(validPayload({...payload(),completion_rate:Infinity}),false);
  assert.equal(validPayload({...payload(),questions:[]}),false);
});

test('feedback waits for its section then refreshes on entry while preserving existing comments on navigation',async()=>{
  const f=fixture({section:'users',withPanel:true});await f.controller.start();
  await [...f.intervals.values()][0].callback();await f.events.get('online')();
  assert.equal(f.requests.length,0);
  const entering=f.select('feedback');assert.equal(f.requests.length,1);
  f.requests[0].resolve(response(payload()));await entering;
  const comment=f.node('feedbackComments').children[0],stamp=f.node('feedbackUpdated').textContent;
  await f.select('statistics');await [...f.intervals.values()][0].callback();
  await f.events.get('online')();assert.equal(f.requests.length,1);
  assert.equal(f.node('feedbackComments').children[0],comment);
  assert.equal(f.node('feedbackUpdated').textContent,stamp);
  const returning=f.select('feedback');assert.equal(f.requests.length,2);
  assert.equal(f.node('feedbackComments').children[0],comment);
  f.requests[1].resolve(response(payload({completed_count:3})));await returning;
  assert.equal(f.node('feedbackCompleted').textContent,'3');
  assert.equal(f.requests[1].url,'/api/ux-survey/admin');
  assert.equal(f.requests[1].options.credentials,'same-origin');
});

test('a hidden browser tab or hidden feedback panel pauses reads until both become visible',async()=>{
  const f=fixture({section:'users',withPanel:true});await f.controller.start();
  f.document.hidden=true;await f.select('feedback');assert.equal(f.requests.length,0);
  f.document.hidden=false;f.panel.hidden=true;
  await f.documentEvents.get('visibilitychange')();assert.equal(f.requests.length,0);
  f.panel.hidden=false;const resume=f.documentEvents.get('visibilitychange')();
  f.requests[0].resolve(response(payload()));await resume;
  assert.equal(f.requests.length,1);
  f.controller.stop();assert.equal(f.documentEvents.size,0);assert.equal(f.intervals.size,0);
});

test('standalone feedback keeps polling if either admin navigation marker is absent',async()=>{
  for(const options of [{section:'users'},{withPanel:true}]){
    const f=fixture(options);await initial(f);assert.equal(f.requests.length,1);
  }
});

test('pending access denial clears comments even after switching away from feedback',async()=>{
  for(const status of [401,403]){
    const f=fixture({section:'feedback',withPanel:true});await initial(f);
    const load=f.controller.refresh();await f.select('integrations');
    assert.equal(f.requests[1].options.signal.aborted,false);
    f.requests[1].resolve(response({},status));await load;
    assert.equal(f.node('feedbackComments').children.length,0);
    assert.equal(f.node('feedbackQuestions').children.length,0);
    assert.equal(f.node('feedbackCompleted').textContent,'—');
    assert.equal(f.node('feedbackResults').hidden,true);
  }
});

test('feedback section changes never overlap an already pending request',async()=>{
  const f=fixture({section:'feedback',withPanel:true});await initial(f);
  const load=f.controller.refresh();await f.select('users');await f.select('feedback');
  assert.equal(f.requests.length,2);
  f.requests[1].resolve(response(payload({completed_count:3})));await load;
  assert.equal(f.node('feedbackCompleted').textContent,'3');
});
