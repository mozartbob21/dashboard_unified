'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const jobs = require('../static/aichat-jobs.js');
const {webcrypto} = require('node:crypto');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');
const template = readFileSync(join(__dirname, '../templates/aichat.html'), 'utf8');
const settle = async () => {for(let i=0;i<14;i++)await Promise.resolve();};
function harness(fetcher, callbacks={}) {
  let next=0, now=0;
  const timers=new Map(),calls=[],states=[],events={};
  const env={crypto:webcrypto,AbortController,addEventListener:(key,fn)=>events[key]=fn,
    removeEventListener:key=>delete events[key],
    setTimeout:(fn,ms)=>{const id=++next;timers.set(id,{fn,at:now+ms});return id;},
    clearTimeout:id=>timers.delete(id),fetch:async (url,options)=>{
      calls.push({url,options});return fetcher(url,options);
    }};
  const ui=jobs.create(env,{onChange:state=>states.push(state),...callbacks});
  async function tick(ms) {
    const until=now+ms;
    for(let guard=0;guard<100;guard++) {
      const item=[...timers].filter(([,v])=>v.at<=until).sort((a,b)=>a[1].at-b[1].at)[0];
      if(!item)break;timers.delete(item[0]);now=item[1].at;item[1].fn();await settle();
    }
    now=until;await settle();
  }
  return {ui,calls,states,events,tick,timers};
}
const reply=(body,status=200)=>({ok:status>=200&&status<300,status,json:async()=>body});
const job=(status='running',extra={})=>({job_id:'job-one',dialog_id:'dialog-one',status,stage:'Получение данных',...extra});

test('acceptance uses durable identifier and does not wait for the task result',async()=>{
  let accepted;
  const h=harness(async(url,options)=>options?.method==='POST'?reply({...job('queued')},202):reply({items:[job()]}),
    {onAccepted:ticket=>accepted=ticket});
  const body=new FormData();body.set('text','Запусти проверку');
  const ticket=await h.ui.submit(body,'stable-key-123456','Запусти проверку');
  assert.equal(ticket.job_id,'job-one');assert.equal(accepted.job_id,'job-one');
  assert.equal(h.calls[0].options.body.get('request_id'),'stable-key-123456');
  assert.equal(h.calls[0].options.body.get('background'),'true');
  assert.equal(h.states.at(-1).pending.length,0);
  h.ui.destroy();
});

test('broken acceptance never retries POST automatically and manual retry keeps exact body and key',async()=>{
  let posts=0;
  const h=harness(async(url,options)=>{
    if(options?.method==='POST'){posts++;if(posts===1)throw new TypeError('connection lost');return reply(job('queued'));}
    return reply({items:[]});
  });
  const body=new FormData();body.set('text','Запусти проверку');
  await h.ui.submit(body,'stable-key-123456','Запусти проверку');
  assert.equal(h.states.at(-1).pending[0].retryable,true);
  await h.tick(30000);assert.equal(posts,1);
  await h.ui.retry('stable-key-123456');assert.equal(posts,2);
  const commands=h.calls.filter(call=>call.options.method==='POST');
  assert.equal(commands[0].options.body,commands[1].options.body);
  assert.equal(commands[1].options.body.get('request_id'),'stable-key-123456');h.ui.destroy();
});

test('long jobs remain tracked beyond request timeout and terminal callback runs once',async()=>{
  let status='running',terminals=0;
  const activity=[];
  const h=harness(async()=>reply({items:[job(status)]}),{onTerminal:()=>terminals++,onActivity:active=>activity.push(active)});
  h.ui.select('dialog-one');await h.tick(0);await h.tick(65000);
  assert.ok(h.calls.length>20);assert.equal(terminals,0);assert.deepEqual(activity,[true]);
  status='succeeded';await h.tick(2000);await h.tick(10000);
  assert.equal(terminals,1);assert.deepEqual(activity,[true,false]);h.ui.destroy();
});

test('lost polling connection backs off, reconnects and never issues a command',async()=>{
  let failed=true;
  const h=harness(async()=>{if(failed)throw new TypeError('offline');return reply({items:[job('failed')]});});
  await h.tick(0);await h.tick(2000);await h.tick(4000);await h.tick(8000);
  assert.equal(h.calls.length,4);assert.match(h.states.at(-1).connection,/повторного запуска не будет/);
  failed=false;h.events.online();await h.tick(0);
  assert.equal(h.states.at(-1).connection,'');
  assert.equal(h.states.at(-1).jobs[0].status,'failed');
  assert.ok(h.calls.every(call=>!call.options.method));h.ui.destroy();
});

test('reload discovers already accepted jobs without replaying command; terminal interrupted is final',async()=>{
  const completed=[];
  const h=harness(async()=>reply({items:[job('interrupted')]}),{onTerminal:items=>completed.push(...items)});
  h.ui.select('dialog-one');await h.tick(0);
  assert.equal(h.calls[0].url,'/aichat/api/jobs?dialog_id=dialog-one');
  assert.equal(completed[0].status,'interrupted');assert.ok(h.calls.every(call=>!call.options.method));h.ui.destroy();
});

test('confirmation is explicit, concurrent clicks share one request and retry uses same proposal URL',async()=>{
  let release;
  const h=harness(async(url,options)=>{
    if(options?.method==='POST')return new Promise(resolve=>release=()=>resolve(reply(job('queued'))));
    return reply({items:[job('succeeded',{result:{action:{label:'ЕДДС',summary:'Текущая проверка'}}})]});
  });
  await h.tick(0);assert.equal(h.calls.filter(call=>call.options.method==='POST').length,0);
  const first=h.ui.confirm('job-one');await h.ui.confirm('job-one');
  assert.equal(h.calls.filter(call=>call.options.method==='POST').length,1);
  release();await first;
  assert.equal(h.states.at(-1).jobs[0].result.action.confirmed_job_id,'job-one');
  assert.equal(h.calls.at(-1).url,'/aichat/api/jobs/job-one/confirm');h.ui.destroy();
});

test('authorisation failure stops polling and does not pretend task completed',async()=>{
  const h=harness(async()=>reply({detail:'login'},401));
  await h.tick(0);await h.tick(60000);
  assert.equal(h.calls.length,1);assert.match(h.states.at(-1).connection,/Сессия завершилась/);h.ui.destroy();
});

test('session expiry clears cached jobs and uncertain drafts before notifying the page',async()=>{
  let expired=false,unauthorized=0,terminal=0,snapshot;
  const activity=[];
  const h=harness(async(url,options)=>options?.method==='POST'?reply({detail:'offline'},503):
    expired?reply({detail:'login'},401):reply({items:[job()]}), {
    onActivity:value=>activity.push(value),onTerminal:()=>terminal++,
    onUnauthorized:error=>{unauthorized++;assert.equal(error.status,401);snapshot=h.states.at(-1);}
  });
  await h.tick(0);await h.ui.submit(new FormData(),'uncertain-request','Секретный запрос');
  assert.equal(h.states.at(-1).jobs.length,1);assert.equal(h.states.at(-1).pending.length,1);
  expired=true;await h.tick(0);
  assert.equal(unauthorized,1);assert.deepEqual(snapshot.jobs,[]);assert.deepEqual(snapshot.pending,[]);
  assert.equal(snapshot.lastUpdated,'');assert.deepEqual(activity,[true,false]);assert.equal(terminal,0);
  const count=h.calls.length;
  await h.ui.confirm('job-one');await h.ui.retry('uncertain-request');await h.ui.submit(new FormData(),'new-request','Новое');
  await h.tick(60000);assert.equal(h.calls.length,count);assert.equal(h.events.online,undefined);h.ui.destroy();
});

test('an accepted response arriving after session expiry cannot restore private data',async()=>{
  let release,accepted=0,unauthorized=0;
  const h=harness(async(url,options)=>options?.method==='POST'?
    new Promise(resolve=>{release=()=>resolve(reply(job('queued')));}):reply({detail:'login'},401),
    {onAccepted:()=>accepted++,onUnauthorized:()=>unauthorized++});
  const sending=h.ui.submit(new FormData(),'pending-request','Текст запроса');await settle();
  await h.tick(0);release();assert.equal(await sending,null);
  assert.equal(unauthorized,1);assert.equal(accepted,0);assert.deepEqual(h.states.at(-1).jobs,[]);
  assert.deepEqual(h.states.at(-1).pending,[]);h.ui.destroy();
});

test('send and confirmation 401 use the same expiry path without a retry action',async()=>{
  for(const operation of ['submit','confirm']) {
    let unauthorized=0;
    const h=harness(async()=>reply({detail:'login'},401),{onUnauthorized:()=>unauthorized++});
    if(operation==='submit')await h.ui.submit(new FormData(),'request-key','Текст');else await h.ui.confirm('proposal');
    assert.equal(unauthorized,1);assert.deepEqual(h.states.at(-1).jobs,[]);assert.deepEqual(h.states.at(-1).pending,[]);
    assert.equal(h.ui.confirming('proposal'),false);await h.tick(30000);assert.equal(h.calls.length,1);h.ui.destroy();
  }
});

test('terminal history read failure is retried without rerunning task',async()=>{
  let reads=0;
  const h=harness(async()=>reply({items:[job('succeeded')]}),{onTerminal:()=>{reads++;if(reads===1)throw new Error('history offline');}});
  await h.tick(0);await h.tick(2000);assert.equal(reads,2);
  assert.ok(h.calls.every(call=>!call.options.method));h.ui.destroy();
});

test('switching dialogs ignores old in-flight poll data',async()=>{
  let resolveOld;
  const h=harness(async url=>url.endsWith('first')?new Promise(resolve=>resolveOld=resolve):reply({items:[job('running',{dialog_id:'second'})]}));
  h.ui.select('first');await h.tick(0);h.ui.select('second');
  resolveOld(reply({items:[job('running',{dialog_id:'first'})]}));await settle();
  assert.equal(h.states.at(-1).jobs.length,0);
  await h.tick(10000);
  assert.equal(h.states.at(-1).jobs[0].dialog_id,'second');h.ui.destroy();
});

function fakeDocument() {
  const doc={activeElement:null};
  class Node {
    constructor(tag){this.tagName=tag;this.childNodes=[];this.attrs={};this.textContent='';this.listeners={};}
    set innerHTML(_){throw new Error('Task rendering must never use HTML');}
    appendChild(node){this.childNodes.push(node);node.parent=this;return node;}
    replaceChildren(fragment){this.childNodes=fragment.childNodes;this.childNodes.forEach(node=>node.parent=this);}
    setAttribute(key,value){this.attrs[key]=value;}
    getAttribute(key){return this.attrs[key];}
    addEventListener(key,fn){this.listeners[key]=fn;}
    contains(node){return node===this||this.childNodes.some(child=>child.contains(node));}
    querySelectorAll(){return this.childNodes.flatMap(node=>[...(node.attrs['data-control']?[node]:[]),...node.querySelectorAll()]);}
    focus(){doc.activeElement=this;}
  }
  doc.createElement=tag=>new Node(tag);doc.createDocumentFragment=()=>new Node('fragment');
  return {doc,host:new Node('div')};
}

test('task labels and stages are safe text; confirmation needs a real click and focus survives poll',()=>{
  const {doc,host}=fakeDocument(),calls=[];
  const controller={confirming:()=>false,confirm:id=>calls.push(id)};
  const attack='<img src=x onerror=alert(1)>';
  const state={jobs:[job('succeeded',{result:{action:{label:attack,summary:'Параметры: '+attack}}})],pending:[],lastUpdated:'2026-11-04T14:00:00Z'};
  jobs.render(doc,host,state,controller);
  assert.equal(host.childNodes[0].childNodes[0].textContent,attack);
  assert.equal(calls.length,0);
  const button=host.querySelectorAll().find(node=>node.attrs['data-control']==='confirm-job-one');
  button.focus();jobs.render(doc,host,state,controller);
  assert.equal(doc.activeElement.attrs['data-control'],'confirm-job-one');
  doc.activeElement.listeners.click();assert.deepEqual(calls,['job-one']);
});

test('ambiguous receipt cannot discard its retry key, rejected draft can be dismissed',async()=>{
  let status=503;
  const dismissed=[];
  const h=harness(async(url,options)=>options?.method==='POST'?reply({detail:'Запрос отклонён'},status):reply({items:[]}),
    {onDismissed:id=>dismissed.push(id)});
  const body=new FormData();await h.ui.submit(body,'stable-key-123456','Запрос');
  h.ui.dismiss('stable-key-123456');assert.equal(h.states.at(-1).pending.length,1);
  status=403;await h.ui.retry('stable-key-123456');h.ui.dismiss('stable-key-123456');
  assert.equal(h.states.at(-1).pending.length,0);assert.deepEqual(dismissed,['stable-key-123456']);h.ui.destroy();
});

function templateCallback(name, next, globals) {
  const start = template.indexOf(name + ':') + name.length + 1;
  const end = template.indexOf(next ? '\n  ' + next + ':' : '\n});', start);
  const source = template.slice(start, end).trim().replace(/,$/, '');
  assert.ok(start > name.length && end > start);
  return vm.runInNewContext('(' + source + ')', globals);
}

test('shared task started by another user refreshes the opened dialogue without any POST',async()=>{
  let phase='running',reads=0;
  const refreshed=[];
  const onTerminal=templateCallback('onTerminal', null, {
    currentDid:'shared-dialog',refreshDialog:async id=>refreshed.push(id),loadDialogs:()=>reads++
  });
  const h=harness(async()=>reply({items:[job(phase,{job_id:'colleague-task',dialog_id:'shared-dialog'})]}),{onTerminal});
  h.ui.select('shared-dialog');await h.tick(0);
  assert.equal(refreshed.length,0);
  phase='succeeded';await h.tick(2000);await h.tick(10000);
  assert.deepEqual(refreshed,['shared-dialog']);assert.equal(reads,1);
  assert.ok(h.calls.every(call=>!call.options.method));h.ui.destroy();
});

test('confirmation of a running shared check shows its existing dialogue and never claims a new run',async()=>{
  const note={textContent:'',hidden:true};
  Object.defineProperty(note,'innerHTML',{set(){throw new Error('Unsafe HTML write');}});
  const navigated=[],refreshed=[],notifications=[];
  const globals={submissions:new Map(),currentDid:'my-proposal',lastSubmission:null,
    selectDialog:id=>navigated.push(id),refreshDialog:async id=>refreshed.push(id),
    loadDialogs:()=>{},$:()=>note};
  const accepted=templateCallback('onAccepted','onTerminal',globals);
  const existing=job('running',{job_id:'existing-check',dialog_id:'shared-dialog',
    already_running:true,message:'Проверка уже запущена. Повторный запуск не создан.'});
  const h=harness(async(url,options)=>options?.method==='POST'?reply(existing,202):reply({items:[existing]}),
    {onAccepted:(ticket,id)=>{notifications.push(ticket);accepted(ticket,id);}});
  const ticket=await h.ui.confirm('my-proposal');await h.tick(20000);
  assert.equal(ticket.job_id,'existing-check');
  assert.equal(notifications.length,1);
  assert.deepEqual(navigated,['shared-dialog']);assert.deepEqual(refreshed,['shared-dialog']);
  assert.equal(note.hidden,false);assert.match(note.textContent,/Проверка уже запущена/);
  assert.equal(h.calls.filter(call=>call.options.method==='POST').length,1);h.ui.destroy();
});

test('occupied shared dialogue keeps the rejected draft pending without pretending it was accepted',async()=>{
  const accepted=[];
  const h=harness(async(url,options)=>options?.method==='POST'?
    reply({detail:'В этом диалоге уже выполняется задача. Дождитесь её завершения.'},409):reply({items:[job()]}),
    {onAccepted:ticket=>accepted.push(ticket)});
  const body=new FormData();body.set('text','Вопрос второго пользователя');
  assert.equal(await h.ui.submit(body,'shared-request-1234','Вопрос второго пользователя'),null);
  await h.tick(30000);
  assert.equal(accepted.length,0);assert.equal(h.states.at(-1).pending.length,1);
  assert.match(h.states.at(-1).pending[0].error,/уже выполняется/);
  assert.equal(h.calls.filter(call=>call.options.method==='POST').length,1);h.ui.destroy();
});

test('unconfirmed response cannot falsely mark a shared proposal as launched',async()=>{
  const accepted=[];
  const h=harness(async()=>reply({already_running:true}),{onAccepted:ticket=>accepted.push(ticket)});
  assert.equal(await h.ui.confirm('proposal'),null);
  assert.equal(accepted.length,0);assert.match(h.states.at(-1).connection,/Подтверждение запуска не получено/);
  h.ui.destroy();
});
