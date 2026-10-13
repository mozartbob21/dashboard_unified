'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {cleanDraft, missingRatings, commentLength, SurveyAPI, SurveySession, safeToInvite} = require('../static/ux-survey.js');
const state = (draft = {}, version = 0) => ({completed:false,draft:cleanDraft(draft),draft_version:version});
const ratings = score => Object.fromEntries(['q1','q2','q3','q4','q5','q6'].map(id => [id,score]));
const deferred = () => {let resolve, reject; const promise = new Promise((a,b)=>{resolve=a;reject=b;}); return {promise,resolve,reject};};
const sessionFor = (t, api, initial = state(), notify) => {const session = new SurveySession(api,initial,notify); t.after(()=>clearTimeout(session.timer)); return session;};

test('zero is an answered value while missing/null is unanswered; invalid ratings never become zero', () => {
  assert.deepEqual(missingRatings(cleanDraft({ratings:ratings(0)})),[]);
  assert.deepEqual(missingRatings(cleanDraft({ratings:{q1:0,q2:null,q3:true,q4:-1,q5:6,q6:'4'}})),['q2','q3','q4','q5','q6']);
  assert.equal(commentLength('🙂А'),2);
});

test('draft preserves zero and safe literal text, then restores a server saved draft on another session', async t => {
  const saved = {ratings:ratings(0),comment:'<img src=x onerror=alert(1)>\nТекст'};
  let server = state();
  const api = {request:async (p,m,d)=>{assert.equal(p,'/draft');assert.equal(m,'PUT');server=state(d,server.draft_version+1);return server;}};
  const session = sessionFor(t,api);session.change(saved);await session.save();
  assert.equal(session.dirty,false);assert.deepEqual(server.draft,saved);
  const other = sessionFor(t,api,server);assert.deepEqual(other.data,saved);assert.equal(other.dirty,false);
});

test('in-flight save never marks newer unsent changes as persisted and queued save uses fresh version', async t => {
  const first = deferred(); const calls=[];
  const api = {request:(p,m,d)=>{calls.push(d);return calls.length===1 ? first.promise : Promise.resolve(state(d,2));}};
  const session=sessionFor(t,api);session.change({ratings:{q1:1}});const save1=session.save();
  session.change({ratings:{q1:5},comment:'newer'});const save2=session.save();
  first.resolve(state(calls[0],1));await Promise.all([save1,save2]);
  assert.equal(calls.length,2);assert.equal(calls[1].draft_version,1);assert.equal(calls[1].ratings.q1,5);
  assert.equal(calls[1].comment,'newer');assert.equal(session.dirty,false);
});

test('network errors keep draft and permit a safe retry', async t => {
  let failing=true;
  const api={request:async (p,m,d)=>{if(failing)throw new Error('offline');return state(d,1);}};
  const session=sessionFor(t,api);session.change({ratings:{q1:0},comment:'draft must remain'});
  await assert.rejects(session.save(),/offline/);assert.equal(session.dirty,true);assert.equal(session.data.comment,'draft must remain');
  failing=false;await session.save();assert.equal(session.dirty,false);assert.equal(session.data.ratings.q1,0);
});

test('two-tab conflict keeps typed answers until explicit keep-mine resolution', async t => {
  let calls=0;
  const current=state({ratings:{q1:4},comment:'other device'},5);
  const api={request:async (p,m,d)=>{
    calls++;if(calls===1){const error=new Error('conflict');error.code='draft_conflict';error.state=current;throw error;}
    assert.equal(d.draft_version,5);return state(d,6);
  }};
  const session=sessionFor(t,api);session.change({ratings:{q1:0},comment:'mine'});
  await assert.rejects(session.save());assert.equal(session.data.comment,'mine');assert.equal(session.conflict,current);
  await session.keepMine();assert.equal(session.data.comment,'mine');assert.equal(session.dirty,false);assert.equal(session.version,6);
});

test('explicit use-latest choice restores the newer server draft', async t => {
  const session=sessionFor(t,{request:()=>{throw new Error('should not save');}});
  session.change({comment:'mine'});session.conflict=state({ratings:{q2:0},comment:'latest'},3);
  session.useLatest();assert.equal(session.data.comment,'latest');assert.equal(session.data.ratings.q2,0);assert.equal(session.dirty,false);
});

test('submit is blocked for unanswered questions, long comment; empty comment and zero ratings accepted', async t => {
  let calls=0;
  const session=sessionFor(t,{request:async(p,m,d)=>{calls++;assert.equal(p,'/submit');assert.deepEqual(d.ratings,ratings(0));assert.equal(d.comment,'');return {completed:true};}});
  await assert.rejects(session.submit());session.change({ratings:ratings(0),comment:'x'.repeat(2001)});await assert.rejects(session.submit());assert.equal(calls,0);
  session.change({ratings:ratings(0),comment:''});await session.submit();assert.equal(calls,1);assert.equal(session.completed,true);assert.equal(session.dirty,false);
});

test('double submit only sends one request and already-submitted server reply is successful', async t => {
  const pending=deferred();let calls=0;
  const session=sessionFor(t,{request:async()=>{calls++;return pending.promise;}});
  session.change({ratings:ratings(2)});const submit1=session.submit();await session.submit();assert.equal(calls,1);
  pending.resolve({completed:true,already_submitted:true});await submit1;await session.submit();assert.equal(calls,1);assert.equal(session.completed,true);
});

test('lost submit response does not complete locally and can safely retry the same answers', async t => {
  let calls=0;const payloads=[];
  const session=sessionFor(t,{request:async(p,m,d)=>{calls++;payloads.push(d);if(calls===1)throw new Error('offline');return {completed:true,already_submitted:true};}});
  session.change({ratings:ratings(3),comment:'feedback'});await assert.rejects(session.submit());assert.equal(session.completed,false);assert.equal(session.data.comment,'feedback');
  await session.submit();assert.deepEqual(payloads[0],payloads[1]);assert.equal(session.completed,true);
});

test('another tab completion during draft save becomes completed without resetting to empty form', async t => {
  const session=sessionFor(t,{request:async()=>{const error=new Error('already');error.code='already_submitted';error.state={completed:true};throw error;}});
  session.change({ratings:{q1:2},comment:'retained'});await session.save();assert.equal(session.completed,true);assert.equal(session.data.comment,'retained');
});

test('API uses authenticated same-origin session and PUT keepalive with no user identifier', async () => {
  let options;
  const api=new SurveyAPI(async(url,opts)=>{assert.equal(url,'/api/ux-survey/draft');options=opts;return {ok:true,redirected:false,status:200,json:async()=>state({},1)};});
  await api.request('/draft','PUT',{ratings:ratings(0),comment:'',draft_version:0},{keepalive:true});
  assert.equal(options.credentials,'same-origin');assert.equal(options.cache,'no-store');assert.equal(options.keepalive,true);
  assert.equal(JSON.parse(options.body).user_id,undefined);
});

test('API preserves conflict metadata and gives safe session expiry error', async () => {
  const current=state({comment:'new'},4);
  const api=new SurveyAPI(async()=>({ok:false,status:409,redirected:false,json:async()=>({code:'draft_conflict',state:current})}));
  await assert.rejects(api.request('/draft','PUT',{}),error=>error.code==='draft_conflict' && error.state.draft_version===4);
  const loggedOut=new SurveyAPI(async()=>({ok:true,status:200,redirected:true,json:async()=>{throw new Error('html');}}));
  await assert.rejects(loggedOut.request('/state'),/Сессия завершилась/);
});

test('idle invitation respects focused input, form edits, active dialogs and background tabs', () => {
  global.getComputedStyle=()=>({visibility:'visible'});
  const doc={visibilityState:'visible',hasFocus:()=>true,activeElement:{matches:()=>false,closest:()=>null},querySelectorAll:()=>[]};
  assert.equal(safeToInvite(doc,0,false,16000),true);
  assert.equal(safeToInvite(doc,5000,false,16000),false);
  assert.equal(safeToInvite(doc,0,true,16000),false);
  doc.activeElement.matches=()=>true;assert.equal(safeToInvite(doc,0,false,16000),false);doc.activeElement.matches=()=>false;
  doc.querySelectorAll=()=>[{getClientRects:()=>[1]}];assert.equal(safeToInvite(doc,0,false,16000),false);doc.querySelectorAll=()=>[];
  doc.visibilityState='hidden';assert.equal(safeToInvite(doc,0,false,16000),false);
});

test('comment rendering never uses HTML and stars expose native labelled keyboard inputs', () => {
  const js=fs.readFileSync(path.join(__dirname,'../static/ux-survey.js'),'utf8');
  const html=fs.readFileSync(path.join(__dirname,'../templates/ux_survey.html'),'utf8');
  assert.doesNotMatch(js,/\.innerHTML\s*=/);
  assert.match(js,/comment\.value = data\.comment/);
  assert.match(js,/input\.type = 'radio'/);
  assert.match(js,/input\.setAttribute\('aria-label'/);
  assert.match(js,/event\.preventDefault\(\); defer\(\)/);
  assert.match(js,/\/invitation\/shown/);
  assert.doesNotMatch(html,/maxlength=/);
  assert.match(js,/comment\.setCustomValidity/);
  assert.match(html,/id="uxSurveyFields" disabled/);
});


test('Unicode limit accepts exactly 2000 emoji and keeps 2001 without truncation or a draft request', async t => {
  const payloads=[];const notices=[];
  const api={request:async(p,m,d)=>{payloads.push(d);return state(d,payloads.length);}};
  const session=sessionFor(t,api,state(),event=>notices.push(event));
  const accepted='🙂'.repeat(2000);
  session.change({ratings:ratings(0),comment:accepted});
  await session.save();assert.equal(payloads.length,1);assert.equal(payloads[0].comment,accepted);
  assert.equal(commentLength(accepted),2000);assert.equal(accepted.length,4000);
  const rejected=accepted+'🙂';
  session.change({ratings:ratings(0),comment:rejected});
  assert.equal(session.data.comment,rejected);assert.equal(session.dirty,true);
  await assert.rejects(session.save(),error=>error.validation===true);
  await assert.rejects(session.submit(),error=>error.validation===true);
  assert.equal(payloads.length,1);assert.ok(notices.includes('invalid_comment'));
  session.change({ratings:ratings(0),comment:accepted});await session.save();
  assert.equal(payloads.length,2);assert.equal(session.dirty,false);assert.equal(session.data.comment,accepted);
});


test('invitation ignores hidden modal ancestors but blocks visible pending form submission', () => {
  global.getComputedStyle=()=>({visibility:'visible'});
  const hidden={getClientRects:()=>[]};
  const busy={getClientRects:()=>[1]};
  let submitting=false;
  const doc={visibilityState:'visible',hasFocus:()=>true,activeElement:{matches:()=>false,closest:()=>null},querySelectorAll:selector=>{
    assert.ok(selector.includes('form button[type="submit"]:disabled'));
    return submitting ? [hidden,busy] : [hidden];
  }};
  assert.equal(safeToInvite(doc,0,false,20000),true);
  submitting=true;assert.equal(safeToInvite(doc,0,false,20000),false);
});
