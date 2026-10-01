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

function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => {resolve=yes; reject=no;});
  return {promise, resolve, reject};
}
const flushPromises = () => new Promise(resolve => setImmediate(resolve));
const plain = value => JSON.parse(JSON.stringify(value));
const incidentHeader = ['id_cds_claim','name_mr','ispolnitel','d_create','d_doklad','type_otkl','water_flag'];
const incidentGrid = id => [incidentHeader,
  [String(id),'Тестовый округ','РСО','01.10.2026 07:00','01.10.2026 08:00','Аварийная заявка','1']];

function liveFixture() {
  const requests=[], notices=[], renders=[];
  const nodes={'lv-go':{disabled:false}, eddsArmRefresh:{disabled:false}, err:{hidden:true,textContent:''}};
  const previous={recs:[{id:'previous'}],spans:[{from:'2026-09-01',to:'2026-09-02'}]};
  let coordinates=0;
  const context=vm.createContext({
    LIVE:true, EMBEDDED:true, DATA:previous,
    req:{from:'2026-10-01',to:'2026-10-02'}, liveBusy:false, livePending:false, liveTimer:null, liveCancelled:false, cancelArmReports:()=>{},
    S:v=>String(v??'').trim(), WATER:[['water_flag','ВЗУ']], PRICH:[], NORM_H:72,
    document:{getElementById:id=>nodes[id]}, clearTimeout:()=>{}, siteDay:value=>value,
    dayStart:value=>new Date(value+'T00:00:00'), dayEnd:value=>new Date(value+'T23:59:59.999'),
    fmtN:String, liveSay:(text,kind)=>notices.push({text,kind}),
    fetchReportCsv:(from,to)=>{
      const pending=deferred();requests.push({from,to,...pending});return pending.promise;
    },
    fetchReport:()=>{throw new Error('Embedded requests must not use the legacy fallback');},
    showData:keepFilters=>renders.push({keepFilters,data:plain(context.DATA)}),
    geoAuto:()=>{coordinates++;},
  });
  // Use the real report parser: a valid header-only export signals an empty
  // period, whereas missing columns are an error and must preserve old data.
  vm.runInContext(source.slice(source.indexOf('function parseDT('),source.indexOf('let DATA = null;')),context);
  vm.runInContext(source.slice(source.indexOf('async function loadLive(){'),source.indexOf('async function load(file){')),context);
  return {context,requests,notices,renders,nodes,previous,get coordinates(){return coordinates;}};
}

test('changing dates while loading requests only the latest period and never renders the stale response', async () => {
  const f=liveFixture();
  const first=f.context.loadLive();
  assert.equal(f.requests.length,1);
  assert.equal(f.nodes['lv-go'].disabled,true);
  f.context.req={from:'2026-10-03',to:'2026-10-04'};
  await f.context.loadLive();
  f.context.req={from:'2026-10-05',to:'2026-10-06'};
  await f.context.loadLive();
  assert.equal(f.requests.length,1,'date changes must not overlap requests');
  f.requests[0].resolve(incidentGrid(1));
  await first;await flushPromises();
  assert.equal(f.requests.length,2);
  assert.deepEqual([f.requests[1].from,f.requests[1].to],['2026-10-05','2026-10-06']);
  assert.equal(f.context.DATA,f.previous);
  assert.equal(f.renders.length,0);
  assert.equal(f.coordinates,0);
  assert.equal(f.nodes.eddsArmRefresh.disabled,true);
  f.requests[1].resolve(incidentGrid(3));
  await flushPromises();
  assert.equal(f.context.DATA.recs[0].id,'3');
  assert.deepEqual(plain(f.context.DATA.spans),[{from:'2026-10-05',to:'2026-10-06'}]);
  assert.equal(f.renders.length,1);
  assert.equal(f.renders[0].keepFilters,true);
  assert.equal(f.coordinates,1);
  assert.equal(f.context.liveBusy,false);
  assert.equal(f.nodes['lv-go'].disabled,false);
  assert.equal(f.nodes.eddsArmRefresh.disabled,false);
});

test('a stale request failure neither replaces data nor shows an error for the newly selected dates', async () => {
  const f=liveFixture();
  const first=f.context.loadLive();
  // Simulate an edit before the debounce calls loadLive again. The completed
  // request itself must notice the changed dates and schedule their load.
  f.context.req={from:'2026-10-07',to:'2026-10-08'};
  f.requests[0].reject(new Error('failure from the old period'));
  await first;await flushPromises();
  assert.equal(f.context.DATA,f.previous);
  assert.equal(f.requests.length,2);
  assert.deepEqual([f.requests[1].from,f.requests[1].to],['2026-10-07','2026-10-08']);
  assert.equal(f.renders.length,0);
  assert.equal(f.nodes.err.hidden,true);
  assert.equal(f.notices.some(item=>item.text.includes('failure from the old period')),false);
  f.requests[1].resolve(incidentGrid(2));
  await flushPromises();
  assert.equal(f.context.DATA.recs[0].id,'2');
  assert.deepEqual(plain(f.context.DATA.spans),[{from:'2026-10-07',to:'2026-10-08'}]);
  assert.equal(f.context.liveBusy,false);
});

test('a valid empty report clears prior incidents and records the requested coverage', async () => {
  const f=liveFixture();
  const pending=f.context.loadLive();
  f.requests[0].resolve([incidentHeader]);
  await pending;
  assert.notEqual(f.context.DATA,f.previous);
  assert.deepEqual(plain(f.context.DATA.recs),[]);
  assert.equal(f.context.DATA.total,0);
  assert.equal(f.context.DATA.broken,0);
  assert.deepEqual(plain(f.context.DATA.spans),[{from:'2026-10-01',to:'2026-10-02'}]);
  assert.equal(f.context.DATA.from.getTime(),new Date('2026-10-01T00:00:00').getTime());
  assert.equal(f.context.DATA.to.getTime(),new Date('2026-10-02T23:59:59.999').getTime());
  assert.equal(f.renders.length,1);
  assert.equal(f.renders[0].keepFilters,true);
  assert.equal(f.coordinates,0);
  assert.match(f.notices.at(-1).text,/0 заявок/);
  assert.equal(f.nodes.err.hidden,true);
});

test('a malformed report preserves prior incidents instead of presenting an empty period', async () => {
  const f=liveFixture();
  const pending=f.context.loadLive();
  f.requests[0].resolve([['unexpected_column']]);
  await pending;
  assert.equal(f.context.DATA,f.previous);
  assert.equal(f.renders.length,0);
  assert.equal(f.coordinates,0);
  assert.equal(f.notices.at(-1).kind,'bad');
  assert.match(f.notices.at(-1).text,/Показан прежний период/);
  assert.equal(f.nodes['lv-go'].disabled,false);
});

function armFixture() {
  const requests=[], nodes={'arm-progress':{textContent:''},'arm-cancel':{hidden:true}};
  const context=vm.createContext({URLSearchParams,AbortController,TextEncoder,
    document:{getElementById:id=>nodes[id]},
    setTimeout:(fn,ms)=>{const timer=setTimeout(fn,ms);timer.unref();return timer;},clearTimeout,
    setInterval:(fn,ms)=>{const timer=setInterval(fn,ms);timer.unref();return timer;},clearInterval,
    fetch:(url,options)=>{
      const pending=deferred();requests.push({url,options,...pending});
      options.signal.addEventListener('abort',()=>pending.reject(Object.assign(new Error('Aborted'),{name:'AbortError'})),{once:true});
      return pending.promise;
    }});
  vm.runInContext(source.slice(source.indexOf('let armQueue = Promise.resolve();'),
    source.indexOf('async function fetchReportCsv(')),context);
  return {context,requests,nodes};
}
const reportResponse = grid => ({ok:true,status:200,redirected:false,json:async()=>({grid})});

test('incident and coordinate exports run serially through response parsing', async () => {
  const f=armFixture();
  const incidents=f.context.armReport('01.10.26','02.10.2026');
  const coordinates=f.context.armReport('2026-10-01','2026-10-02',true);
  await flushPromises();
  assert.equal(f.requests.length,1);
  const firstUrl=new URL(f.requests[0].url,'https://dashboard.invalid');
  assert.equal(firstUrl.pathname,'/edds/arm/report');
  assert.equal(firstUrl.searchParams.get('from_date'),'2026-10-01');
  assert.equal(firstUrl.searchParams.get('to_date'),'2026-10-02');
  assert.equal(firstUrl.searchParams.get('coordinates'),'false');
  const body=deferred();
  f.requests[0].resolve({ok:true,status:200,redirected:false,json:()=>body.promise});
  await flushPromises();
  assert.equal(f.requests.length,1,'the next export must wait until JSON parsing finishes');
  body.resolve({grid:incidentGrid(1)});
  assert.deepEqual(await incidents,incidentGrid(1));
  await flushPromises();
  assert.equal(f.requests.length,2);
  const secondUrl=new URL(f.requests[1].url,'https://dashboard.invalid');
  assert.equal(secondUrl.searchParams.get('coordinates'),'true');
  const points=[['id_cds_claim','latitude','longitude'],['1','55.7','37.6']];
  f.requests[1].resolve(reportResponse(points));
  assert.deepEqual(await coordinates,points);
});

test('a failed incident export rejects its caller but does not block queued coordinates', async () => {
  const f=armFixture();
  const incidents=f.context.armReport('2026-10-01','2026-10-02');
  const failed=assert.rejects(incidents,/report temporarily unavailable/);
  const coordinates=f.context.armReport('2026-10-01','2026-10-02',true);
  await flushPromises();
  assert.equal(f.requests.length,1);
  f.requests[0].resolve({ok:false,status:502,redirected:false,
    json:async()=>({detail:'report temporarily unavailable'})});
  await failed;await flushPromises();
  assert.equal(f.requests.length,2);
  f.requests[1].resolve(reportResponse([['id_cds_claim','latitude','longitude']]));
  assert.deepEqual(await coordinates,[['id_cds_claim','latitude','longitude']]);
});


test('malformed rows and missing water fields never clear the previously loaded map', async () => {
  for (const grid of [
    [incidentGrid(1)[0], ['broken', 'Округ', 'РСО', 'invalid-date', '', '', '1']],
    [incidentGrid(1)[0].filter(name=>name!=='water_flag')],
  ]) {
    const f=liveFixture(), pending=f.context.loadLive();
    f.requests[0].resolve(grid);
    await pending;
    assert.equal(f.context.DATA, f.previous);
    assert.equal(f.renders.length, 0);
    assert.ok(f.notices.some(n=>n.kind==='bad' && n.text.includes('Показан прежний период')));
  }
});


test('long periods show completed days and use contiguous inclusive week requests', async () => {
  const f=armFixture();
  const result=f.context.armReport('2026-01-01','2026-01-16');
  await flushPromises();
  const ranges=[];
  for(let i=0;i<3;i++){
    const url=new URL(f.requests[i].url,'https://dashboard.invalid');
    ranges.push([url.searchParams.get('from_date'),url.searchParams.get('to_date')]);
    f.requests[i].resolve(reportResponse([['id_cds_claim'],[String(i+1)]]));
    await flushPromises();
    if(i===0)assert.match(f.nodes['arm-progress'].textContent,/7 из 16 дн/);
  }
  assert.deepEqual(ranges,[['2026-01-01','2026-01-07'],['2026-01-08','2026-01-14'],['2026-01-15','2026-01-16']]);
  assert.deepEqual(plain(await result),[['id_cds_claim'],['1'],['2'],['3']]);
  assert.match(f.nodes['arm-progress'].textContent,/загружено 16 дн/);
  assert.equal(f.nodes['arm-cancel'].hidden,true);
});

test('cancelling a stalled period frees the queue for the latest date and skips stale queued coordinates', async () => {
  const f=armFixture();
  const first=f.context.armReport('2026-01-01','2026-10-01');
  const failed=assert.rejects(first,/отменена/);
  const staleCoordinates=f.context.armReport('2026-01-01','2026-10-01',true);
  const skipped=assert.rejects(staleCoordinates,/отменена/);
  await flushPromises();
  f.context.cancelArmReports();
  const latest=f.context.armReport('2026-10-01','2026-10-01');
  await failed;await skipped;await flushPromises();
  assert.equal(f.requests.length,2);
  assert.equal(new URL(f.requests[1].url,'https://dashboard.invalid').searchParams.get('from_date'),'2026-10-01');
  assert.equal(new URL(f.requests[1].url,'https://dashboard.invalid').searchParams.get('coordinates'),'false');
  f.requests[1].resolve(reportResponse([['id_cds_claim'],['latest']]));
  assert.deepEqual(plain(await latest),[['id_cds_claim'],['latest']]);
});

test('overall deadline aborts a stuck fetch with a clear timeout and releases the queue', async () => {
  const f=armFixture();
  vm.runInContext('Date.now = () => 1000',f.context);
  const pending=f.context.armReport('2026-01-01','2026-10-01');
  const failed=assert.rejects(pending,/лимиту 10 минут/);
  await flushPromises();
  vm.runInContext('Date.now = () => 602000; for(const task of armTasks)task.controller.abort()',f.context);
  await failed;
  const next=f.context.armReport('2026-10-01','2026-10-01');
  await flushPromises();
  assert.equal(f.requests.length,2);
  f.requests[1].resolve(reportResponse([['id_cds_claim']]));
  await next;
});


test('temporary server busy is retried for the same window without overlapping exports', async () => {
  const f=armFixture();
  const pending=f.context.armReport('2026-10-01','2026-10-01');
  await flushPromises();
  f.requests[0].resolve({ok:false,status:409,redirected:false,json:async()=>({detail:'Previous export running',code:'busy'})});
  await flushPromises();
  assert.match(f.nodes['arm-progress'].textContent,/ждём завершения/);
  assert.equal(f.requests.length,1);
  await new Promise(resolve=>setTimeout(resolve,1010));
  assert.equal(f.requests.length,2);
  assert.equal(f.requests[0].url,f.requests[1].url);
  f.requests[1].resolve(reportResponse([['id_cds_claim'],['1']]));
  assert.deepEqual(plain(await pending),[['id_cds_claim'],['1']]);
});

test('a stuck server worker ends busy retries with an actionable error', async () => {
  const f=armFixture();
  vm.runInContext('Date.now=()=>1000',f.context);
  const pending=f.context.armReport('2026-10-01','2026-10-01');
  const failed=assert.rejects(pending,/Проверьте окно Chromium-GOST/);
  await flushPromises();
  const busy={ok:false,status:409,redirected:false,json:async()=>({detail:'Busy',code:'busy'})};
  f.requests[0].resolve(busy);
  await flushPromises();
  vm.runInContext('Date.now=()=>152000',f.context);
  await new Promise(resolve=>setTimeout(resolve,1010));
  f.requests[1].resolve(busy);
  await failed;
  assert.equal(f.requests.length,2);
  assert.equal(f.nodes['arm-cancel'].hidden,true);
});


test('CAPTCHA and other non-busy conflicts are surfaced immediately without retry', async () => {
  const f=armFixture();
  const pending=f.context.armReport('2026-10-01','2026-10-01');
  const failed=assert.rejects(pending,/Complete CAPTCHA/);
  await flushPromises();
  f.requests[0].resolve({ok:false,status:409,redirected:false,json:async()=>({detail:'Complete CAPTCHA in Chromium-GOST'})});
  await failed;await flushPromises();
  assert.equal(f.requests.length,1);
  assert.equal(f.nodes['arm-cancel'].hidden,true);
});


test('overall deadline during a busy retry pause reports the limit and releases the queue', async () => {
  const f=armFixture();
  vm.runInContext('Date.now = () => 1000',f.context);
  const pending=f.context.armReport('2026-01-01','2026-10-01');
  const failed=assert.rejects(pending,/лимиту 10 минут/);
  await flushPromises();
  f.requests[0].resolve({ok:false,status:409,redirected:false,json:async()=>({detail:'Busy',code:'busy'})});
  await flushPromises();
  assert.match(f.nodes['arm-progress'].textContent,/ждём завершения/);
  vm.runInContext('Date.now = () => 602000; for(const task of armTasks)task.controller.abort()',f.context);
  await failed;
  assert.equal(f.requests.length,1);
  assert.equal(f.nodes['arm-cancel'].hidden,true);
  assert.match(f.nodes['arm-progress'].textContent,/Новые данные не применены/);
  const next=f.context.armReport('2026-10-01','2026-10-01');
  await flushPromises();
  f.requests[1].resolve(reportResponse([['id_cds_claim']]));
  await next;
});
