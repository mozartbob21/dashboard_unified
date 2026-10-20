const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const script = fs.readFileSync(require.resolve('../static/water-dashboard.js'), 'utf8');

function render(sources) {
  const elements = new Map();
  function element(id) {
    if (!elements.has(id)) elements.set(id, {innerHTML:'', textContent:'', open:false,
      addEventListener(){}, classList:{toggle(){}}, querySelector(){return element(id+' child');}});
    return elements.get(id);
  }
  element('waterSnapshot').textContent = JSON.stringify({metric_schema:1, sources, table:[], kpis:{}});
  const context = {
    document:{getElementById:element, querySelector:element, querySelectorAll:()=>[]},
    // Polling is suspended: this is a deterministic frontend contract check.
    fetch:()=>new Promise(()=>{}), setInterval:()=>0, Date, Set, URL,
    location:{href:'http://localhost/water-dashboard'}
  };
  vm.runInNewContext(script, context);
  return element('kpiGrid').innerHTML;
}
const metric = (id,label,value,unit='')=>({id,label,value,unit});
function source(metrics) { return {metric_schema:1,ok:true,metrics,updated_at:'2099-01-01'}; }
function article(html,id) {
  const found = html.match(new RegExp('<article[^>]+data-source="'+id+'">([\\s\\S]*?)</article>'));
  assert.ok(found,'expected source card: '+id); return found[1];
}
function primary(html) { return html.slice(html.indexOf('<div class="wd-primary-metrics'), html.indexOf('<div class="wd-rows">')); }

test('water and capital repair cards show both semantic address totals with equal classes',()=>{
  for (const id of ['sys_vs','sys_kr']) {
    const html=article(render({[id]:source([
      metric('resonant_addresses','Резонансные адреса',25,'адр.'),
      metric('system_addresses','Системные адреса',9,'адр.')
    ])}),id), head=primary(html);
    assert.match(head,/wd-primary-metrics--paired/);
    assert.equal((head.match(/class="wd-num"/g)||[]).length,2);
    assert.ok(head.indexOf('data-metric="system_addresses"')<head.indexOf('data-metric="resonant_addresses"'));
    assert.match(head,/>9 адр\.</);assert.match(head,/>25 адр\.</);
    assert.doesNotMatch(html.slice(html.indexOf('<div class="wd-rows">')),/9 адр\.|25 адр\./);
  }
});

test('EDO promotes electronic document share, preserving signature indicators as secondary',()=>{
  const html=article(render({edo_rso:source([
    metric('signers_ecp_pct','Подписанты с ЭЦП',80,'%'),
    metric('signers_total','Должностные лица с правом подписи',20,'чел.'),
    metric('signers_ecp','Подписанты с ЭЦП',16,'чел.'),
    metric('electronic_documents_pct','Электронные документы в общем количестве',32,'%')
  ])}),'edo_rso');
  assert.match(primary(html),/data-metric="electronic_documents_pct"/);
  assert.match(primary(html),/>32 %</);
  assert.doesNotMatch(primary(html),/>80 %</);
  assert.match(html.slice(html.indexOf('<div class="wd-rows">')),/>80 %</);
  assert.equal((html.match(/>32 %</g)||[]).length,1);
});

test('missing document share is never replaced by signature percentage; zero remains a value',()=>{
  const missing=article(render({edo_rso:source([metric('signers_ecp_pct','Подписанты с ЭЦП',90,'%')])}),'edo_rso');
  assert.match(primary(missing),/class="wd-num">—</);assert.doesNotMatch(primary(missing),/>90 %</);
  const zero=article(render({edo_rso:source([metric('electronic_documents_pct','Доля',0,'%')])}),'edo_rso');
  assert.match(primary(zero),/>0 %</);
  const addresses=article(render({sys_vs:source([metric('system_addresses','Системные адреса',0,'адр.')])}),'sys_vs');
  assert.match(primary(addresses),/>0 адр\.</);assert.match(primary(addresses),/class="wd-num">—</);
});

test('unverified data cannot populate primary cards and source labels are escaped',()=>{
  const html=render({edo_rso:{metrics:[metric('electronic_documents_pct','Доля',99,'%')]},
    sys_kr:source([metric('system_addresses','<img src=x onerror=alert(1)>',7,'адр.')])});
  assert.doesNotMatch(primary(article(html,'edo_rso')),/>99/);
  assert.doesNotMatch(html,/<img src=x/);assert.match(html,/&lt;img/);
});

test('other source cards retain their leading metric and additional indicators',()=>{
  const html=article(render({valves:source([metric('inserted','Внесено',10,'шт.'),metric('plan','План',20,'шт.')])}),'valves');
  assert.match(primary(html),/data-metric="inserted"/);assert.match(primary(html),/>10 шт\.</);
  assert.match(html.slice(html.indexOf('<div class="wd-rows">')),/>20 шт\.</);
});
