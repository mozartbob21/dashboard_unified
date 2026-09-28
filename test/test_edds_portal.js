// Run with node --test test/test_edds_portal.js. No network or credentials used.
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../services/edds/portal.js'), 'utf8');
const origin = 'https://zkh-kontur.mosreg.ru';
const defaults = {fromDate:'2026-09-17',toDate:'2026-09-24',coordinates:false,timeout:1000,limit:200000};
async function run(responses, options={}) {
  const calls=[];
  const fetch=async (url, init)=>{
    calls.push({url,init});
    const item=responses.shift();
    if(item instanceof Error) throw item;
    if(item instanceof Response) return item;
    return new Response(item,{status:200});
  };
  const context={window:{fetch},location:{origin:options.origin || origin},URL,URLSearchParams,
    TextDecoder,Uint8Array,AbortController,setTimeout,clearTimeout};
  const report=vm.runInNewContext('('+source+')',context);
  const result=await report({...defaults,...options});
  return {result:JSON.parse(JSON.stringify(result)),calls};
}

test('summary uses the archive CSV endpoint, date format and every resource',async()=>{
  const {result,calls}=await run(['\ufeffid_cds_claim;text_message\n7;Прорыв']);
  assert.deepEqual(result.grid,[['id_cds_claim','text_message'],['7','Прорыв']]);
  assert.equal(calls.length,1);
  assert.equal(calls[0].url,origin+'/new9/?act=cds_report_svod&id=3608');
  assert.deepEqual(Object.fromEntries(calls[0].init.body),{
    act:'cds_report_svod',id:'3608',date_ot:'17.09.26',date_do:'24.09.26',id_tu_mun_raion:'0',saveToCSV:'1'});
  assert.equal(calls[0].init.method,'POST');
  assert.equal(calls[0].init.mode,'same-origin');
  assert.equal(calls[0].init.credentials,'same-origin');
});

test('coordinates preserve archive fields, four-digit year and Russian headings',async()=>{
  const {result,calls}=await run(['Номер заявки;Широта;Долгота\n7;55,75;37,61'],{coordinates:true});
  assert.deepEqual(result.grid,[['Номер заявки','Широта','Долгота'],['7','55,75','37,61']]);
  assert.equal(calls[0].url,origin+'/new9/?act=cds_claim_report&id=3654');
  assert.deepEqual(Object.fromEntries(calls[0].init.body),{
    start:'17.09.2026',end:'24.09.2026','fields[id_cds_claim]':'on','fields[lat_]':'on','fields[lon_]':'on',submit:'1'});
});

test('coordinates download follows a same-origin generated CSV link',async()=>{
  const {result,calls}=await run(['<html><a href="exports/EDDS_claim_report_1.csv">CSV</a></html>',
    'id_cds_claim;lat_;lon_\n7;55;37'],{coordinates:true});
  assert.equal(calls.length,2);
  assert.equal(calls[1].url,origin+'/new9/exports/EDDS_claim_report_1.csv');
  assert.deepEqual(result.grid,[['Номер заявки','Широта','Долгота'],['7','55','37']]);
});

test('external CSV links cannot leave the portal origin',async()=>{
  for(const href of ['https://other.test/report.csv','http://zkh-kontur.mosreg.ru/report.csv','https://user:pass@zkh-kontur.mosreg.ru/report.csv']){
    const {result,calls}=await run(['<html><a href="'+href+'">CSV</a></html>'],{coordinates:true});
    assert.equal(result.status,502);assert.match(result.error,/другой адрес/);assert.equal(calls.length,1);
  }
});

test('foreign tab is rejected before making any requests',async()=>{
  const {result,calls}=await run([],{origin:'https://other.test'});
  assert.equal(result.status,409);assert.equal(calls.length,0);
});

test('expired login, permission error and malformed report never become an empty success',async()=>{
  for(const body of ['<html><form action="?act=login">Вход</form></html>',new Response('Denied',{status:403})]){
    const {result}=await run([body]);assert.equal(result.status,403);assert.equal(result.grid,undefined);
  }
  const {result}=await run(['garbage;data\n1;2']);assert.equal(result.status,502);
});

test('archive parser preserves multiline descriptions, embedded quotes and comma delimiters',async()=>{
  const {result}=await run(['id_cds_claim;text_message\n1;"Первая строка\nСНТ ""Сырьево"""\n2;Текст "без кавычек"']);
  assert.deepEqual(result.grid,[['id_cds_claim','text_message'],['1','Первая строка\nСНТ "Сырьево"'],['2','Текст "без кавычек"']]);
  const comma=await run(['id_cds_claim,text_message\n1,"a;b;c;d"']);
  assert.deepEqual(comma.result.grid,[['id_cds_claim','text_message'],['1','a;b;c;d']]);
});

test('UTF-8 and Windows-1251 responses decode inside the browser',async()=>{
  const bytes=Buffer.concat([Buffer.from('id_cds_claim;text_message\n7;'),Buffer.from([0xd2,0xe5,0xf1,0xf2])]);
  const {result}=await run([new Response(bytes)]);
  assert.deepEqual(result.grid[1],['7','Тест']);
});

test('stream and content-length limits bound downloads',async()=>{
  for(const response of [new Response('id_cds_claim;text_message\n7;too long'),
    new Response('small',{headers:{'content-length':'500'}})]){
    const {result}=await run([response],{limit:10});assert.equal(result.status,413);
  }
});

test('network failure returns a safe error without transport details',async()=>{
  const {result}=await run([new Error('token=private')]);
  assert.equal(result.status,502);assert.doesNotMatch(result.error,/private/);
});
