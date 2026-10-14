'use strict';
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const html=fs.readFileSync(path.join(__dirname,'..','templates','zip_curator_original.html'),'utf8');
const dateCode=html.slice(html.indexOf('function sourceDateText('),html.indexOf('\nconst fmt='));
const badge={textContent:''},warning={textContent:''},button={textContent:'Обновить',disabled:false,setAttribute(){},removeAttribute(){}};
let source={provider:'mosreg',label:'Облако МО',configured:true};
const calls=[];
const context=vm.createContext({
  $:selector=>({'#sourceBadge':badge,'#cfgWarn':warning,'#btnDrive':button}[selector]),
  fetch:async(url,options)=>{
    calls.push([url,options]);
    if(url.endsWith('/source'))return {ok:source.configured,json:async()=>source};
    assert.equal(url,'/zip_curator/api/scan');assert.equal(options.method,'POST');assert.deepEqual(JSON.parse(options.body),{});
    return {ok:true,status:200,json:async()=>({ok:true,files:1,downloaded:1,state:{pending:[],clean:{}}})};
  },
  toast:()=>{},renderDriveSync:()=>{},applyServerState:()=>{},safeScanError:(_,fallback)=>fallback,
});
vm.runInContext(dateCode,context);
function format(record){context.record=record;return vm.runInContext('sourceDateText(record)',context);}
assert.match(format({source_created_at:'2025-02-03T08:15:00Z',source_modified_at:'2026-10-13T10:00:00Z',loaded_at:'2026-10-14T11:00:00Z'}),/^Создан в источнике:.*2025/);
assert.match(format({source_modified_at:'2026-10-13T10:00:00Z'}),/^Изменён в источнике:/);
assert.equal(format({uploaded:1700000000000,loaded_at:'2026-10-14T11:00:00Z'}),'Дата источника неизвестна');
assert.equal(format({source_created_at:'invalid'}),'Дата источника неизвестна');
assert.equal(format({source_created_at:'2026-02-30T12:00:00Z'}),'Дата источника неизвестна');
assert.doesNotMatch(html,/YANDEX_FOLDER_A_URL/);
assert.match(html,/<th>Дата файла в источнике<\/th>/);
const scanCode=html.slice(html.indexOf('async function fetchFromDrive(){'),html.indexOf('// сравнение дат вида'));
vm.runInContext('let driveSyncRunning=false;\n'+scanCode,context);
(async()=>{
  await vm.runInContext('loadSource()',context);assert.equal(badge.textContent,'Источник: Облако МО');
  await vm.runInContext('fetchFromDrive()',context);assert.equal(button.disabled,false);
  source={provider:'yandex',label:'Яндекс Диск',configured:true};
  await vm.runInContext('loadSource()',context);assert.equal(badge.textContent,'Источник: Яндекс Диск');
  source={configured:false};await vm.runInContext('loadSource()',context);assert.match(warning.textContent,/вручную/);
  assert.equal(button.disabled,false);
  console.log('ZIP source-date labels, source badge and server-selected scan tests passed');
})().catch(error=>{console.error(error);process.exitCode=1;});
