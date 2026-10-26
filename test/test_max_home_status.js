'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {MaxHomeStatus} = require('../static/max-home-status.js');
function fixture(active) {
  const listeners = {}, calls = [];
  const labels = [{textContent:''}, {textContent:''}];
  const links = labels.map(label => ({dataset:{}, querySelector:()=>label, setAttribute(key,value){this[key]=value;}}));
  const win = {document:{hidden:false, querySelectorAll:()=>links, addEventListener:(name,fn)=>listeners[name]=fn},
    setTimeout:()=>1,clearTimeout(){},setInterval:()=>2,clearInterval(){},addEventListener:(name,fn)=>listeners[name]=fn,
    fetch:async (url, options)=>{calls.push({url,options});return {ok:true,redirected:false,json:async()=>({active,enabled:true})};}};
  return {win,links,labels,listeners,calls,status:new MaxHomeStatus(win)};
}
test('connected label is read from account server state on every new session; no browser storage', async()=>{
  const first=fixture(true);await first.status.refresh(true);
  assert.ok(first.labels.every(label=>label.textContent==='Подключено'));
  assert.equal(first.calls[0].url,'/max-bot/api/status');
  assert.equal(first.calls[0].options.credentials,'same-origin');assert.equal(first.calls[0].options.cache,'no-store');
  assert.equal(first.links[0]['aria-label'],'MAX подключён. Управление подключением');
  const another=fixture(true);await another.status.refresh(true);assert.equal(another.labels[0].textContent,'Подключено');
  const otherAccount=fixture(false);await otherAccount.status.refresh(true);assert.equal(otherAccount.labels[0].textContent,'Подключить MAX');
});
test('server disconnect and unavailable/forbidden status never claim an active connection',async()=>{
  const f=fixture(true);await f.status.refresh(true);
  f.win.fetch=async()=>({ok:true,json:async()=>({enabled:true,active:false})});await f.status.refresh(true);assert.equal(f.labels[0].textContent,'Подключить MAX');
  for(const response of [{ok:false,status:403},{ok:true,redirected:true},{ok:true,json:async()=>({})}]) {
    f.win.fetch=async()=>response;await f.status.refresh(true);assert.equal(f.labels[0].textContent,'MAX');assert.equal(f.links[0].dataset.maxState,'unknown');
  }
});
test('background tabs and overlapping refreshes do not generate duplicate requests',async()=>{
  const f=fixture(true);f.win.document.hidden=true;await f.status.refresh(true);assert.equal(f.calls.length,0);
  f.win.document.hidden=false;let finish;f.win.fetch=()=>new Promise(resolve=>{finish=resolve;});
  const first=f.status.refresh(true);await f.status.refresh(true);assert.equal(f.status.pending,true);
  finish({ok:true,json:async()=>({active:true,enabled:true})});await first;assert.equal(f.status.pending,false);assert.equal(f.labels[0].textContent,'Подключено');
});
test('absent access entry makes no status request',async()=>{const f=fixture(true);f.status.links=[];f.status.init();await f.status.refresh(true);assert.equal(f.calls.length,0);});
