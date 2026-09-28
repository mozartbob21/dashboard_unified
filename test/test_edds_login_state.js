'use strict';
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');
const source = readFileSync(join(__dirname,'../services/edds/chrome.py'),'utf8');
function fixture() {
  const nodes={password:[],captcha:[],verification:[],errors:[],entry:[],report:false,meta:null};
  const listeners={};
  const document={readyState:'complete',
    querySelectorAll(selector) {
      if(selector==='input[type="password"]') return nodes.password;
      if(selector.startsWith('iframe')) return nodes.captcha;
      if(selector.startsWith('input[autocomplete=')) return nodes.verification;
      if(selector.startsWith('[role=')) return nodes.errors;
      if(selector.startsWith('button,')) return nodes.entry;
      throw new Error('unexpected selector');
    },
    querySelector(selector) {
      return selector.startsWith('meta[') ? nodes.meta : nodes.report ? {} : null;
    },
    createElement(){ return {dataset:{}}; },
    addEventListener(name,fn){listeners[name]=fn;},
    head:{appendChild(meta){nodes.meta=meta;}},
  };
  const context=vm.createContext({document,window:{},getComputedStyle:()=>({visibility:'visible'})});
  function load(name){return vm.runInContext('('+source.match(new RegExp(name+" = r'''([\\s\\S]*?)'''"))[1]+')',context);}
  return {nodes,listeners,context,state:load('LOGIN_STATE'),guard:load('LOGIN_GUARD')};
}
const node=(text='',visible=true)=>({textContent:text,getClientRects:()=>visible?[{}]:[],getAttribute:()=>null});
test('entry button is detected before the modal and visible password takes precedence',()=>{
  const f=fixture();f.nodes.entry=[node('Авторизоваться')];assert.equal(f.state(),'entry');
  f.nodes.password=[node('',false)];assert.equal(f.state(),'entry');
  f.nodes.password=[node()];assert.equal(f.state(),'login');
});
test('hidden entry and ESIA are not treated as the password login entry',()=>{
  const f=fixture();f.nodes.entry=[node('Авторизоваться',false),node('Войти через Госуслуги'),
    {...node('Войти'),id:'esia-auth-button'}];assert.equal(f.state(),'other');
});
test('entry uses button caption rather than its submitted value, and input value as caption',()=>{
  const f=fixture();f.nodes.entry=[{...node('Авторизоваться'),tagName:'BUTTON',value:'login'}];
  assert.equal(f.state(),'entry');
  f.nodes.entry=[{...node(),tagName:'INPUT',value:'Войти'}];assert.equal(f.state(),'entry');
});
test('AJAX submission is not completed just because a modal hides and entry reappears',()=>{
  const f=fixture();f.nodes.entry=[node('Авторизоваться')];f.guard();assert.equal(f.state(),'');
});
test('password form waits for login and report requires both report fields',()=>{
  const f=fixture();f.nodes.password=[node()];assert.equal(f.state(),'login');
  f.nodes.password=[node('',false)];f.nodes.report=true;assert.equal(f.state(),'report');
});
test('hiding an AJAX login form cannot trigger premature navigation',()=>{
  const f=fixture();f.nodes.password=[node()];f.guard();f.nodes.password=[];
  assert.equal(f.state(),'');
  delete f.context.window.__neuronaAuthAttemptActive;
  assert.equal(f.state(),'other');
});
test('stale validation alert is ignored while newly changed error is detected',()=>{
  const f=fixture();const alert=node('Старая ошибка');f.nodes.password=[node()];f.nodes.errors=[alert];
  assert.equal(f.state(),'login_error');f.guard();assert.equal(f.state(),'login');
  alert.textContent='Неверный пароль';assert.equal(f.state(),'login_error');
});
test('hidden challenge does not block login; visible challenge requires user',()=>{
  const f=fixture();f.nodes.password=[node()];f.nodes.captcha=[node('',false)];assert.equal(f.state(),'login');
  f.nodes.captcha=[node()];assert.equal(f.state(),'captcha');
  f.nodes.captcha=[];f.nodes.verification=[node()];assert.equal(f.state(),'verification');
});
test('policy violations produce a bounded blocked state without collecting payloads',()=>{
  const f=fixture();f.guard();
  assert.equal(f.nodes.meta.content,"form-action 'self'; connect-src 'self'");
  f.listeners.securitypolicyviolation({effectiveDirective:'connect-src'});
  assert.equal(f.state(),'blocked');
});
