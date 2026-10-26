'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const {accountEditPolicy,accountSaveRequest}=require('../static/users.js');
const delegate={id:2,can_manage_users:true};

test('delegate opens own module editor while other managers remain protected',()=>{
  assert.deepEqual(accountEditPolicy(delegate,2,false),{self:true,ownModulesOnly:true,protectedUser:false});
  assert.deepEqual(accountEditPolicy({id:3,can_manage_users:true},2,false),{self:false,ownModulesOnly:false,protectedUser:true});
});
test('own module submit contains no identity, status or delegation fields',()=>{
  const payload={username:'delegate',password:'must-not-submit',email:'example@example.invalid',is_active:false,can_manage_users:true,modules:['zips']};
  assert.deepEqual(accountSaveRequest(delegate,2,false,payload),{
    url:'/api/users/me/modules',method:'PUT',payload:{modules:['zips']}
  });
});
test('parent keeps existing editor and may assign own modules',()=>{
  const user={id:1,can_manage_users:true,is_manager:true}, payload={username:'parent',modules:['zips']};
  assert.deepEqual(accountEditPolicy(user,1,true),{self:true,ownModulesOnly:false,protectedUser:false});
  assert.deepEqual(accountSaveRequest(user,1,true,payload),{url:'/api/users/1',method:'PUT',payload});
});
test('new and ordinary accounts keep the established save routes',()=>{
  const payload={username:'ordinary',modules:['edo']};
  assert.deepEqual(accountSaveRequest(null,2,false,payload),{url:'/api/users',method:'POST',payload});
  assert.deepEqual(accountSaveRequest({id:4},2,false,payload),{url:'/api/users/4',method:'PUT',payload});
});
test('unknown identity does not enable self editing',()=>{
  for(const id of [undefined,null,'2',false]) assert.equal(accountEditPolicy(delegate,id,false).protectedUser,true);
});

const {aiTokenItems}=require('../static/users.js');
test('AI token settings reject missing or malformed response with a readable message',()=>{
  for(const data of [null,{}, {items:null}, {items:{}}, {items:[]}, {items:[null]}, {items:[{module_id:'aichat'}]},
    {items:[{module_id:'aichat',label:'Чат'},{module_id:'aichat',label:'Другой'}]}]){
    assert.throws(()=>aiTokenItems(data),/Сервер вернул некорректный список настроек ИИ/);
  }
});
test('AI token settings accept the complete list before rendering and copy its array',()=>{
  const data={items:[{module_id:'aichat',label:'ИИ-чат',source:'saved'},{module_id:'tools',label:'Инструменты',source:'none'}]};
  const items=aiTokenItems(data);
  assert.deepEqual(items,data.items); assert.notEqual(items,data.items);
});
