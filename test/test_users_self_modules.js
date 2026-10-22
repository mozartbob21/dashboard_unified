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
