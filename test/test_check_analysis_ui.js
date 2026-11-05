const test=require('node:test');
const assert=require('node:assert/strict');
const {pageScope,links,projection,createController}=require('../static/check-analysis.js');

test('read-only analysis routes keep distinct checks under the same permission separate',()=>{
  assert.deepEqual(pageScope('/camera-prescriptions/'),['camera_prescriptions']);
  assert.deepEqual(pageScope('/cameras'),['cameras']);
  assert.deepEqual(pageScope('/mingkh/treatment'),['mingkh','treatment']);
  assert.deepEqual(pageScope('/mingkh/collective'),['collective']);
  assert.deepEqual(pageScope('/mingkh/sms/system'),['mingkh','sms']);
  assert.deepEqual(pageScope('/zips'),['zips']);
  assert.deepEqual(pageScope('/zip_curator_original.html'),['zips','zips_scan']);
  assert.deepEqual(pageScope('/edds'),['edds']);
  assert.deepEqual(pageScope('/tools'),['tools']);
  assert.equal(pageScope('/login'),null);
  assert.equal(pageScope('/mingkh/api/dataset'),null);
  assert.equal(pageScope('/appeals/12'),null);
  assert.deepEqual(links('mingkh','treatment'),{
    api:'/api/check-history/mingkh/analysis?operation=treatment',
    chat:'/aichat?history=mingkh&operation=treatment'
  });
});

test('render limits metrics and keeps a real zero; missing history does not acquire a date',()=>{
  const data=projection({summary:'Нет сохранённых проверок.',metrics:[]});
  assert.equal(data.date,'');
  assert.deepEqual(data.metrics,[]);
  const metric=(value)=>({label:'Метрика',value,unit:'шт.'});
  const recorded=projection({summary:'Проверка завершилась ошибкой',status:'error',latest_at:'invalid',
    metrics:[metric(null), metric(0), metric(3),metric(4),metric(5)]});
  assert.equal(recorded.status,'error');
  assert.equal(recorded.date,'');
  assert.deepEqual(recorded.metrics.map(x=>x.value),[0,3,4]);
});

test('lost network retains previous result with a warning; revoked permission clears it',async()=>{
  let calls=0, value, notice;
  const read=async()=>{
    calls++;
    if(calls===1) return {ok:true,json:async()=>({summary:'Прочитаны цифры',metrics:[]})};
    if(calls===2) throw new Error('disconnect');
    return {status:403};
  };
  const controller=createController({read,visible:()=>true,render:data=>value=data.summary,
    reset:()=>value=null,notice:text=>notice=text});
  await controller.refresh();
  await controller.refresh();
  assert.equal(value,'Прочитаны цифры'); assert.match(notice,/предыдущий результат/);
  await controller.refresh();
  assert.equal(value,null); assert.match(notice,/Нет доступа/);
});

test('hidden cards do not poll; concurrent refresh and inflight GET are coalesced',async()=>{
  let visible=false,calls=0,done;
  const controller=createController({visible:()=>visible, read:async()=>{
    ++calls; await new Promise(resolve=>done=resolve); return {ok:true,json:async()=>({summary:'Нет проверок.'})};
  }, render:()=>{},notice:()=>{},reset:()=>{}});
  await controller.refresh(); assert.equal(calls,0);
  visible=true;
  const first=controller.refresh(); await controller.refresh(); assert.equal(calls,1);
  done(); await first;
});
