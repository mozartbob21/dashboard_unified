'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const api=require('../static/aichat-agents');
const ID='complaints-neighbors-rating';
const file=name=>new File(['synthetic input'],name);
const formValues=()=>({city:'Чехов',geo:file('справочник.xlsx'),complaints:file('жалобы.xlsx'),bounds:['37.1','37.9','54.95','55.35'],radius:'100',wide:'200'});

test('two files travel under explicit agent roles, no ordinary file upload',()=>{
  const values=formValues();
  const run=api.prepare(values);
  const body=api.bodyFor(run,'dialog-1');
  assert.equal(body.get('agent_id'),ID);
  assert.equal(body.get('dialog_id'),'dialog-1');
  assert.equal(body.get('file'),null);
  assert.equal(body.get('geo_file').name,values.geo.name);
  assert.equal(body.get('complaints_file').name,values.complaints.name);
  assert.equal(body.get('agent_bbox'),'37.1,37.9,54.95,55.35');
  assert.ok(!body.get('text').includes('xlsx'));
  assert.equal(api.sameRun(run,api.prepare(values)),true);
  assert.equal(api.sameRun(run,api.prepare({...values, complaints:file('новая.xlsx')})),false);
});

test('invalid files, reversed coordinates and radii do not enter the chat',()=>{
  for(const override of [{city:' '},{city:'123'},{city:'а'.repeat(81)},{wide:50001},{bounds:['37','','54','55']},{bounds:[38,37,54,55]},
     {bounds:[37,38,55,100]}, {wide:'99'}, {radius:'0'}, {geo:file('map.xlsm')},
     {complaints:{name:'big.xlsx',size:9*1024*1024}}]){
    assert.throws(()=>api.prepare({...formValues(),...override}));
  }
});

test('only server-shaped artifact identity becomes a fixed local XLSX link',()=>{
  const doc={createElement:tag=>({tag})}, id='a'.repeat(32);
  for(const fake of [null,{}, {id:'https://external.invalid/file'}, {id:'../../etc.zip'}, {id:id+'/fake'}]){
    assert.equal(api.downloadLink(doc,fake),null);
  }
  const link=api.downloadLink(doc,{id,url:'https://external.invalid/file',filename:'../../secret'});
  assert.equal(link.href,'/aichat/api/agents/artifacts/'+id);
  assert.equal(link.textContent,'Скачать рейтинг в Excel');
});

function harness(permission) {
  const nodes=new Map(), sent=[];
  class Element {
    constructor(id){this.id=id;this.hidden=true;this.disabled=false;this.events={};this.value='';this.files=[];this.open=false;}
    addEventListener(event,cb){this.events[event]=cb;}
    setAttribute(name,value){this[name]=value;}
    contains(){return false;}
    reset(){populate();nodes.get('agentGeoFile').files=[];nodes.get('agentComplaintsFile').files=[];}
    trigger(ev){return this.events[ev]?.({target:this,preventDefault(){}});}
  }
  const doc={getElementById(id){if(!nodes.has(id))nodes.set(id,new Element(id));return nodes.get(id);}};
  const el=id=>doc.getElementById(id);
  function populate(){const values=formValues();['agentGeoFile','agentComplaintsFile'].forEach((k,i)=>el(k).files=[i?values.complaints:values.geo]);
    ['agentWest','agentEast','agentSouth','agentNorth'].forEach((k,i)=>el(k).value=values.bounds[i]);
    el('agentCity').value='Чехов';el('agentRadius').value='100';el('agentRadiusWide').value='200';}
  populate();
  let sequence=0;
  const ui=api.create({doc,request:async()=>({agents:permission?[{id:ID,title:'Рейтинг',description:'Расстояния'}]:[]}),
    dialog:()=> 'dialog',openModal(){},closeModal(){},key:()=>String(++sequence),isBusy:()=>false,
    submit:async(body,id)=>{sent.push({body,id});return null;}});
  return {el,ui,sent,setPermission(value){permission=value;}};
}

test('access, uncertain submission retry, corrected files and new city boundaries',async()=>{
  const h=harness(true);
  await h.ui.start();
  assert.equal(h.el('agentsBtn').hidden,false);
  assert.equal(h.el('agentRatingTitle').textContent,'Рейтинг');
  await h.el('agentRatingForm').trigger('submit');
  await h.el('agentRatingForm').trigger('submit');
  assert.equal(h.sent.length,2);assert.equal(h.sent[0].id,h.sent[1].id);
  h.el('agentComplaintsFile').files=[file('updated.xlsx')];
  await h.el('agentRatingForm').trigger('submit');
  assert.notEqual(h.sent[0].id,h.sent[2].id);
  h.ui.accept(h.sent[2].id);
  assert.deepEqual(h.el('agentComplaintsFile').files,[]);
  h.el('agentCity').value='Химки';h.el('agentCity').trigger('input');
  assert.equal(h.el('agentWest').value,'');assert.equal(h.el('agentOptions').open,true);
  h.setPermission(false);
  await h.el('agentsRetry').trigger('click');
  assert.equal(h.el('agentsBtn').hidden,true);assert.equal(h.el('agentRatingForm').hidden,true);
});
