// Run with: node --test test/test_zip_curator_dictionary_frontend.js
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root=process.env.ZIP_CURATOR_UI_ROOT || path.join(__dirname,'..');
const script=fs.readFileSync(path.join(root,'static/zip-curator-dictionary.js'),'utf8');
const template=fs.readFileSync(path.join(root,'templates/zip_curator_original.html'),'utf8');
const initial={dict:{'старая позиция':['Старая категория','Старая группа',0]},categories:['Старая категория'],groupsByCat:{'Старая категория':['Старая группа']},kws:[],units:['шт'],umap:{'шт':'шт'}};
const imported={...initial,dict:{...initial.dict,'новая позиция':['Новая категория','Новая группа',1]},categories:['Старая категория','Новая категория'],groupsByCat:{...initial.groupsByCat,'Новая категория':['Новая группа']}};
const clone=value=>JSON.parse(JSON.stringify(value));
const flush=()=>new Promise(resolve=>setImmediate(resolve));
function deferred(){let resolve,reject;const promise=new Promise((yes,no)=>{resolve=yes;reject=no;});return {promise,resolve,reject};}
class Element{
  constructor(id){this.id=id;this.textContent=id;this.disabled=false;this.value='';this.attributes={};this.listeners={};}
  setAttribute(name,value){this.attributes[name]=value;}
  removeAttribute(name){delete this.attributes[name];}
  addEventListener(name,handler){this.listeners[name]=handler;}
  replaceChildren(...children){this.children=children;}
  reset(){}
  showModal(){}
  close(){}
  focus(){}
}
function fixture(){
  const requests=[],renders=[],toasts=[],downloads=[],nodes={};
  const document={getElementById:id=>nodes[id] ||= new Element(id),createElement:tag=>new Element(tag)};
  const D=clone(initial);
  const context=vm.createContext({D,dictDirty:false,PENDING:[],document,
    toast:message=>toasts.push(message),render:()=>renders.push(Object.keys(D.dict).length),
    toggleView:()=>{},persistEdit:async()=>true,
    FormData:class{constructor(){this.parts=[];}append(name,file){this.parts.push([name,file]);}},
    XLSX:{utils:{aoa_to_sheet:rows=>rows,book_new:()=>({}),book_append_sheet:(book,sheet)=>{book.rows=sheet;}},writeFile:(book,name)=>downloads.push({rows:clone(book.rows),name})},
    fetch:(url,options)=>{const pending=deferred();requests.push({url,options,...pending});return pending.promise;}});
  vm.runInContext(script,context);
  return {context,D,requests,renders,toasts,downloads,nodes,document};
}
const response=(body,status=200)=>({ok:status>=200&&status<300,status,json:async()=>body});
async function initialized(){const f=fixture();f.requests[0].resolve(response(clone(initial)));await flush();return f;}

test('the embedded scripts and dictionary controller compile',()=>{
  new vm.Script(script);
  for(const match of template.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))new vm.Script(match[1]);
});

test('initial server dictionary refresh also renders its persisted record count',async()=>{
  const f=fixture();assert.equal(f.renders.length,0);
  f.requests[0].resolve(response(clone(imported)));await flush();
  assert.deepEqual(f.D.dict,imported.dict);assert.deepEqual(f.renders,[2]);
});

test('import sends the selected workbook and announces success only after durable server acknowledgement',async()=>{
  const f=await initialized();const file={name:'справочник.xlsx'};
  const pending=f.context.importCuratorDictionary(file);
  const request=f.requests[1];assert.equal(request.url,'/zip_curator/api/dictionary/import');
  assert.equal(request.options.method,'POST');assert.deepEqual(request.options.body.parts,[['file',file]]);
  assert.equal(request.options.headers,undefined,'the browser must set the multipart boundary');
  assert.deepEqual(f.D.dict,initial.dict);assert.equal(f.toasts.length,0);
  assert.equal(f.nodes.btnLoadDict.disabled,true);
  request.resolve(response({ok:true,imported:1,dictionary:clone(imported)}));
  assert.equal(await pending,true);assert.deepEqual(f.D.dict,imported.dict);
  assert.match(f.toasts.at(-1),/сохранён на сервере: 1/);
  assert.equal(f.renders.at(-1),2);assert.equal(f.nodes.fiDict.disabled,false);
  assert.equal(f.nodes.fiDict.value,'');assert.equal(f.nodes.btnLoadDict.disabled,false);
});

test('a late initial GET cannot replace the dictionary already saved by an import',async()=>{
  const f=fixture();const pending=f.context.importCuratorDictionary({name:'fresh.xlsx'});
  f.requests[1].resolve(response({ok:true,imported:1,dictionary:clone(imported)}));await pending;
  f.requests[0].resolve(response(clone(initial)));await flush();
  assert.deepEqual(f.D.dict,imported.dict);assert.deepEqual(f.renders,[2]);
});

test('an obsolete GET failure does not report a failure after a successful import',async()=>{
  const f=fixture();const pending=f.context.importCuratorDictionary({name:'fresh.xlsx'});
  f.requests[1].resolve(response({ok:true,imported:1,dictionary:clone(imported)}));await pending;
  f.requests[0].reject(new Error('stale initial request failed'));await flush();
  assert.equal(f.toasts.length,1);assert.match(f.toasts[0],/сохранён на сервере/);
});

test('duplicate imports are ignored until the first one finishes',async()=>{
  const f=await initialized();const first=f.context.importCuratorDictionary({name:'first.xlsx'});
  assert.equal(await f.context.importCuratorDictionary({name:'second.xlsx'}),false);
  assert.equal(f.requests.length,2);
  f.requests[1].resolve(response({ok:true,imported:1,dictionary:clone(imported)}));await first;
  assert.equal(f.nodes.btnLoadDict.disabled,false);
});

test('a rejected import keeps the previous dictionary and resets controls so the same file can be retried',async()=>{
  const f=await initialized();f.document.getElementById('fiDict').value='C:\\fakepath\\bad.xlsx';
  const first=f.context.importCuratorDictionary({name:'bad.xlsx'});
  f.requests[1].resolve(response({ok:false,error:'Не найдены столбцы справочника.'},400));
  assert.equal(await first,false);assert.deepEqual(f.D.dict,initial.dict);
  assert.match(f.toasts.at(-1),/Не найдены столбцы/);assert.equal(f.nodes.fiDict.value,'');
  for(const id of ['btnLoadDict','btnSaveDict','btnCategories','fiDict'])assert.equal(f.nodes[id].disabled,false);
  const retry=f.context.importCuratorDictionary({name:'bad.xlsx'});assert.equal(f.requests.length,3);
  f.requests[2].resolve(response({ok:true,imported:1,dictionary:clone(imported)}));assert.equal(await retry,true);
});

test('a transport failure cannot change the dictionary or claim success',async()=>{
  const f=await initialized();const pending=f.context.importCuratorDictionary({name:'dict.xlsx'});
  f.requests[1].reject(new Error('Connection lost'));assert.equal(await pending,false);
  assert.deepEqual(f.D.dict,initial.dict);assert.equal(f.toasts.some(text=>text.includes('сохранён на сервере:')),false);
  assert.equal(f.nodes.btnLoadDict.attributes['aria-busy'],undefined);
});

test('a malformed success response cannot partially replace the dictionary',async()=>{
  const f=await initialized();const pending=f.context.importCuratorDictionary({name:'dict.xlsx'});
  f.requests[1].resolve(response({ok:true,imported:1,dictionary:{dict:imported.dict}}));assert.equal(await pending,false);
  assert.deepEqual(f.D.dict,initial.dict);assert.match(f.toasts.at(-1),/неполный справочник/);
});

test('download refreshes the server dictionary before creating the exported workbook',async()=>{
  const f=await initialized();const pending=f.context.downloadCuratorDictionary();
  assert.equal(f.requests[1].url,'/zip_curator/api/dictionary');assert.equal(f.requests[1].options.cache,'no-store');
  assert.equal(f.downloads.length,0);assert.equal(f.nodes.btnSaveDict.disabled,true);
  f.requests[1].resolve(response(clone(imported)));assert.equal(await pending,true);
  assert.equal(f.downloads.length,1);assert.equal(f.downloads[0].name,'справочник.xlsx');
  assert.ok(f.downloads[0].rows.some(row=>row[0]==='новая позиция'&&row[3]==='да'));
  assert.equal(f.nodes.btnSaveDict.disabled,false);
});

test('failed dictionary refresh prevents exporting a stale embedded copy',async()=>{
  const f=await initialized();const pending=f.context.downloadCuratorDictionary();
  f.requests[1].resolve(response({error:'failure'},503));assert.equal(await pending,false);
  assert.equal(f.downloads.length,0);assert.deepEqual(f.D.dict,initial.dict);
  assert.equal(f.nodes.btnSaveDict.disabled,false);
});

test('download cannot race an in-progress import',async()=>{
  const f=await initialized();const pending=f.context.importCuratorDictionary({name:'dict.xlsx'});
  assert.equal(await f.context.downloadCuratorDictionary(),false);assert.equal(f.requests.length,2);
  assert.equal(f.downloads.length,0);
  f.requests[1].resolve(response({ok:true,imported:1,dictionary:clone(imported)}));await pending;
});
