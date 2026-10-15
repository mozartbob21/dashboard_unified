'use strict';
const assert=require('node:assert/strict'),test=require('node:test'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const root=process.env.ZIP_CURATOR_UI_ROOT||path.join(__dirname,'..');
const code=fs.readFileSync(path.join(root,'static/zip-curator-review.js'),'utf8');
const template=fs.readFileSync(path.join(root,'templates/zip_curator_original.html'),'utf8');
const esc=value=>String(value??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const item=()=>({name:'Труба',cat:'Трубы',grp:'Стальная',qty:0,unit:'пог. м',unitUnknown:false,via:'override'});
const record=(count=1)=>({rso:'Водоканал',okrug:'Королёв',revision:'v1',items:Array.from({length:count},(_,i)=>({...item(),name:'Труба '+i}))});
const response=(body,status=200)=>({ok:status<300,status,json:async()=>body});
function node(){return {innerHTML:'',textContent:'',disabled:false,isConnected:true,attributes:{},classList:{toggle(){},add(){},remove(){}},setAttribute(name,value){this.attributes[name]=value;},querySelector(){return null;},scrollIntoView(){}};}
function fixture(count=1){
 const nodes={'#pv0':node(),'#reviewToggle0':node(),'#approve0':node(),'#approvalNotes0':node()};
 const calls=[],notices=[],states=[];let result=null;
 const PENDING=[record(count)],D={categories:['Трубы','Другое'],groupsByCat:{Трубы:['Стальная'],Другое:[]},units:['шт','пог. м'],umap:{шт:'шт','пог. м':'пог. м'},dict:{}};
 const ctx=vm.createContext({PENDING,D,esc,norm:value=>value.toLowerCase(),normUnit:value=>({canon:D.umap[value],unknown:!D.umap[value]}),fmt:String,$:selector=>nodes[selector],window:{addEventListener(){}},toast:text=>notices.push(text),applyServerState:state=>states.push(state),fetch:async(url,options)=>{const body=JSON.parse(options.body);calls.push({url,body});if(result)return typeof result==='function'?result(body):await result;const next=JSON.parse(JSON.stringify(PENDING));Object.assign(next[body.pi].items[body.ii],body);next[body.pi].revision='v'+(calls.length+1);next[body.pi].approval_issues=[];return response({ok:true,state:{pending:next,clean:{}}});}});
 vm.runInContext(code,ctx);
 return {ctx,nodes,calls,notices,states,PENDING,D,setResult:value=>result=value,run:script=>vm.runInContext(script,ctx)};
}
function editor(f,ii=0,values={}){
 const controls=Object.fromEntries(['cat','grp','qty','unit'].map(field=>[field,{value:values[field]??f.PENDING[0].items[ii][field],disabled:false,focus(){this.focused=true;}}]));
 const note=node(),row=node();row.querySelector=selector=>selector==='.review-save-note'?note:controls[selector.match(/data-edit="([^"]+)/)?.[1]];row.querySelectorAll=()=>Object.values(controls);
 f.ctx.editorRow=row;f.ctx.editorValues=Object.fromEntries(Object.entries(controls).map(([k,v])=>[k,v.value]));
 f.run(`zipReviewEditors.set(PENDING[0].items[${ii}],{record:PENDING[0],values:editorValues,row:editorRow,error:''})`);
 return {row,note,controls};
}
test('complete registry accepts zero and running metres; missing OMSU, units and quantity have concrete blockers',()=>{
 const f=fixture();assert.equal(f.run('zipApprovalIssues(PENDING[0]).length'),0);
 f.PENDING[0].okrug=' ';f.PENDING[0].items[0].qty=null;f.PENDING[0].items[0].unit=null;
 assert.deepEqual(Array.from(f.run('zipApprovalIssues(PENDING[0]).map(x=>x.field)')),['okrug','unit','qty']);
 f.ctx.updateApprovalStatus(0);assert.equal(f.nodes['#approve0'].disabled,true);assert.match(f.nodes['#approvalNotes0'].innerHTML,/Укажите ОМСУ/);
});
test('no groups are required for categories without groups, but a stale group cannot approve',()=>{
 const f=fixture();Object.assign(f.PENDING[0].items[0],{cat:'Другое',grp:null});assert.equal(f.run('zipApprovalIssues(PENDING[0]).length'),0);
 f.PENDING[0].items[0].grp='Неизвестная';assert.equal(f.run('zipApprovalIssues(PENDING[0])[0].field'),'grp');
});
test('all 606 rows are reachable using absolute indices and last page is clamped',async()=>{
 const f=fixture(606);await f.ctx.changeReviewPage(0,10);let html=f.nodes['#pv0'].innerHTML;
 assert.match(html,/Позиции 501–550 из 606/);assert.match(html,/editRow\(this,0,500\)/);assert.doesNotMatch(html,/editRow\(this,0,0\)/);
 await f.ctx.changeReviewPage(0,999);html=f.nodes['#pv0'].innerHTML;assert.match(html,/Позиции 601–606 из 606/);assert.match(html,/data-item-index="605"/);
 await f.ctx.changeReviewPage(0,0,200);assert.match(f.nodes['#pv0'].innerHTML,/Позиции 1–200 из 606/);
 assert.doesNotMatch(template,/slice\(0,400\)|первые 400/);
});
test('moving pages persists unit and decimal quantity for the absolute row before moving',async()=>{
 const f=fixture(606);f.run('zipReviewView(PENDING[0]).page=10');editor(f,503,{unit:'пог. м',qty:'12,5'});
 await f.ctx.changeReviewPage(0,11);
 assert.equal(f.calls.length,1);assert.equal(f.calls[0].body.ii,503);assert.equal(f.calls[0].body.unit,'пог. м');assert.equal(f.calls[0].body.qty,12.5);assert.equal(f.calls[0].body.revision,'v1');
 assert.equal(f.PENDING[0].items[503].qty,12.5);assert.equal(f.PENDING[0].revision,'v2');assert.equal(f.run('zipReviewView(PENDING[0]).page'),11);assert.equal(f.run('zipReviewEditors.size'),0);
});
test('network failure keeps drafts and the previous page; a safe retry saves them',async()=>{
 const f=fixture(606);f.run('zipReviewView(PENDING[0]).page=10');const edit=editor(f,503,{qty:'55'});
 f.setResult(async()=>{throw new Error('Нет соединения');});await f.ctx.changeReviewPage(0,11);
 assert.equal(f.run('zipReviewView(PENDING[0]).page'),10);assert.equal(f.run('zipReviewEditors.size'),1);assert.equal(f.PENDING[0].items[503].qty,0);assert.match(edit.note.textContent,/не сохранена/);
 f.setResult(null);await f.ctx.changeReviewPage(0,11);assert.equal(f.PENDING[0].items[503].qty,55);assert.equal(f.run('zipReviewView(PENDING[0]).page'),11);
});
test('simultaneous saves serialize revisions and repeated row clicks share a single request',async()=>{
 const f=fixture(2);editor(f,0,{qty:'1'});editor(f,1,{qty:'2'});
 await Promise.all([f.ctx.saveReviewRow(0,0),f.ctx.saveReviewRow(0,0),f.ctx.saveReviewRow(0,1)]);
 assert.equal(f.calls.length,2);assert.equal(f.calls[0].body.revision,'v1');assert.equal(f.calls[1].body.revision,'v2');assert.equal(f.PENDING[0].revision,'v3');
});
test('negative or malformed quantity never saves or changes the page',async()=>{
 for(const value of ['-1','invalid','Infinity']){const f=fixture(100);const edit=editor(f,0,{qty:value});await f.ctx.changeReviewPage(0,1);assert.equal(f.calls.length,0);assert.equal(f.run('zipReviewView(PENDING[0]).page'),0);assert.match(edit.note.textContent,/не меньше нуля/);}
});
test('unsafe names and error messages remain inert HTML and valid approval is guarded',async()=>{
 const f=fixture();f.PENDING[0].rso='\"><img src=x onerror=alert(1)>';f.PENDING[0].items[0].name='<script>evil()</script>';
 await f.ctx.changeReviewPage(0,0);assert.doesNotMatch(f.nodes['#pv0'].innerHTML,/<script>evil|<img/);assert.match(f.nodes['#pv0'].innerHTML,/&lt;script&gt;/);
 f.PENDING[0].okrug='';await f.ctx.approve(0);assert.equal(f.calls.length,0);
 const markup=f.run('zipUnitSelect("пог. м")');assert.match(markup,/погонный метр/);assert.match(markup,/aria-label="Единица измерения"/);
});
test('stale revision response retains the draft and does not silently approve or overwrite',async()=>{
 const f=fixture();editor(f,0,{qty:'19'});f.setResult(response({ok:false,error:'Реестр изменён в другой вкладке.'},409));assert.equal(await f.ctx.saveReviewRow(0,0),false);assert.equal(f.run('zipReviewEditors.size'),1);assert.equal(f.PENDING[0].revision,'v1');assert.equal(f.PENDING[0].items[0].qty,0);
});
test('approved legacy rows without municipality can be repaired in the mapping editor',()=>{
 assert.match(template,/Указать ОМСУ/);assert.match(template,/openMunicipalityEditor\(null,decodeURIComponent/);assert.match(template,/function openMunicipalityEditor\(pi,organization\)/);
});
test('double approval clicks issue one request after awaiting editor flush',async()=>{
 const f=fixture();let resolve;f.setResult(new Promise(done=>{resolve=done;}));
 const first=f.ctx.approve(0),second=f.ctx.approve(0);await new Promise(done=>setImmediate(done));
 assert.equal(f.calls.length,1);assert.equal(f.calls[0].body.revision,'v1');
 resolve(response({ok:true,state:{pending:[],clean:{}}}));await Promise.all([first,second]);assert.equal(f.states.length,1);
});
