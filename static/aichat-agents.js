/* A workbook agent is an explicit chat job; its files never enter the ordinary composer. */
(function (root, factory) {
  'use strict';
  const api=factory();
  if(typeof module==='object' && module.exports) module.exports=api;
  else root.NeuronaAgents=api;
})(globalThis, function () {
  'use strict';
  const ID='complaints-neighbors-rating', MAX_FILE_BYTES=8*1024*1024;
  const DEFAULT_BOUNDS=['37.1','37.9','54.95','55.35'];
  const COORDINATES=['agentWest','agentEast','agentSouth','agentNorth'];
  const TITLE='Рейтинг домов по жалобам и соседям';

  function checkFile(file, name) {
    if(!file || !/\.xlsx$/i.test(file.name || '') || !file.size) throw Error('Выберите файл XLSX для поля «'+name+'».');
    if(file.size>MAX_FILE_BYTES) throw Error('Файл в поле «'+name+'» превышает 8 МБ.');
  }

  function prepare(values) {
    checkFile(values.geo,'Справочник домов');
    checkFile(values.complaints,'Выгрузка жалоб');
    const city=values.city.trim();
    if(!/^[а-яёa-z][а-яёa-z0-9 .-]{0,79}$/i.test(city)) throw Error('Укажите название округа длиной до 80 символов: буквы, цифры, пробелы, точки или дефисы. Первый символ — буква.');
    const bounds=values.bounds.map(value=>String(value).trim());
    const coordinates=bounds.map(value=>value ? Number(value.replace(',','.')) : NaN);
    const [west,east,south,north]=coordinates;
    if(coordinates.length!==4 || !coordinates.every(Number.isFinite) || west< -180 || east>180 || south< -90 || north>90 ||
       west>=east || south>=north) throw Error('Проверьте границы округа: западная долгота должна быть меньше восточной, южная широта — меньше северной.');
    const radius=Number(values.radius),wide=Number(values.wide);
    if(!Number.isSafeInteger(radius) || !Number.isSafeInteger(wide) || radius<1 || wide<=radius || wide>50000) {
      throw Error('Введите радиусы в целых метрах. Расширенный радиус должен быть больше основного и не превышать 50 000 м.');
    }
    return {city,bbox:coordinates.join(','),radius,wide,geo:values.geo,complaints:values.complaints};
  }

  function sameRun(a,b) {
    return Boolean(a&&b && a.city===b.city && a.bbox===b.bbox && a.radius===b.radius && a.wide===b.wide &&
      a.geo===b.geo && a.complaints===b.complaints);
  }

  function bodyFor(run,dialog,body=new FormData()) {
    body.set('agent_id',ID);
    body.set('dialog_id',dialog||'');
    body.set('text',TITLE+' · '+run.city);
    body.set('agent_city',run.city);
    body.set('agent_bbox',run.bbox);
    body.set('agent_radius',String(run.radius));
    body.set('agent_radius_wide',String(run.wide));
    body.set('geo_file',run.geo,run.geo.name);
    body.set('complaints_file',run.complaints,run.complaints.name);
    return body;
  }

  function downloadLink(doc, artifact) {
    if(!artifact || typeof artifact.id!=='string' || !/^[a-f0-9]{32}$/.test(artifact.id)) return null;
    const link=doc.createElement('a');
    link.className='ac-agent-download';
    link.href='/aichat/api/agents/artifacts/'+artifact.id;
    link.download='';
    link.textContent='Скачать рейтинг в Excel';
    return link;
  }

  function create(options) {
    const {doc,request,isBusy,dialog,key,submit,openModal,closeModal}=options;
    const el=id=>doc.getElementById(id);
    const button=el('agentsBtn'), form=el('agentRatingForm'), status=el('agentsStatus'), retry=el('agentsRetry');
    if(!button || !form)return null;
    const city=el('agentCity'), details=el('agentOptions');
    let previousCity='чехов', descriptor=null, draft=null, sequence=0, sending=false;

    function syncBusy(){
      button.disabled=isBusy() || sending;
      el('agentRatingRun').disabled=isBusy() || sending || !descriptor;
    }
    function selection(){
      return prepare({geo:el('agentGeoFile').files[0],complaints:el('agentComplaintsFile').files[0],city:city.value,
                      bounds:COORDINATES.map(id=>el(id).value),radius:el('agentRadius').value,wide:el('agentRadiusWide').value});
    }
    function reset(){
      draft=null; form.reset(); previousCity='чехов'; details.open=false;
    }
    function accept(id){
      if(!draft || draft.id!==id)return;
      let unchanged=false;
      try{unchanged=sameRun(draft.run,selection());}catch(_){}
      if(unchanged) reset();
      else draft=null;
    }
    function dismiss(id){if(draft && draft.id===id)draft=null;}
    async function load(){
      const current=++sequence;
      retry.hidden=true; descriptor=null; form.hidden=true;
      status.textContent='Загружаем каталог…'; syncBusy();
      try{
        const data=await request('/aichat/api/agents');
        if(current!==sequence)return;
        const items=data && data.agents;
        if(!Array.isArray(items))throw Error('Не удалось прочитать каталог.');
        descriptor=items.find(item=>item && item.id===ID && typeof item.title==='string')||null;
        button.hidden=!descriptor;
        if(!descriptor){
          reset();status.textContent='Для вашего аккаунта пока нет доступных агентов.';return;
        }
        el('agentRatingTitle').textContent=descriptor.title;
        el('agentRatingDescription').textContent=typeof descriptor.description==='string'?descriptor.description:'';
        form.hidden=false;status.textContent='';
      }catch(error){
        if(current!==sequence)return;
        if(error.status===401 || error.status===403){
          reset();button.hidden=true;status.textContent='Для просмотра каталога проверьте вход и права.';
        }else{
          button.hidden=false;retry.hidden=false;status.textContent='Каталог временно недоступен.';
        }
      }finally{if(current===sequence)syncBusy();}
    }
    function cityChanged(){
      const name=city.value.trim().toLocaleLowerCase('ru');
      if(name===previousCity)return;
      // Entering another district always requires its own coordinates.
      COORDINATES.forEach((id,i)=>{el(id).value=name==='чехов'?DEFAULT_BOUNDS[i]:'';});
      if(name!=='чехов')details.open=true;
      previousCity=name;
    }
    async function startRun(event){
      event.preventDefault();
      if(isBusy() || sending || !descriptor)return;
      let run;
      try{run=selection();}
      catch(error){status.textContent=error.message;return;}
      const did=dialog()||'';
      if(!draft || draft.dialog!==did || !sameRun(draft.run,run)){
        draft={id:key(),dialog:did,run};
      }
      status.textContent='';sending=true;syncBusy();
      closeModal('#agentsOvl');
      try{
        await submit(bodyFor(run,did),draft.id,TITLE+' · '+run.city);
      }finally{sending=false;syncBusy();}
    }
    function start(){
      button.addEventListener('click',()=>{
        if(isBusy())return;
        button.setAttribute('aria-expanded','true');
        openModal('#agentsOvl','#agentsClose');load();
      });
      el('agentsClose').addEventListener('click',()=>closeModal('#agentsOvl'));
      retry.addEventListener('click',load);
      city.addEventListener('input',cityChanged);
      form.addEventListener('invalid',event=>{if(details.contains(event.target))details.open=true;},true);
      form.addEventListener('submit',startRun);
      return load();
    }
    return {start,syncBusy,accept,dismiss};
  }
  return {create,prepare,sameRun,bodyFor,downloadLink};
});
