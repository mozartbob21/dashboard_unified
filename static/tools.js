(() => {
  const root=document.querySelector('.tools-page'), $=id=>document.getElementById(id);
  const jobKey='neurona-tools-job:'+root.dataset.owner;
  let mergeFiles=[],busy=false;
  root.querySelectorAll('[data-panel]').forEach(button=>button.addEventListener('click',()=>{
    root.querySelectorAll('[data-panel]').forEach(b=>b.setAttribute('aria-pressed',String(b===button)));
    root.querySelectorAll('[data-tool-panel]').forEach(p=>p.hidden=p.dataset.toolPanel!==button.dataset.panel);
  }));
  root.querySelectorAll('[data-conversion]').forEach(button=>button.addEventListener('click',()=>{
    root.querySelectorAll('[data-conversion]').forEach(b=>b.setAttribute('aria-pressed',String(b===button)));
    root.querySelectorAll('[data-converter]').forEach(p=>p.hidden=p.dataset.converter!==button.dataset.conversion);
  }));
  function status(message,error=false){$('toolResult').hidden=false;$('toolResult').classList.toggle('error',error);$('toolStatus').textContent=message;}
  async function api(url,options){const r=await fetch(url,options);let d;try{d=await r.json();}catch{throw new Error('Сессия завершилась или сервер недоступен. Обновите страницу.');}if(!r.ok||d.ok===false)throw new Error(d.message||d.detail||'Не удалось выполнить действие.');return d;}
  function setBusy(value){busy=value;root.querySelectorAll('button[type=submit]').forEach(b=>b.disabled=value);}
  function clearResult(){$('toolDownload').hidden=true;$('macroResult').hidden=true;}
  async function poll(id){
    setBusy(true);
    try{
      const d=await api('/tools/api/jobs/'+encodeURIComponent(id));
      if(d.state==='queued'||d.state==='running'){status(d.state==='queued'?'Задание в очереди…':'Обрабатываю… Можно переключить инструмент, результат появится здесь.');setTimeout(()=>poll(id),1500);return;}
      sessionStorage.removeItem(jobKey);setBusy(false);
      if(d.state==='error'){status(d.message,true);return;}
      status(d.message||'Готово.');
      if(d.url&&d.url.startsWith('/tools/download/')){$('toolDownload').href=d.url;$('toolDownload').hidden=false;}
      if(d.code){$('macroCode').textContent=d.code;$('macroInstructions').textContent=d.instructions||'';$('macroAssumptions').textContent=d.assumptions||'';$('macroResult').hidden=false;}
    }catch(e){sessionStorage.removeItem(jobKey);setBusy(false);status(e.message,true);}
  }
  root.querySelectorAll('form[data-kind]').forEach(form=>form.addEventListener('submit',async event=>{
    event.preventDefault();if(busy)return;clearResult();setBusy(true);status('Подготавливаю задание…');
    const data=new FormData(form);data.set('tool',form.dataset.kind);
    if(form.dataset.kind==='merge'){data.delete('file');mergeFiles.forEach(file=>data.append('file',file,file.name));}
    if(form.dataset.kind==='html_pptx'&&!form.elements.file.files.length)data.delete('file');
    try{const d=await api('/tools/api/run',{method:'POST',body:data});sessionStorage.setItem(jobKey,d.job_id);poll(d.job_id);}catch(e){setBusy(false);status(e.message,true);}
  }));
  function selected(files,folder){mergeFiles=Array.from(files).filter(f=>/\.(xlsx|xls|xlsb)$/i.test(f.name)&&!f.name.startsWith('~$')&&f.name.toUpperCase()!=='ОБЪЕДИНЕННЫЙ_РЕЗУЛЬТАТ.XLSX'&&(!folder||f.webkitRelativePath.split('/').length===2));$('mergeSelection').textContent='Выбрано файлов: '+mergeFiles.length;}
  $('mergeFiles').addEventListener('change',e=>selected(e.target.files,false));$('mergeFolder').addEventListener('change',e=>selected(e.target.files,true));
  const htmlForm=root.querySelector('[data-kind=html_pptx]');htmlForm.elements.mode.addEventListener('change',()=>{$('htmlAiNote').hidden=htmlForm.elements.mode.value!=='ai';});
  $('previewHtml').addEventListener('click',async()=>{
    try{const html=htmlForm.elements.file.files[0]?await htmlForm.elements.file.files[0].text():htmlForm.elements.html.value;const fd=new FormData();fd.append('html',html);const d=await api('/tools/html2pptx/preview',{method:'POST',body:fd});$('htmlPreview').replaceChildren();$('htmlPreview').hidden=false;d.slides.forEach((slide,index)=>{const box=document.createElement('div');box.className='tool-preview-slide';const title=document.createElement('strong');title.textContent=(index+1)+'. '+(slide.title||'Без заголовка');const text=document.createElement('p');text.textContent=[...(slide.bullets||[]),...(slide.texts||[])].join(' · ');box.append(title,text);$('htmlPreview').append(box);});}catch(e){status(e.message,true);}
  });
  function templates(list){const selected=$('toolTemplate').value;$('toolTemplate').replaceChildren(new Option('Стандартный',''));$('templateList').replaceChildren();list.forEach(name=>{$('toolTemplate').add(new Option(name,name));const row=document.createElement('div');row.className='tool-template-row';const label=document.createElement('span');label.textContent=name;const remove=document.createElement('button');remove.type='button';remove.textContent='×';remove.setAttribute('aria-label','Удалить шаблон '+name);remove.onclick=async()=>{if(!confirm('Удалить шаблон «'+name+'»?'))return;try{const fd=new FormData();fd.append('name',name);const d=await api('/tools/template/delete',{method:'POST',body:fd});templates(d.templates);}catch(e){status(e.message,true);}};row.append(label,remove);$('templateList').append(row);});if(list.includes(selected))$('toolTemplate').value=selected;}
  [['templateUpload','/tools/template','template_file'],['emblemUpload','/tools/emblem','emblem_file']].forEach(([id,url,key])=>$(id).addEventListener('change',async event=>{const file=event.target.files[0];if(!file)return;try{const fd=new FormData();fd.append(key,file);const d=await api(url,{method:'POST',body:fd});if(d.templates)templates(d.templates);status('Оформление сохранено.');}catch(e){status(e.message,true);}event.target.value='';}));
  api('/tools/api/options').then(d=>{templates(d.templates||[]);$('officeAvailability').textContent=d.docx_pdf?'':'Для этой конвертации администратору нужно установить LibreOffice на сервер.';$('pdfAvailability').textContent=d.pdf_docx?'':'Обновите зависимости сервера: нужен pdf2docx.';if($('exeAvailability')&&!d.exe)$('exeAvailability').textContent='Нужны Windows Sandbox и отдельная среда Python с PyInstaller. Настройка: docs/tools-and-mingkh.md.';}).catch(e=>status(e.message,true));
  try{const pending=sessionStorage.getItem(jobKey);if(pending)poll(pending);}catch{}
})();
