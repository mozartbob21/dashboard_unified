(() => {
  const root=document.querySelector('.tools-page'), $=id=>document.getElementById(id);
  const jobKey='neurona-tools-job:'+root.dataset.owner;
  let mergeFiles=[],busy=false,pdfFiles=[],pdfReading=false,pdfVersion=0;
  const pdfCounts=new Map();
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
  function setBusy(value){busy=value;root.querySelectorAll('button[type=submit]').forEach(b=>b.disabled=value);$('pdfRun').disabled=value||pdfReading;}
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
    if(form.dataset.kind==='pdf_pages'){
      if(pdfReading||!pdfFiles.length){setBusy(false);status('Выберите PDF и дождитесь проверки файла.',true);return;}
      data.delete('file');pdfFiles.forEach(file=>data.append('file',file,file.name));
    }
    if(form.dataset.kind==='html_pptx'&&!form.elements.file.files.length)data.delete('file');
    try{const d=await api('/tools/api/run',{method:'POST',body:data});sessionStorage.setItem(jobKey,d.job_id);poll(d.job_id);}catch(e){setBusy(false);status(e.message,true);}
  }));
  function selected(files,folder){mergeFiles=Array.from(files).filter(f=>/\.(xlsx|xls|xlsb)$/i.test(f.name)&&!f.name.startsWith('~$')&&f.name.toUpperCase()!=='ОБЪЕДИНЕННЫЙ_РЕЗУЛЬТАТ.XLSX'&&(!folder||f.webkitRelativePath.split('/').length===2));$('mergeSelection').textContent='Выбрано файлов: '+mergeFiles.length;}
  $('mergeFiles').addEventListener('change',e=>selected(e.target.files,false));$('mergeFolder').addEventListener('change',e=>selected(e.target.files,true));
  const pdfActions={
    merge:['Объединить PDF →','Файлы объединятся в порядке списка. Меняйте его кнопками со стрелками.','Страницы'],
    extract:['Извлечь страницы →','Укажите нужные страницы, например 1, 3-5. Они попадут в один новый PDF.','Страницы для извлечения'],
    delete:['Удалить страницы →','Укажите страницы, которые нужно убрать, например 2, 5-7. Исходный файл сохранится у вас.','Страницы для удаления'],
    reorder:['Изменить порядок →','Перечислите все страницы ровно по одному разу: например 3, 1-2. Для обратного порядка можно указать 5-1.','Новый порядок страниц'],
    rotate:['Повернуть страницы →','Укажите страницы или оставьте поле пустым, чтобы повернуть весь документ.','Страницы — пусто означает все'],
    split:['Разделить PDF →','Каждая выбранная страница станет отдельным PDF. Результат — ZIP. Пустое поле означает все страницы.','Страницы — пусто означает все']
  };
  function renderPdfFiles(){
    $('pdfFileList').replaceChildren();
    pdfFiles.forEach((file,index)=>{
      const row=document.createElement('li'),name=document.createElement('span'),meta=document.createElement('small'),controls=document.createElement('div');
      name.className='tool-file-name';name.textContent=(index+1)+'. '+file.name;
      const size=file.size<1024*1024?Math.max(1,Math.ceil(file.size/1024))+' КБ':(file.size/1024/1024).toLocaleString('ru-RU',{maximumFractionDigits:1})+' МБ';
      meta.textContent=size+(pdfCounts.has(file)?' · страниц: '+pdfCounts.get(file):'');name.append(meta);
      controls.className='tool-file-controls';
      const button=(text,label,action,disabled=false)=>{const b=document.createElement('button');b.type='button';b.textContent=text;b.setAttribute('aria-label',label);b.disabled=disabled;b.addEventListener('click',action);controls.append(b);return b;};
      if($('pdfAction').value==='merge'){
        button('↑','Поднять '+file.name,()=>{[pdfFiles[index-1],pdfFiles[index]]=[pdfFiles[index],pdfFiles[index-1]];renderPdfFiles();},index===0);
        button('↓','Опустить '+file.name,()=>{[pdfFiles[index+1],pdfFiles[index]]=[pdfFiles[index],pdfFiles[index+1]];renderPdfFiles();},index===pdfFiles.length-1);
      }
      button('×','Убрать '+file.name,()=>{pdfFiles.splice(index,1);renderPdfFiles();if(!pdfFiles.length)$('pdfFiles').value='';}).dataset.remove='true';
      row.append(name,controls);$('pdfFileList').append(row);
    });
    if(pdfFiles.length&&pdfFiles.every(f=>pdfCounts.has(f)))$('pdfInspection').textContent='Файлов: '+pdfFiles.length+' · всего страниц: '+pdfFiles.reduce((n,f)=>n+pdfCounts.get(f),0)+'. Обработка на сервере организации.';
    if(!pdfFiles.length)$('pdfInspection').textContent='Выберите PDF-файлы. Всего до 500 страниц.';
  }
  function pdfAction(){
    const action=$('pdfAction').value,config=pdfActions[action];
    $('pdfRun').textContent=config[0];$('pdfActionHint').textContent=config[1];$('pdfPagesLabel').textContent=config[2];
    $('pdfPagesField').hidden=action==='merge';$('pdfAngleField').hidden=action!=='rotate';
    $('pdfPageNumbers').required=['extract','delete','reorder'].includes(action);
    $('pdfFiles').multiple=action==='merge';$('pdfFileLabel').textContent=action==='merge'?'PDF-файлы':'PDF-документ';
    renderPdfFiles();
  }
  $('pdfAction').addEventListener('change',pdfAction);
  $('pdfFiles').addEventListener('change',async event=>{
    const version=++pdfVersion;pdfFiles=Array.from(event.target.files);const selectedFiles=[...pdfFiles];pdfCounts.clear();renderPdfFiles();
    if(!pdfFiles.length){pdfReading=false;setBusy(busy);return;}
    pdfReading=true;setBusy(busy);$('pdfInspection').classList.remove('error');$('pdfInspection').textContent='Проверяю число страниц…';
    try{
      if(pdfFiles.length>20||pdfFiles.some(f=>f.size>20*1024*1024)||pdfFiles.reduce((sum,f)=>sum+f.size,0)>60*1024*1024)throw new Error('До 20 PDF: каждый до 20 МБ, суммарно до 60 МБ.');
      const fd=new FormData();selectedFiles.forEach(f=>fd.append('file',f,f.name));const data=await api('/tools/api/pdf-info',{method:'POST',body:fd});
      if(version!==pdfVersion)return;
      data.files.forEach((item,index)=>pdfCounts.set(selectedFiles[index],item.pages));renderPdfFiles();
    }catch(e){if(version===pdfVersion){$('pdfInspection').textContent=e.message;$('pdfInspection').classList.add('error');}}
    finally{if(version===pdfVersion){pdfReading=false;setBusy(busy);}}
  });
  const htmlForm=root.querySelector('[data-kind=html_pptx]');htmlForm.elements.mode.addEventListener('change',()=>{$('htmlAiNote').hidden=htmlForm.elements.mode.value!=='ai';$('toolTemplate').disabled=htmlForm.elements.mode.value==='shots';});
  $('previewHtml').addEventListener('click',async()=>{
    try{const html=htmlForm.elements.file.files[0]?await htmlForm.elements.file.files[0].text():htmlForm.elements.html.value;const fd=new FormData();fd.append('html',html);const d=await api('/tools/html2pptx/preview',{method:'POST',body:fd});$('htmlPreview').replaceChildren();$('htmlPreview').hidden=false;d.slides.forEach((slide,index)=>{const box=document.createElement('div');box.className='tool-preview-slide';const title=document.createElement('strong');title.textContent=(index+1)+'. '+(slide.title||'Без заголовка');const text=document.createElement('p');text.textContent=[...(slide.bullets||[]),...(slide.texts||[])].join(' · ');box.append(title,text);$('htmlPreview').append(box);});}catch(e){status(e.message,true);}
  });
  function templates(list){const selected=$('toolTemplate').value;$('toolTemplate').replaceChildren(new Option('Стандартный',''));$('templateList').replaceChildren();list.forEach(name=>{$('toolTemplate').add(new Option(name,name));const row=document.createElement('div');row.className='tool-template-row';const label=document.createElement('span');label.textContent=name;const remove=document.createElement('button');remove.type='button';remove.textContent='×';remove.setAttribute('aria-label','Удалить общий шаблон '+name);remove.onclick=async()=>{if(!confirm('Удалить общий шаблон «'+name+'» для всех аккаунтов?'))return;try{const fd=new FormData();fd.append('name',name);const d=await api('/tools/template/delete',{method:'POST',body:fd});branding(d);}catch(e){status(e.message,true);}};row.append(label,remove);$('templateList').append(row);});if(list.includes(selected))$('toolTemplate').value=selected;}
  function branding(data){templates(data.templates||[]);$('emblemPreview').hidden=!data.emblem_url;if(data.emblem_url&&data.emblem_url.startsWith('/tools/emblem'))$('emblemImage').src=data.emblem_url;else $('emblemImage').removeAttribute('src');$('brandingWarnings').textContent=(data.branding_warnings||[]).join('\n');$('brandingWarnings').hidden=!data.branding_warnings?.length;}
  [['templateUpload','/tools/template','template_file'],['emblemUpload','/tools/emblem','emblem_file']].forEach(([id,url,key])=>$(id).addEventListener('change',async event=>{const file=event.target.files[0];if(!file)return;try{const fd=new FormData();fd.append(key,file);const d=await api(url,{method:'POST',body:fd});branding(d);status('Общее оформление сохранено и доступно всем аккаунтам.');}catch(e){status(e.message,true);}event.target.value='';}));
  api('/tools/api/options').then(d=>{branding(d);$('pdfAvailability').textContent=d.pdf_docx?'':'Обновите зависимости сервера: нужен pdf2docx.';if($('exeAvailability')&&!d.exe)$('exeAvailability').textContent='Нужны Windows Sandbox и отдельная среда Python с PyInstaller. Настройка: docs/tools-and-mingkh.md.';}).catch(e=>status(e.message,true));
  try{const pending=sessionStorage.getItem(jobKey);if(pending)poll(pending);}catch{}
})();
