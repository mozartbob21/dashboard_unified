(() => {
'use strict';
const $ = id => document.getElementById(id);
const badges = {pending:'На утверждении',approved:'Утверждён',revision:'На доработке'};
const backendName = value => ({qwen:'Госчат · ИИ',algo:'Алгоритм',gigachat:'GigaChat · ИИ'})[value] || 'Движок не указан';
const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const statusName = value => badges[value] || 'Статус не указан';
const statusClass = value => Object.hasOwn(badges,value) ? value : 'pending';
let current = null, historyReport = null, historyOffset = 0, historyRequest = 0, detailRequest = 0;
let historyTotal = 0, deleteTarget = null, deleting = false;
const deletedReports = new Set();
const trashIcon = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18M9 6V4h6v2M5 6l1 14h12l1-14M10 10v6M14 10v6"/></svg>';
const DEMO = 'Иванов: В 10:30 произошла авария на водоводе в Серпухове. Отключено холодное водоснабжение 12 домов.\nПетрова: Бригада водоканала на месте, 4 человека и 2 единицы техники.\nИванов: Плановое завершение работ в 14:00. Жители проинформированы.';
function setStatus(text, type) { $('statusLine').textContent=text; $('statusLine').className='sm-status '+(type||''); }
async function api(path, payload, method) {
    const options=payload===undefined ? {cache:'no-store'} : {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)};
    if(method)options.method=method;
    const response = await fetch(path, options);
    if(response.redirected || !(response.headers.get('content-type')||'').includes('application/json')) throw new Error('Сессия завершена. Войдите в систему заново.');
    const data = await response.json();
    if(!response.ok || data.ok===false) {
        const error=new Error(data.message||data.detail||data.error||'Не удалось выполнить запрос.');
        error.status=response.status;
        if(response.status===409)error.report=data.report;
        throw error;
    }
    return data;
}
async function copyText(text,button) {
    const label=button.textContent;
    try {
        if(navigator.clipboard && window.isSecureContext) await navigator.clipboard.writeText(text);
        else {
            const box=document.createElement('textarea'); box.value=text; box.style.position='fixed'; box.style.opacity='0';
            (button.closest('dialog')||document.body).appendChild(box); box.select();
            try { if(!document.execCommand('copy')) throw new Error('copy'); } finally { box.remove(); }
        }
        button.textContent='Скопировано';
    } catch(error) { button.textContent='Выделите и скопируйте текст'; }
    setTimeout(()=>{button.textContent=label;},1800);
}
$('demoBtn').addEventListener('click',()=>{$('srcText').value=DEMO;});
$('summaryBtn').addEventListener('click',async()=>{
    const text=$('srcText').value.trim();
    if(text.length<50) return setStatus('Нужно минимум 50 символов текста.','err');
    $('summaryBtn').disabled=true; setStatus('Анализирую текст…');
    try {
        const data=await api('/summarizer/api/summary',{text,backend:$('engineSelect').value});
        current=data.report; renderReport(current); await loadHistory();
        setStatus('Отчёт готов и сохранён в истории.','ok');
    } catch(error) { setStatus(error.message,'err'); }
    finally { $('summaryBtn').disabled=false; }
});
function renderReport(rep) {
    $('reportCard').hidden=false;
    const res=rep.result||{}, stats=res.stats||{}, approvals=Object.keys(rep.approvals||{}), reportText=res.report_text||'';
    $('reportBox').innerHTML=`
      <div class="sm-report">
        <div class="sm-stats"><span class="sm-stat">${esc(backendName(res.backend))}</span><span class="sm-stat">${esc(stats.time||'—')}</span><span class="sm-stat">${esc(stats.place||'—')}</span><span class="sm-badge ${statusClass(rep.status)}">${esc(statusName(rep.status))}</span></div>
        ${res.warning?`<p class="sm-status">${esc(res.warning)}</p>`:''}
        <details class="sm-source-details"><summary>Что суммировали · ${esc(rep.created_at)} · ${esc(rep.author)}</summary><pre class="sm-history-text">${esc(rep.source||'Исходный текст не сохранён')}</pre></details>
        <pre class="sm-history-text" id="docBox">${esc(reportText)}</pre>
        <div class="sm-dialog-foot"><button class="sm-btn ghost" id="copyBtn" type="button">Скопировать результат</button><span class="sm-sub">Сжатие ${esc(stats.compression||0)}% · источников: ${esc(stats.authors||0)}</span></div>
        ${(res.key_points||[]).length?`<div class="sm-h">Не распознано</div><ul>${res.key_points.map(p=>`<li>${esc(p)}</li>`).join('')}</ul>`:''}
      </div>
      <div class="sm-approve">
        ${rep.status==='pending'?'<button class="sm-btn green" id="approveBtn" type="button">Утвердить</button>':''}
        ${rep.status!=='approved'?'<button class="sm-btn red" id="rejectBtn" type="button">На доработку</button>':''}
        ${rep.status==='revision'?'<button class="sm-btn primary" id="regenBtn" type="button">Пересобрать с учётом комментария</button>':''}
        <div class="sm-badges">${approvals.map(u=>`<span class="sm-badge approved">${esc(u)}</span>`).join('')}</div>
        <span class="sm-sub">Для утверждения нужны 2 сотрудника</span>
      </div>
      ${rep.status!=='approved'?`<textarea class="sm-comment" id="commentBox" aria-label="Комментарий для доработки" placeholder="Что нужно исправить…">${esc(rep.revision_comment||'')}</textarea>`:''}`;
    $('copyBtn').addEventListener('click',event=>copyText(reportText,event.currentTarget));
    [['approveBtn','approve'],['rejectBtn','reject'],['regenBtn','regenerate']].forEach(([id,action])=>{
        const button=$(id); if(!button)return;
        button.addEventListener('click',async()=>{
            const comment=($('commentBox')||{}).value||'';
            if(action==='reject'&&!comment.trim())return setStatus('Укажите комментарий для доработки.','err');
            $('reportBox').querySelectorAll('button').forEach(b=>b.disabled=true);
            try {
                const data=await api('/summarizer/api/'+action,{id:rep.id,comment,backend:$('engineSelect').value});
                if(deletedReports.has(rep.id))return;
                current=data.report; renderReport(current); await loadHistory(); setStatus('Изменения сохранены.','ok');
            } catch(error) {
                if(deletedReports.has(rep.id))return;
                if(error.report){current=error.report;renderReport(current);await loadHistory();}
                setStatus(error.message,'err');
            }
            finally { $('reportBox').querySelectorAll('button').forEach(b=>b.disabled=false); }
        });
    });
}
async function loadHistory(append=false) {
    const sequence=++historyRequest, offset=append?historyOffset:0;
    $('refreshHistoryBtn').disabled=true; $('moreHistoryBtn').disabled=true;
    $('clearHistoryBtn').disabled=true;
    $('historyState').textContent='Загружаю историю…';
    try {
        const data=await api('/summarizer/api/reports?limit=20&offset='+offset);
        if(sequence!==historyRequest)return;
        if(!append)$('historyBox').replaceChildren();
        data.items.forEach(item=>{
            const row=document.createElement('div');row.className='sm-hrow';
            const button=document.createElement('button');button.type='button';button.className='sm-hitem';
            button.innerHTML=`<span class="sm-history-meta"><span>${esc(item.created_at)} · ${esc(item.author)} · ${esc(backendName(item.backend))}</span><span class="sm-badge ${statusClass(item.status)}">${esc(statusName(item.status))}</span></span><span class="sm-history-preview"><span><b>Исходный текст</b>${esc(item.source_preview||'Не сохранён')}</span><span><b>Полученный результат</b>${esc(item.result_preview||'Нет текста результата')}</span></span><span class="sm-history-open">Открыть полностью${item.version_count>1?' · версий: '+esc(item.version_count):''} →</span>`;
            button.addEventListener('click',()=>openHistory(item.id));
            const remove=document.createElement('button');remove.type='button';remove.className='sm-btn sm-delete sm-delete-icon';
            remove.innerHTML=trashIcon;remove.title='Удалить запись';
            remove.setAttribute('aria-label','Удалить запись от '+item.created_at+' · '+item.author);
            remove.addEventListener('click',()=>askDelete(item));
            row.append(button,remove);$('historyBox').appendChild(row);
        });
        historyTotal=data.total;
        historyOffset=offset+data.items.length;
        $('moreHistoryBtn').hidden=!data.has_more;
        $('historyState').textContent=data.total?'Показано '+historyOffset+' из '+data.total:'Здесь появятся исходные тексты и результаты после первого суммирования.';
    } catch(error) {if(sequence===historyRequest)$('historyState').textContent=error.message;}
    finally { if(sequence===historyRequest){$('refreshHistoryBtn').disabled=false;$('moreHistoryBtn').disabled=false;$('clearHistoryBtn').disabled=!historyTotal||deleting;} }
}
function askDelete(item=null) {
    if(deleting)return;
    deleteTarget=item;
    $('deleteTitle').textContent=item?'Удалить запись?':'Очистить всю историю?';
    $('deleteDescription').textContent=item
        ? 'Запись от '+item.created_at+' будет удалена из общей истории для всех сотрудников вместе с исходным текстом, результатом и версиями. Восстановить её не получится.'
        : 'Будут удалены все записи общей истории, включая записи на следующих страницах, исходные тексты, результаты и версии. Это действие нельзя отменить.';
    $('confirmDeleteBtn').textContent=item?'Удалить запись':'Очистить всё';
    $('deleteState').textContent='';$('smDeleteDialog').showModal();$('cancelDeleteBtn').focus();
}
function forgetDeletedReport(id) {
    if(id)deletedReports.add(id);
    if(current && (!id||current.id===id)){
        deletedReports.add(current.id);current=null;$('reportBox').replaceChildren();$('reportCard').hidden=true;
        setStatus('');
    }
    if(historyReport && (!id||historyReport.id===id)){
        historyReport=null;detailRequest++;$('smHistoryDialog').close();
    }
}
$('clearHistoryBtn').addEventListener('click',()=>askDelete());
$('cancelDeleteBtn').addEventListener('click',()=>{$('smDeleteDialog').close();});
$('smDeleteDialog').addEventListener('cancel',event=>{if(deleting)event.preventDefault();});
$('confirmDeleteBtn').addEventListener('click',async()=>{
    if(deleting)return;
    const target=deleteTarget;
    deleting=true;$('confirmDeleteBtn').disabled=true;$('cancelDeleteBtn').disabled=true;$('clearHistoryBtn').disabled=true;
    $('deleteState').textContent='Удаляю…';
    historyRequest++;
    try {
        try { await api('/summarizer/api/reports'+(target?'/'+encodeURIComponent(target.id):''),undefined,'DELETE'); }
        catch(error){if(!(target && error.status===404))throw error;}
        forgetDeletedReport(target?target.id:null);
        $('smDeleteDialog').close();
        await loadHistory();
        $('historyState').textContent=(target?'Запись удалена. ':'История очищена. ')+$('historyState').textContent;
        $('refreshHistoryBtn').focus();
    } catch(error){$('deleteState').textContent=error.message;}
    finally{
        deleting=false;$('confirmDeleteBtn').disabled=false;$('cancelDeleteBtn').disabled=false;
        $('refreshHistoryBtn').disabled=false;$('moreHistoryBtn').disabled=false;$('clearHistoryBtn').disabled=!historyTotal;
    }
});
async function openHistory(id) {
    const sequence=++detailRequest;
    historyReport=null;$('historyDetail').hidden=true;$('historyMeta').textContent='';$('historyDetailState').textContent='Загружаю отчёт…';
    $('smHistoryDialog').showModal();$('closeHistoryBtn').focus();
    try {
        const data=await api('/summarizer/api/reports/'+encodeURIComponent(id));
        if(sequence!==detailRequest || !$('smHistoryDialog').open)return;
        historyReport=data.report;
        $('historyMeta').textContent=historyReport.created_at+' · '+historyReport.author+' · '+statusName(historyReport.status);
        $('historyVersion').replaceChildren();
        (historyReport.versions||[]).forEach((version,index)=>{
            const option=document.createElement('option');option.value=String(index);option.textContent=(index+1)+'. '+version.created_at+' · '+version.author;$('historyVersion').appendChild(option);
        });
        $('historyVersion').value=String(historyReport.versions.length-1);
        renderVersion();$('historyDetail').hidden=false;$('historyDetailState').textContent='';
    } catch(error) { if(sequence===detailRequest)$('historyDetailState').textContent=error.message; }
}
function renderVersion() {
    if(!historyReport)return;
    const version=historyReport.versions[Number($('historyVersion').value)], result=version.result||{};
    $('historySource').textContent=version.source||'Исходный текст не сохранён.';
    $('historyResult').textContent=result.report_text||'Текст результата не сохранён.';
    $('historyBackend').textContent=backendName(result.backend);
    const warnings=[historyReport.source_may_be_truncated?'В старой версии сохранялись только первые 20 000 символов. Полный исходник этой записи может отсутствовать.':'',result.warning||''].filter(Boolean);
    $('historyWarning').textContent=warnings.join(' ');$('historyWarning').hidden=!warnings.length;
}
$('refreshHistoryBtn').addEventListener('click',()=>loadHistory());
$('moreHistoryBtn').addEventListener('click',()=>loadHistory(true));
$('closeHistoryBtn').addEventListener('click',()=>{$('smHistoryDialog').close();});
$('smHistoryDialog').addEventListener('close',()=>{detailRequest++;});
$('historyVersion').addEventListener('change',renderVersion);
$('copySourceBtn').addEventListener('click',event=>copyText($('historySource').textContent,event.currentTarget));
$('copyResultBtn').addEventListener('click',event=>copyText($('historyResult').textContent,event.currentTarget));
$('historyReviewBtn').addEventListener('click',()=>{if(!historyReport)return;current=historyReport;renderReport(current);$('smHistoryDialog').close();$('reportCard').scrollIntoView({behavior:'smooth',block:'start'});});
loadHistory();
})();
