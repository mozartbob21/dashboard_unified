// Shared, read-only projection of saved checks. No model or collector is started here.
(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else api.mount(root);
})(globalThis, function() {
  'use strict';
  const routes = new Map();
  for (const id of ['edds','mingkh','edo','overdue','watercontrol','utnkr','cameras','appeals','cds','ecur',
                    'municipality-report','water-dashboard','omsu-rating','tools','summarizer','telegram']) {
    routes.set('/'+id, [id]);
  }
  routes.set('/mgkh-rm', ['mgkh_rm']);
  routes.set('/water-rm', ['water_rm']);
  routes.set('/mingkh/collective', ['collective']);
  routes.set('/mingkh/water-map', ['mingkh-water-map']);
  routes.set('/mingkh/treatment', ['mingkh','treatment']);
  for (const suffix of ['', '/summary', '/system', '/solutions', '/resonance']) {
    routes.set('/mingkh/sms'+suffix, ['mingkh','sms']);
  }
  for (const path of ['/prescriptions','/camera-prescriptions']) routes.set(path, ['camera_prescriptions']);
  routes.set('/zips', ['zips']);
  for (const path of ['/zip','/zip_curator','/zip_curator.html','/zip_curator_original.html','/zip-curator']) {
    routes.set(path, ['zips','zips_scan']);
  }
  function pageScope(path) {return routes.get(path.replace(/\/+$/,'') || '/') || null;}
  function links(moduleId, operation) {
    const filter=new URLSearchParams();
    if (operation) filter.set('operation',operation);
    const chat=new URLSearchParams({history:moduleId});
    if (operation) chat.set('operation',operation);
    return {api:'/api/check-history/'+encodeURIComponent(moduleId)+'/analysis'+(filter.size?'?'+filter:''),
            chat:'/aichat?'+chat};
  }
  function projection(data) {
    if (!data || typeof data.summary !== 'string' || !data.summary.trim()) throw new Error('Incomplete analysis');
    const date=data.latest_at && new Date(data.latest_at);
    const metrics=Array.isArray(data.metrics)?data.metrics.filter(item=>item && typeof item.label==='string'
      && typeof item.value==='number' && Number.isFinite(item.value)).slice(0,3):[];
    return {summary:data.summary, status:data.status, metrics,
            period:typeof data.period_label==='string'?data.period_label:'',
            date:date && Number.isFinite(date.getTime())?date.toLocaleString('ru-RU'):''};
  }
  function createController({read, render, notice, reset, visible}) {
    let busy=false, hasResult=false;
    async function refresh() {
      if (busy || !visible()) return;
      busy=true;
      try {
        const response=await read();
        if (response.status===401 || response.status===403 || response.redirected) {
          reset(); hasResult=false; notice('Нет доступа к истории. Проверьте вход и права.'); return;
        }
        if (!response.ok) throw new Error('History unavailable');
        render(projection(await response.json())); hasResult=true; notice('');
      } catch (_) {
        notice(hasResult?'Обновить данные не удалось. Показан предыдущий результат.':
                         'Не удалось загрузить историю. Повторим запрос, когда связь восстановится.');
      } finally { busy=false; }
    }
    return {refresh};
  }
  function mount(env) {
    if(env.top!==env.self) return;
    const doc=env.document;
    if(doc.body.classList.contains('sms-embedded')) return;
    let card=doc.querySelector('[data-check-analysis]');
    const scope=card?.dataset.checkAnalysis ? [card.dataset.checkAnalysis,card.dataset.checkOperation] : pageScope(env.location.pathname);
    if(!scope) return;
    if(!card) {
      card=doc.createElement('section');
      const zip=doc.body.classList.contains('zip-curator-page') && doc.querySelector('.wrap');
      const header=doc.querySelector('header.topbar,header.unified-shell-topbar,header.tg-header') || doc.querySelector('header');
      if(zip) zip.prepend(card);
      else if(header) header.after(card);
      else (doc.querySelector('main,.container')||doc.body).prepend(card);
    }
    const [moduleId,operation]=scope;
    card.dataset.checkAnalysis=moduleId;
    card.classList.add('check-analysis-card');
    card.setAttribute('aria-label','Краткий анализ Нейроны');
    const element=(tag,classname,text) => {
      const node=doc.createElement(tag);
      node.className=classname;
      if(text) node.textContent=text;
      return node;
    };
    const heading=element('div','check-analysis-heading');
    const title=element('h2','check-analysis-title','Краткий анализ Нейроны');
    const link=element('a','check-analysis-link','Спросить в ИИ-чате →');
    const urls=links(moduleId,operation); link.href=urls.chat;
    heading.append(title,link);
    const summary=element('p','check-analysis-summary');
    const date=element('p','check-analysis-date');
    const metrics=element('dl','check-analysis-metrics');
    const notice=element('p','check-analysis-notice','Загрузка истории…');
    notice.setAttribute('role','status');
    card.replaceChildren(heading,summary,date,metrics,notice);
    const controller=createController({
      visible:()=>!doc.hidden && !card.closest('[hidden]') && (card.getBoundingClientRect().top<=env.innerHeight+100 && card.getBoundingClientRect().bottom>=-100),
      read:()=>env.fetch(urls.api,{credentials:'same-origin',cache:'no-store', signal:globalThis.AbortSignal?.timeout?.(12000)}),
      notice:text=>{notice.textContent=text; notice.hidden=!text;},
      reset:()=>{summary.textContent=''; date.textContent=''; metrics.replaceChildren(); delete card.dataset.state;},
      render:data=>{
        summary.textContent=data.summary;
        date.textContent=[data.date && ('Последняя проверка: '+data.date),data.period].filter(Boolean).join(' · ');
        card.dataset.state=['error','interrupted','failed','unavailable'].includes(data.status)?'error':'neutral';
        metrics.replaceChildren(...data.metrics.map(item=>{
          const box=element('div','check-analysis-metric');
          box.append(element('dt','',item.label),element('dd','',item.value.toLocaleString('ru-RU')+(item.unit?' '+item.unit:'')));
          return box;
        }));
      }
    });
    controller.refresh();
    const interval=env.setInterval(controller.refresh,60000);
    let observer;
    if (env.IntersectionObserver) {
      observer=new env.IntersectionObserver(entries=>{
        if (entries.some(entry=>entry.isIntersecting) && !summary.textContent) controller.refresh();
      });
      observer.observe(card);
    }
    doc.addEventListener('visibilitychange',controller.refresh);
    env.addEventListener('online',controller.refresh);
    env.addEventListener('pagehide',event=>{if(!event.persisted){env.clearInterval(interval);observer?.disconnect();}});
  }
  return {mount,pageScope,links,projection,createController};
});
