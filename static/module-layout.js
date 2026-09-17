(() => {
  const buttons=[...document.querySelectorAll('[data-module-view]')];
  const summaries={
    edo:'Проверяет полноту и корректность данных и готовит персональные уведомления по найденным проблемам.',
    overdue:'Выявляет просроченные задачи по территориям и организациям и помогает оценить их риск.',
    watercontrol:'Проверяет обязательные поля в карточках контроля воды и формирует сводку по нарушениям.',
    utnkr:'Анализирует таблицу УТНКР и распределяет проблемные объекты по уровням риска.',
    cameras:'Сверяет адресную таблицу с системой камер, проверяет видеопотоки и готовит предписания.',
    appeals:'Готовит мягкие и корректные ответы на обращения с учётом эмоции и критичности.',
    summarizer:'Превращает переписку в краткий отчёт ЖКХ-Центра с последующим согласованием.',
    cds:'Выгружает обращения из диспетчерской 1С в структурированный список для анализа и отчётности.',
    mgkh_rm:'Проверяет задачи по качеству воды в Redmine и помогает обработать нарушения сроков.',
    zips:'Показывает остатки ЗиП по РСО с фильтрацией, сравнением периодов и динамикой.',
    telegram:'Позволяет читать выбранные чаты и отправлять сообщения из личного Telegram-аккаунта.',
    ecur:'Выгружает жалобы из ЕЦУР и помогает контролировать категории, кураторов и сроки.',
    edds:'Показывает задержки докладов по объектам водоснабжения и водоотведения, динамику и свод жалоб Добродела.',
    'water-dashboard':'Объединяет восемь операционных дашбордов DataLens в единый срез по 56 муниципалитетам.',
    'municipality-report':'Собирает единый управленческий отчёт по выбранному муниципалитету из данных всех модулей.'
  };

  document.querySelectorAll('.module-card[data-module]').forEach(card=>{
    const text=summaries[card.dataset.module];
    if(!text || card.querySelector('.module-list-summary')) return;
    const summary=document.createElement('p');
    summary.className='module-list-summary';
    summary.textContent=text;
    const actions=card.querySelector('.action-panel');
    if(actions) card.insertBefore(summary,actions);
    else card.appendChild(summary);
  });

  const grid=document.querySelector('.modules-grid');
  const regularCards=[...document.querySelectorAll('.modules-grid > .module-card[data-module]')];
  const favoritesStatus=document.getElementById('favoritesStatus');
  let preferences={can_favorite:false, eligible:[], favorites:[]};
  try { preferences=JSON.parse(document.getElementById('homePreferences')?.textContent || '{}'); } catch (_) {}
  let saving=false;
  const star='<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m12 3 2.8 5.7 6.3.9-4.55 4.45 1.08 6.28L12 17.36l-5.63 2.97 1.08-6.28L2.9 9.6l6.3-.9Z"/></svg>';

  function applyFavorites(){
    const favorites=preferences.favorites || [];
    regularCards.forEach(card=>{
      let button=card.querySelector('.favorite-toggle');
      const eligible=preferences.can_favorite && preferences.eligible?.includes(card.dataset.module);
      if(!eligible){button?.remove();return;}
      if(!button){
        button=document.createElement('button');
        button.type='button'; button.className='favorite-toggle'; button.innerHTML=star;
        button.addEventListener('click',()=>saveFavorite(card.dataset.module));
        card.querySelector('.module-head')?.appendChild(button);
      }
      const selected=favorites.includes(card.dataset.module);
      const label=(selected?'Убрать из избранного: ':'В избранное: ')+card.querySelector('.module-title').textContent.trim();
      button.setAttribute('aria-pressed',String(selected));
      button.setAttribute('aria-label',label); button.title=label;
    });
    if(grid){
      const byId=new Map(regularCards.map(card=>[card.dataset.module,card]));
      [...favorites.map(id=>byId.get(id)).filter(Boolean), ...regularCards.filter(card=>!favorites.includes(card.dataset.module))]
        .forEach(card=>grid.appendChild(card));
    }
  }

  async function saveFavorite(id){
    if(saving)return;
    const current=preferences.favorites || [];
    const favorites=current.includes(id)?current.filter(value=>value!==id):[id,...current];
    saving=true;
    document.querySelectorAll('.favorite-toggle').forEach(button=>button.disabled=true);
    try {
      const response=await fetch('/api/me/home-favorites',{
        method:'PUT', credentials:'same-origin', cache:'no-store',
        headers:{'Content-Type':'application/json'},body:JSON.stringify({favorites})
      });
      if(response.redirected || response.status===401)throw new Error('Войдите в систему заново. Избранное не изменено.');
      const data=await response.json();
      if(!response.ok)throw new Error(typeof data.detail==='string'?data.detail:'Не удалось сохранить избранное.');
      preferences=data; applyFavorites();
      if(favoritesStatus)favoritesStatus.textContent=current.includes(id)?'Модуль убран из избранного.':'Модуль добавлен в избранное и перемещён наверх.';
    } catch(error){
      if(favoritesStatus)favoritesStatus.textContent=error.message || 'Не удалось сохранить избранное. Повторите попытку.';
    } finally {
      saving=false;
      document.querySelectorAll('.favorite-toggle').forEach(button=>button.disabled=false);
      regularCards.find(card=>card.dataset.module===id)?.querySelector('.favorite-toggle')?.focus({preventScroll:true});
    }
  }
  // The initial order is supplied by the server and belongs to the signed-in account.
  applyFavorites();
  document.querySelectorAll('.module-card .tag, .module-card .pill').forEach(tag=>tag.title=tag.textContent.trim());
  document.querySelectorAll('.module-card .action-panel .secondary-button').forEach(link=>link.setAttribute('aria-label','Открыть источник'));

  function setView(view){
    view=view==='list'?'list':'cards';
    document.documentElement.dataset.moduleView=view;
    buttons.forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.moduleView===view)));
    try{localStorage.setItem('neurona-module-view',view);}catch(_){}
  }
  let initial='cards';try{initial=localStorage.getItem('neurona-module-view');}catch(_){}
  setView(initial);
  buttons.forEach(b=>b.addEventListener('click',()=>setView(b.dataset.moduleView)));

  const search=document.getElementById('moduleSearch');
  const clear=document.getElementById('moduleSearchClear');
  const status=document.getElementById('moduleSearchStatus');
  const empty=document.getElementById('moduleSearchEmpty');
  const cards=[...document.querySelectorAll('.module-card[data-module]')];
  const normalize=value=>String(value||'').toLocaleLowerCase('ru-RU').replaceAll('ё','е').replace(/\s+/g,' ').trim();

  function filterModules(){
    const query=normalize(search?.value);
    let visible=0;
    cards.forEach(card=>{
      const matches=!query||normalize(card.textContent).includes(query);
      card.hidden=!matches;
      if(matches) visible+=1;
    });
    document.querySelectorAll('.wide-pair').forEach(group=>{
      group.hidden=[...group.querySelectorAll('.module-card[data-module]')].every(card=>card.hidden);
    });
    if(clear) clear.hidden=!query;
    if(empty) empty.hidden=visible!==0;
    if(status) status.textContent=query ? `Найдено: ${visible}` : '';
  }

  if(search){
    search.addEventListener('input',filterModules);
    search.addEventListener('keydown',event=>{
      if(event.key==='Escape'&&search.value){search.value='';filterModules();search.focus();}
    });
  }
  if(clear) clear.addEventListener('click',()=>{search.value='';filterModules();search.focus();});
})();
