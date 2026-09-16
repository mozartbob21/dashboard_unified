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
