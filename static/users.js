(() => {
  'use strict';
  const form = document.getElementById('userForm');
  const message = document.getElementById('message');
  const list = document.getElementById('userList');
  const search = document.getElementById('userSearch');
  let users = [], selected = null, isParentManager = false;
  const names={username:'Логин',email:'Почта',password:'Новый пароль',modules:'Доступные блоки',is_active:'Учётная запись активна'};
  let sessionExpired=false;
  function notify(text, error=false) {
    message.textContent=text; message.classList.toggle('error',error);
    if(error) message.scrollIntoView({block:'nearest'});
  }
  function clearErrors() {
    form.querySelectorAll('[aria-invalid]').forEach(el=>el.removeAttribute('aria-invalid'));
    form.querySelectorAll('.field-error').forEach(el=>el.remove());
  }
  function fieldError(name, text) {
    const el=form.querySelector(`[name="${name}"]`);
    if(!el) return;
    el.setAttribute('aria-invalid','true');
    const hint=document.createElement('span'); hint.className='field-error'; hint.textContent=text;
    el.parentElement.append(hint);
  }
  async function api(url, options={}) {
    const response = await fetch(url, {credentials:'same-origin', ...options});
    if (response.status===401 || response.redirected) {
      sessionExpired=true;
      document.getElementById('sessionNotice').hidden=false;
      throw new Error('Сессия завершена или недействительна (401). Войдите учётной записью с правом управления пользователями. Это не ошибка заполнения полей; изменения не сохранены.');
    }
    if(response.status===403) {
      const data=await response.json().catch(()=>({}));
      throw new Error(data.detail||'Нет доступа (403). Требуется право «Управление пользователями».');
    }
    const data = await response.json().catch(()=>({}));
    if (!response.ok) {
      if(Array.isArray(data.detail)) {
        const errors=data.detail.map(item=>{
          const field=item.loc?.find(part=>Object.hasOwn(names,part));
          const rules={username:'укажите логин длиной 3–32 символа',email:'укажите корректную почту, не более 254 символов',password:'пароль — не более 72 байт UTF-8',modules:'выберите блоки из списка',is_active:'укажите состояние учётной записи'};
          const text=item.type==='missing'?'обязательное поле':rules[field]||'некорректное значение';
          if(field&&!url.includes('/integrations/')) fieldError(field,text);
          return `${names[field]||'Запрос'}: ${text}`;
        });
        throw new Error(errors.join('; '));
      }
      const text=data.detail||data.message||`Ошибка сервера (${response.status}). Изменения не сохранены. Повторите позже.`;
      const field=/парол/i.test(text)?'password':/почт/i.test(text)?'email':/логин/i.test(text)?'username':/блок/i.test(text)?'modules':null;
      if(field&&!url.includes('/integrations/')) fieldError(field,text);
      throw new Error(text);
    }
    sessionExpired=false; document.getElementById('sessionNotice').hidden=true;
    return data;
  }
  function renderList() {
    const query=search.value.trim().toLowerCase();
    const mode=document.getElementById('userArchiveFilter').value;
    const shown=users.filter(u=>(mode==='all'||(mode==='archived'?!!u.archived_at:!u.archived_at))&&[u.username,u.email||''].some(s=>s.toLowerCase().includes(query)));
    document.getElementById('userCount').textContent=`Показано ${shown.length} из ${users.length}`;
    list.replaceChildren();
    for (const user of shown) {
      const button=document.createElement('button');
      button.type='button'; button.className='user-item'; button.classList.toggle('selected',selected?.id===user.id);
      const name=document.createElement('strong'); name.textContent=user.username;
      const details=document.createElement('small');
      details.textContent=user.is_manager ? 'Родительская учётная запись' : `${user.is_active?'Активен':'Отключён'} · блоков: ${user.modules.length}${user.can_manage_users?' · Управляющий':''}`;
      if(user.archived_at) details.textContent='В архиве · вход запрещён';
      const email=document.createElement('small'); email.textContent=user.email||'Почта не указана';
      const activity=document.createElement('small');
      activity.textContent=user.activity?.last_seen_at ? `Активность: ${activityDate(user.activity.last_seen_at)}` : 'Активность ещё не учтена';
      button.append(name,details,email,activity); button.addEventListener('click',()=>openUser(user)); list.append(button);
    }
  }
  function openUser(user=null) {
    if (!form.hidden && form.dataset.dirty==='true' && !confirm('Перейти без сохранения изменений?')) return;
    selected=user; clearErrors(); form.reset(); form.hidden=false; form.dataset.dirty='false';
    const protectedUser=!!user?.can_manage_users&&!isParentManager;
    form.querySelectorAll('input,button,fieldset').forEach(el=>el.disabled=protectedUser);
    document.getElementById('editorTitle').textContent=user ? user.username : 'Новый пользователь';
    document.getElementById('editorHint').textContent=user ? 'Изменения прав действуют со следующего запроса пользователя.' : 'После создания передайте пользователю логин и пароль.';
    form.elements.username.value=user?.username||''; form.elements.username.readOnly=!!user;
    form.elements.email.value=user?.email||''; form.elements.password.value=''; form.elements.password.required=!user;
    form.elements.is_active.checked=user?.is_active??true; form.elements.is_active.disabled=protectedUser||!!user?.is_manager;
    if(user?.archived_at) form.elements.is_active.disabled=true;
    document.getElementById('moduleChoices').disabled=protectedUser||!!user?.is_manager;
    form.elements.can_manage_users.checked=!!user?.can_manage_users;
    form.elements.can_manage_users.disabled=!isParentManager||!!user?.is_manager;
    document.getElementById('delegationOption').hidden=!isParentManager;
    document.getElementById('protectedNote').hidden=!protectedUser;
    const archive=document.getElementById('archiveUser');
    archive.hidden=!user||!!user.is_manager||protectedUser;
    archive.textContent=user?.archived_at?'Восстановить из архива':'В архив';
    document.getElementById('archiveNote').hidden=!user?.archived_at;
    document.getElementById('managerNote').hidden=!user?.is_manager;
    form.querySelectorAll('[name=modules]').forEach(input=>input.checked=!!user?.modules.includes(input.value));
    renderList();
    loadActivity(user);
  }
  async function loadUsers() { const data=await api('/api/users'); users=data.users; isParentManager=data.is_parent_manager; renderList(); }
  const activityPanel=document.getElementById('userActivity');
  const activityStatus=document.getElementById('activityStatus');
  const activityContent=document.getElementById('activityContent');
  let activityVersion=0, activityData=null;
  function activityDate(value) {
    if(!value) return 'Не учтена';
    const date=new Date(value);
    return Number.isNaN(date.getTime()) ? '—' : date.toLocaleString('ru-RU');
  }
  function renderActivityRows() {
    const rows=document.getElementById('activityRows'); rows.replaceChildren();
    const modules=(activityData?.modules||[]).filter(item=>!document.getElementById('activityUsedOnly').checked||item.last_seen_at);
    for(const item of modules) {
      const row=document.createElement('tr');
      [item.label,item.visits,item.page_views,item.actions,activityDate(item.last_seen_at)].forEach((value,index)=>{
        const cell=document.createElement(index===0?'th':'td');
        if(index===0) cell.scope='row';
        cell.textContent=value; row.append(cell);
      });
      rows.append(row);
    }
    if(!modules.length) {
      const row=document.createElement('tr'), cell=document.createElement('td');
      cell.colSpan=5; cell.textContent='За время сбора активность в блоках не учтена.';
      row.append(cell); rows.append(row);
    }
  }
  async function loadActivity(user) {
    const version=++activityVersion;
    activityPanel.hidden=!user; activityContent.hidden=true; activityData=null;
    if(!user) return;
    activityStatus.textContent='Загрузка статистики…';
    try {
      const data=await api(`/api/users/${user.id}/activity`);
      if(version!==activityVersion||selected?.id!==user.id) return;
      activityData=data;
      document.getElementById('activityVisits').textContent=data.platform.visits;
      document.getElementById('activityViews').textContent=data.platform.page_views;
      document.getElementById('activityActions').textContent=data.platform.actions;
      document.getElementById('activityLastSeen').textContent=`Последняя активность: ${activityDate(data.platform.last_seen_at)}`;
      document.getElementById('activitySince').textContent=`Сбор начат: ${activityDate(data.tracking_since)}. Время показано в часовом поясе браузера. Данные за всё время сбора.`;
      activityStatus.textContent=data.platform.last_seen_at?'':'Для этой учётной записи активность ещё не учтена.';
      user.activity=data.platform; renderList(); renderActivityRows(); activityContent.hidden=false;
    } catch(error) {
      if(version===activityVersion) activityStatus.textContent=error.message||'Не удалось загрузить статистику.';
    }
  }
  document.getElementById('activityUsedOnly').addEventListener('change',renderActivityRows);
  document.getElementById('refreshActivity').addEventListener('click',()=>loadActivity(selected));
  search.addEventListener('input',renderList);
  document.getElementById('userArchiveFilter').addEventListener('change',renderList);
  document.getElementById('archiveUser').addEventListener('click',async()=>{
    if(!selected) return;
    const restore=!!selected.archived_at;
    if(!confirm(restore?'Восстановить из архива? Вход останется отключённым до отдельного включения.':'Переместить пользователя в архив? Вход будет запрещён, данные сохранятся. Несохранённые изменения формы не применятся.'))return;
    const id=selected.id;
    try{
      await api(`/api/users/${id}/${restore?'restore':'archive'}`,{method:'POST'});
      form.dataset.dirty='false'; await loadUsers(); openUser(users.find(u=>u.id===id));
      notify(restore?'Пользователь восстановлен. При необходимости включите учётную запись.':'Пользователь в архиве. Доступ к системе закрыт.');
      loadNotifications().catch(()=>{});
    }catch(e){notify(e.message,true);}
  });
  document.getElementById('newUser').addEventListener('click',()=>openUser());
  form.addEventListener('input',()=>form.dataset.dirty='true');
  form.addEventListener('invalid',event=>{
    event.target.setAttribute('aria-invalid','true');
    notify(`${names[event.target.name]||'Поле'}: ${event.target.validationMessage}`,true);
  },true);
  for (const [id,checked] of [['selectAll',true],['clearAll',false]]) {
    document.getElementById(id).addEventListener('click',()=>{form.querySelectorAll('[name=modules]').forEach(input=>input.checked=checked);form.dataset.dirty='true';});
  }
  form.addEventListener('submit',async event=>{
    event.preventDefault(); const button=document.getElementById('saveUser'); button.disabled=true;
    clearErrors();
    const payload={username:form.elements.username.value,email:form.elements.email.value,password:form.elements.password.value,
      is_active:form.elements.is_active.checked, modules:[...form.querySelectorAll('[name=modules]:checked')].map(el=>el.value)};
    if(isParentManager) payload.can_manage_users=form.elements.can_manage_users.checked;
    let committed=false;
    try {
      await api(selected?`/api/users/${selected.id}`:'/api/users',{method:selected?'PUT':'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
      committed=true;
      form.dataset.dirty='false'; form.elements.password.value='';
      await loadUsers(); openUser(users.find(u=>u.username.toLowerCase()===payload.username.trim().toLowerCase()));
      notify('Учётная запись сохранена. Права обновлены.');
      loadNotifications().catch(()=>{});
    } catch(error) {
      notify(committed?'Изменения сохранены, но обновить список не удалось. Обновите страницу после повторного входа.':error instanceof TypeError?'Не удалось связаться с сервером. Результат сохранения неизвестен; проверьте соединение и список пользователей.':error.message,true);
      if(!sessionExpired) loadNotifications().catch(()=>{});
    }
    finally { button.disabled=false; }
  });
  loadUsers().catch(error=>notify(error.message,true));
  let notificationSnapshot='', latestNotification=0, loadingNotifications=false;
  async function loadNotifications() {
    if(loadingNotifications) return;
    loadingNotifications=true;
    try {
      const data=await api('/api/users/notifications');
      const snapshot=JSON.stringify(data);
      if(snapshot===notificationSnapshot) return;
      notificationSnapshot=snapshot;
      latestNotification=data.notifications[0]?.id||0;
      document.getElementById('notificationCount').textContent=`Уведомления управляющего · непрочитанных: ${data.unread}`;
      const list=document.getElementById('notificationList'); list.replaceChildren();
      for(const item of data.notifications) {
        const li=document.createElement('li');
        const date=new Date(item.created_at.replace(' ','T')+'Z');
        li.textContent=`${date.toLocaleString('ru-RU')} — ${item.message}`;
        li.classList.toggle('unread',!item.is_read); list.append(li);
      }
      if(!data.notifications.length) list.textContent='Новых событий пока нет.';
    } finally {loadingNotifications=false;}
  }
  document.getElementById('refreshNotifications').addEventListener('click',()=>loadNotifications().catch(e=>notify(e.message,true)));
  document.getElementById('readNotifications').addEventListener('click',async()=>{
    try {await api('/api/users/notifications/read',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({through_id:latestNotification})}); await loadNotifications();}
    catch(e) {notify(e.message,true);}
  });
  loadNotifications().catch(e=>notify(e.message,true));
  setInterval(()=>{if(!document.hidden&&!sessionExpired) loadNotifications().catch(()=>{});},30000);
  const integrationForm=document.getElementById('integrationForm');
  const integrationStatus=document.getElementById('integrationStatus');
  const checkEddsLogin=document.getElementById('checkEddsLogin');
  const checkEddsHint=document.getElementById('checkEddsHint');
  let integrationVersion=0, integrationSaved=false, integrationBusy=false;
  function updateIntegrationCheck(){
    const isEdds=integrationForm.elements.service.value==='edds';
    checkEddsLogin.hidden=!isEdds; checkEddsHint.hidden=!isEdds;
    checkEddsLogin.disabled=integrationBusy||!integrationSaved;
  }
  function setIntegrationBusy(busy){
    integrationBusy=busy;
    integrationForm.querySelectorAll('input,select,button').forEach(el=>el.disabled=busy);
    updateIntegrationCheck();
  }
  async function loadIntegration(){
    const version=++integrationVersion;
    integrationSaved=false; setIntegrationBusy(true);
    integrationForm.elements.username.value=''; integrationForm.elements.password.value='';
    try{
      const data=await api('/api/users/integrations/'+integrationForm.elements.service.value);
      if(version!==integrationVersion)return;
      integrationForm.elements.username.value=data.username;
      integrationSaved=!!data.configured;
      integrationStatus.textContent=data.configured?'Доступ сохранён. Пароль скрыт.':'Доступ ещё не настроен.';
    }catch(e){if(version===integrationVersion)integrationStatus.textContent=e.message;}
    finally{if(version===integrationVersion)setIntegrationBusy(false);}
  }
  integrationForm.elements.service.addEventListener('change',loadIntegration);
  for(const name of ['username','password']){
    integrationForm.elements[name].addEventListener('input',()=>{
      integrationSaved=false; updateIntegrationCheck();
      integrationStatus.textContent='Сначала сохраните изменения доступа.';
    });
  }
  integrationForm.addEventListener('submit',async event=>{
    event.preventDefault();if(integrationBusy)return;
    integrationSaved=false;setIntegrationBusy(true);
    try{
      await api('/api/users/integrations/'+integrationForm.elements.service.value,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:integrationForm.elements.username.value,password:integrationForm.elements.password.value})});
      integrationSaved=true;
      integrationForm.elements.password.value='';integrationStatus.textContent='Настройки доступа сохранены.';
      loadNotifications().catch(()=>{});
    }catch(e){integrationStatus.textContent=e.message;}
    finally{setIntegrationBusy(false);}
  });
  checkEddsLogin.addEventListener('click',async()=>{
    if(integrationBusy||!integrationSaved||integrationForm.elements.service.value!=='edds')return;
    setIntegrationBusy(true);integrationStatus.textContent='Проверяем вход в Добродел…';
    try{
      const data=await api('/api/users/integrations/edds/check',{method:'POST'});
      integrationStatus.textContent=data.message;
    }catch(e){integrationStatus.textContent=e instanceof TypeError?'Не удалось связаться с сервером. Повторите проверку.':e.message;}
    finally{setIntegrationBusy(false);}
  });
  loadIntegration();
})();
