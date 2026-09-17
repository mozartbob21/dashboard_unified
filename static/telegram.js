(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = {stage: 'disconnected', chats: [], chat: null, messages: [], before: 0,
    busy: false, loading: false, epoch: 0, reply: null, attempt: null, dialogs: new Map(), selected: new Set(), drafts: new Map()};
  const kind = value => ({user: 'Личный чат', group: 'Группа', channel: 'Канал'}[value] || 'Чат');
  function element(tag, text, className) {
    const node = document.createElement(tag);
    if (text != null) node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  function notice(text = '', error = false) {
    $('tgNotice').textContent = text;
    $('tgNotice').hidden = !text;
    $('tgNotice').dataset.error = String(error);
  }
  function requestId() {
    // crypto.randomUUID is unavailable on some HTTP LAN addresses.
    if (crypto.randomUUID) return crypto.randomUUID();
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 15) | 64; bytes[8] = (bytes[8] & 63) | 128;
    const hex = [...bytes].map(n => n.toString(16).padStart(2, '0')).join('');
    return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
  }
  async function api(path, method = 'GET', body = null, binary = false) {
    const response = await fetch('/telegram/api' + path, {method, cache: 'no-store',
      headers: {'X-Neurona-Telegram': '1', ...(body !== null ? {'Content-Type': 'application/json'} : {})},
      body: body !== null ? JSON.stringify(body) : undefined});
    if (response.redirected) {
      throw new Error('Сессия «Нейроны» завершена или сервер недоступен. Обновите страницу и войдите снова.');
    }
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      const detail = typeof data.detail === 'string' ? data.detail : 'Проверьте введённые данные.';
      throw new Error(detail);
    }
    if (!response.headers.get('content-type')?.includes(binary ? 'octet-stream' : 'json')) {
      throw new Error('Сервер вернул неожиданный ответ. Обновите страницу.');
    }
    return binary ? response.blob() : response.json();
  }
  async function action(button, operation) {
    if (state.busy || state.loading) {notice('Дождитесь завершения текущего запроса.');return;}
    state.busy = true;
    if (button) button.disabled = true;
    notice();
    try {await operation();}
    catch (error) {notice(error.message, true);}
    finally {state.busy = false; if (button) button.disabled = false;}
  }
  function applyStatus(data) {
    state.stage = data.stage;
    state.chats = data.chats || [];
    const ready = state.stage === 'ready';
    $('tgSetup').hidden = data.configured;
    $('tgLogin').hidden = !data.configured || ready;
    $('tgWorkspace').hidden = !ready;
    $('tgChoose').hidden = !ready;
    $('tgDisconnect').hidden = !ready && state.stage !== 'expired';
    $('tgProfile').textContent = ready ? data.profile?.name || 'Telegram подключён' : '';
    $('tgPhoneField').hidden = ['code', 'password'].includes(state.stage);
    $('tgCodeField').hidden = state.stage !== 'code';
    $('tgPasswordField').hidden = state.stage !== 'password';
    $('tgRestart').hidden = !['code', 'password'].includes(state.stage);
    $('tgLoginSubmit').textContent = state.stage === 'code' ? 'Подтвердить код' : state.stage === 'password' ? 'Войти' : 'Получить код';
    $('tgLoginHint').textContent = state.stage === 'code' ? 'Введите код, который Telegram отправил в приложение или по SMS.' :
      state.stage === 'password' ? 'Для вашего аккаунта включена двухэтапная защита. Введите пароль Telegram.' :
      'Подключите свой аккаунт и выберите чаты для работы. Подключение принадлежит только вашей учётной записи «Нейроны».';
    renderChats();
    if (!state.chats.some(chat => chat.id === state.chat?.id)) clearChat();
  }
  function clearChat() {
    state.chat = null;state.messages = [];state.epoch++;state.attempt = null;state.reply = null;
    $('tgMessages').replaceChildren(element('p', 'Выберите чат из списка или добавьте его через «Выбрать чаты».', 'tg-empty'));
    $('tgChatTitle').textContent = 'Выберите чат';
    $('tgComposer').hidden = true;$('tgRefresh').hidden = true;$('tgOlder').hidden = true;
    $('tgText').value = '';updateReply();
  }
  function renderChats() {
    const query = $('tgChatSearch').value.trim().toLocaleLowerCase('ru');
    $('tgChats').replaceChildren();
    for (const chat of state.chats.filter(chat => chat.title.toLocaleLowerCase('ru').includes(query))) {
      const button = element('button', chat.title, 'tg-chat');
      button.type = 'button';button.setAttribute('aria-current', String(state.chat?.id === chat.id));
      button.append(element('small', kind(chat.kind)));
      button.onclick = () => action(button, async () => {
        if (state.chat) state.drafts.set(state.chat.id, $('tgText').value);
        state.chat = chat;state.epoch++;state.messages = [];state.reply = null;state.attempt = null;
        $('tgText').value = state.drafts.get(chat.id) || '';
        $('tgChatTitle').textContent = chat.title;
        $('tgComposer').hidden = false;$('tgRefresh').hidden = false;
        $('tgMessages').replaceChildren(element('p', 'Загружаю сообщения…', 'tg-empty'));
        updateReply();renderChats();await loadMessages(false, true);
      });
      $('tgChats').append(button);
    }
  }
  function updateReply() {
    $('tgReply').hidden = !state.reply;
    $('tgReplyText').textContent = state.reply ? `Ответ: ${state.reply.sender} · ${state.reply.text.slice(0, 100) || 'Вложение'}` : '';
  }
  function renderMessages(toBottom = false) {
    const container = $('tgMessages');
    const bottom = container.scrollHeight - container.scrollTop - container.clientHeight < 90;
    const scroll = container.scrollTop;
    container.replaceChildren();
    if (!state.messages.length) container.append(element('p', 'В этом чате пока нет сообщений.', 'tg-empty'));
    for (const message of state.messages) {
      const article = element('article', null, 'tg-message' + (message.out ? ' out' : ''));
      article.append(element('header', message.out ? 'Вы' : message.sender));
      if (message.reply_to) article.append(element('div', `Ответ на сообщение №${message.reply_to}`, 'tg-muted tg-small'));
      article.append(element('div', message.text || (message.service ? 'Служебное сообщение' : message.unsupported_media ? 'Этот тип сообщения пока не поддерживается.' : ''), 'tg-message-text'));
      if (message.file) {
        const file = element('button', `↓ ${message.file.name} · ${Math.ceil(message.file.size / 1024)} КБ`, 'tg-button tg-file');
        file.type = 'button';file.disabled = !message.file.downloadable;
        if (!message.file.downloadable) file.title = 'Поддерживаются вложения до 20 МБ';
        file.onclick = () => action(file, async () => {
          const blob = await api(`/chats/${encodeURIComponent(state.chat.id)}/attachments/${message.id}`, 'GET', null, true);
          const url = URL.createObjectURL(blob), link = document.createElement('a');
          link.href = url;link.download = message.file.name;link.click();setTimeout(() => URL.revokeObjectURL(url), 10000);
        });
        article.append(file);
      }
      const footer = element('footer');
      footer.append(element('time', new Date(message.date).toLocaleString('ru-RU', {day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'})));
      const reply = element('button', 'Ответить');reply.type = 'button';
      reply.onclick = () => {state.reply = message;updateReply();$('tgText').focus();};
      footer.append(reply);article.append(footer);container.append(article);
    }
    if (toBottom || bottom) container.scrollTop = container.scrollHeight;
    else container.scrollTop = scroll;
  }
  async function loadMessages(older = false, toBottom = false) {
    if (!state.chat || state.loading) return;
    state.loading = true;
    const epoch = state.epoch, chat = state.chat.id;
    try {
      const before = older && state.messages.length ? state.messages[0].id : 0;
      const data = await api(`/chats/${encodeURIComponent(chat)}/messages${before ? '?before=' + before : ''}`);
      if (epoch !== state.epoch) return;
      let retained = state.messages;
      if (!older) {
        const first = data.items[0]?.id ?? Infinity;
        retained = data.items.length ? state.messages.filter(item => item.id < first) : [];
      }
      const merged = new Map(retained.map(item => [item.id,item]));
      data.items.forEach(item => merged.set(item.id,item));
      state.messages = [...merged.values()].sort((a,b) => a.id-b.id);
      if (older || !before && retained.length === 0) $('tgOlder').hidden = !data.has_more;
      const oldHeight = $('tgMessages').scrollHeight, oldScroll = $('tgMessages').scrollTop;
      renderMessages(toBottom);
      if (older) $('tgMessages').scrollTop = oldScroll + $('tgMessages').scrollHeight - oldHeight;
    } finally {state.loading = false;}
  }
  function renderDialogs() {
    const query = $('tgDialogSearch').value.trim().toLocaleLowerCase('ru');
    $('tgDialogList').replaceChildren();
    for (const chat of state.dialogs.values()) {
      if (!chat.title.toLocaleLowerCase('ru').includes(query)) continue;
      const row = element('label', null, 'tg-dialog-row'), checkbox = document.createElement('input');
      checkbox.type = 'checkbox';checkbox.checked = state.selected.has(chat.id);
      checkbox.onchange = () => {
        checkbox.checked ? state.selected.add(chat.id) : state.selected.delete(chat.id);
        $('tgSelectedCount').textContent = `Выбрано: ${state.selected.size} из 100`;
      };
      row.append(checkbox, element('span', chat.title + ' · ' + kind(chat.kind)));$('tgDialogList').append(row);
    }
    $('tgSelectedCount').textContent = `Выбрано: ${state.selected.size} из 100`;
  }
  async function loadDialogs(more = false) {
    $('tgPickerNotice').textContent = 'Загружаю список…';
    try {
      const data = await api('/dialogs' + (more ? '?more=true' : ''));
      for (const chat of data.items) state.dialogs.set(chat.id, chat);
      $('tgMoreDialogs').hidden = !data.has_more;
      $('tgPickerNotice').textContent = 'Загружено чатов: ' + state.dialogs.size;
      renderDialogs();
    } catch (error) {$('tgPickerNotice').textContent = error.message;throw error;}
  }
  function chooseChats() {
    action($('tgChoose'), async () => {
      state.dialogs = new Map(state.chats.map(chat => [chat.id,chat]));
      state.selected = new Set(state.chats.map(chat => chat.id));
      $('tgDialogSearch').value = '';$('tgMoreDialogs').hidden = true;
      renderDialogs();$('tgPicker').showModal();await loadDialogs();
    });
  }
  $('tgChoose').onclick = chooseChats;$('tgChooseEmpty').onclick = chooseChats;
  $('tgPickerClose').onclick = () => {if (!state.busy) $('tgPicker').close();};
  $('tgPicker').addEventListener('cancel', event => {if (state.busy) event.preventDefault();});
  $('tgDialogSearch').oninput = renderDialogs;
  $('tgMoreDialogs').onclick = () => action($('tgMoreDialogs'), () => loadDialogs(true));
  $('tgPickerForm').onsubmit = event => {
    event.preventDefault();
    action($('tgSaveChats'), async () => {
      if (state.selected.size > 100) throw new Error('Можно выбрать не больше 100 чатов.');
      applyStatus(await api('/chats', 'PUT', {ids:[...state.selected]}));$('tgPicker').close();
    });
  };
  $('tgLoginForm').onsubmit = event => {
    event.preventDefault();
    action($('tgLoginSubmit'), async () => {
      if (['code','password'].includes(state.stage)) {
        const payload = state.stage === 'code' ? {code:$('tgCode').value.trim()} : {password:$('tgPassword').value};
        try {await api('/auth/confirm', 'POST', payload);}
        finally {$('tgCode').value = '';$('tgPassword').value = '';}
      } else await api('/auth/start', 'POST', {phone:$('tgPhone').value.trim()});
      applyStatus(await api('/status'));
      if (state.stage === 'ready') notice('Аккаунт подключён. Нажмите «Выбрать чаты», чтобы начать работу.');
    });
  };
  async function disconnect() {
    await api('/disconnect', 'POST');
    state.drafts.clear();$('tgPhone').value = '';$('tgCode').value = '';$('tgPassword').value = '';
    clearChat();applyStatus(await api('/status'));
  }
  $('tgRestart').onclick = () => action($('tgRestart'), disconnect);
  $('tgDisconnect').onclick = () => {
    if (confirm('Отключить Telegram от вашей учётной записи «Нейроны»? Эта Telegram-сессия будет завершена.')) action($('tgDisconnect'), disconnect);
  };
  $('tgChatSearch').oninput = renderChats;
  $('tgRefresh').onclick = () => action($('tgRefresh'), () => loadMessages());
  $('tgOlder').onclick = () => action($('tgOlder'), () => loadMessages(true));
  $('tgCancelReply').onclick = () => {state.reply = null;updateReply();};
  $('tgComposer').onsubmit = event => {
    event.preventDefault();
    action($('tgSend'), async () => {
      const text = $('tgText').value;
      if (!state.chat || !text.trim()) return;
      const fingerprint = JSON.stringify([state.chat.id,text,state.reply?.id || null]);
      if (state.attempt?.fingerprint !== fingerprint) state.attempt = {fingerprint,id:requestId()};
      $('tgText').readOnly = true;
      try {
        await api(`/chats/${encodeURIComponent(state.chat.id)}/messages`, 'POST', {
          text, request_id:state.attempt.id, reply_to:state.reply?.id || null});
        $('tgText').value = '';state.drafts.delete(state.chat.id);state.attempt = null;state.reply = null;updateReply();
        notice('Сообщение отправлено.');await loadMessages(false,true);
      } finally {$('tgText').readOnly = false;}
    });
  };
  $('tgText').addEventListener('keydown', event => {
    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {event.preventDefault();$('tgComposer').requestSubmit();}
  });
  setInterval(() => {
    if (document.hidden || state.busy || state.loading || !state.chat || $('tgPicker').open) return;
    loadMessages().catch(error => notice(error.message,true));
  }, 15000);
  action(null, async () => {applyStatus(await api('/status'));});
})();
