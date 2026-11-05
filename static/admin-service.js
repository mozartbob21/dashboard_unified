(() => {
  const panel = document.querySelector('[data-admin-panel="service"]');
  if (!panel) return;
  const metrics = document.getElementById('serviceMetrics');
  const button = document.getElementById('serviceRefresh');
  const stamp = document.getElementById('serviceStamp');
  const format = new Intl.NumberFormat('ru-RU', {maximumFractionDigits:1});
  const unknown = 'Нет данных';
  const gib = value => Number.isFinite(value) ? format.format(value / 2**30) + ' ГиБ' : unknown;
  const capacity = (used,total) => Number.isFinite(used) && Number.isFinite(total) ? gib(used)+' из '+gib(total) : unknown;
  function elapsed(seconds) {
    if (!Number.isFinite(seconds)) return unknown;
    if (seconds < 60) return Math.floor(seconds) + ' с';
    return Math.floor(seconds / 3600) + ' ч ' + Math.floor(seconds % 3600 / 60) + ' мин';
  }
  function render(data) {
    const fields = [
      ['Процессор', Number.isFinite(data.cpu_percent) ? format.format(data.cpu_percent)+' %' : unknown],
      ['Оперативная память', capacity(data.ram_used, data.ram_total)],
      ['Диск', capacity(data.disk_used, data.disk_total)],
      ['Приложение', data.application || unknown],
      ['Время работы процесса', elapsed(data.uptime_seconds)],
      ['База данных', data.database || unknown],
      ['Задачи ИИ-чата', data.chat_jobs ? `В очереди: ${data.chat_jobs.queued} · выполняются: ${data.chat_jobs.running}` : unknown]
    ];
    metrics.replaceChildren(...fields.map(([label,value])=>{
      const box=document.createElement('div'), name=document.createElement('dt'), detail=document.createElement('dd');
      name.textContent=label; detail.textContent=value; box.append(name,detail); return box;
    }));
    document.getElementById('serviceScope').textContent=data.scope || '';
  }
  async function refresh(){
    if(panel.hidden || document.hidden || button.disabled) return;
    button.disabled=true;
    try {
      const response=await fetch('/api/users/service-center',{credentials:'same-origin',cache:'no-store', signal:globalThis.AbortSignal?.timeout?.(12000)});
      if(!response.ok || response.redirected) throw new Error(response.status===403?'Доступ отозван.':'Не удалось обновить показатели.');
      const data=await response.json();
      render(data);
      stamp.textContent='Обновлено '+new Date(data.updated_at).toLocaleString('ru-RU')+' · автоматически каждые 15 секунд';
    } catch(error) {
      render({});
      stamp.textContent=error.message=== 'Доступ отозван.'?error.message:'Нет связи с сервисным центром. Повторим через 15 секунд.';
    } finally {button.disabled=false;}
  }
  button.addEventListener('click',refresh);
  document.addEventListener('admin:sectionchange',refresh);
  document.addEventListener('visibilitychange',refresh);
  window.addEventListener('online',refresh);
  setInterval(refresh,15000);
  refresh();
})();
