(() => {
  const button=document.getElementById('eddsRefresh'), message=document.getElementById('eddsStatus');
  const armButton=document.getElementById('eddsArmRefresh');
  let armStarted=false, complaintsRunning=false;
  if(!button) return;
  async function request(path,options={}){
    const r=await fetch('/edds/'+path,{cache:'no-store',...options});
    if(r.redirected||r.status===401) throw new Error('Войдите в систему заново');
    const data=await r.json();
    if(!r.ok) throw new Error(data.detail||'Ошибка запроса');
    return data;
  }
  async function poll(){
    try{const data=await request('status'); button.disabled=data.running||!data.complaints_configured;
      button.title=data.complaints_configured?'Обновить свод жалоб':'Настройте отдельный доступ к Доброделу в разделе «Пользователи»';
      const text=data.message||(data.arm_configured?`АРМ ЕДДС: учётные данные сохранены · ${data.arm_transport==='chrome'?'Chrome на сервере':'прямое подключение'}.`:'Настройте доступ к «АРМ ЕДДС» в разделе «Пользователи → Логины и пароли».');
      if(message.textContent!==text)message.textContent=text;
      if(complaintsRunning&&!data.running)window.dispatchEvent(new Event('edds-complaints-updated'));
      complaintsRunning=data.running;
      if(data.arm_configured&&!armStarted&&typeof window.eddsLoadArm==='function'){
        armStarted=true;window.eddsLoadArm();
      }
    }catch(e){message.textContent=e.message;}
  }
  button.addEventListener('click',async()=>{
    button.disabled=true;
    try{const data=await request('refresh',{method:'POST'});message.textContent=data.message;}
    catch(e){message.textContent=e.message;button.disabled=false;}
  });
  if(armButton)armButton.addEventListener('click',()=>{if(typeof window.eddsLoadArm==='function')window.eddsLoadArm();});
  poll(); setInterval(()=>{if(!document.hidden)poll();},15000);
})();
