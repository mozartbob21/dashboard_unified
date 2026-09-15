(() => {
  const button=document.getElementById('eddsRefresh'), message=document.getElementById('eddsStatus');
  if(!button) return;
  async function request(path,options={}){
    const r=await fetch('/edds/'+path,options);
    if(r.redirected||r.status===401) throw new Error('Войдите в систему заново');
    const data=await r.json();
    if(!r.ok) throw new Error(data.detail||'Ошибка запроса');
    return data;
  }
  async function poll(){
    try{const data=await request('status'); button.disabled=data.running;
      const text=data.message||'Данные АРМ загрузите из файла; жалобы доступны из серверного свода.';
      if(message.textContent!==text)message.textContent=text;
    }catch(e){message.textContent=e.message;}
  }
  button.addEventListener('click',async()=>{
    button.disabled=true;
    try{const data=await request('refresh',{method:'POST'});message.textContent=data.message;}
    catch(e){message.textContent=e.message;button.disabled=false;}
  });
  poll(); setInterval(()=>{if(!document.hidden)poll();},15000);
})();
