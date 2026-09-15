(() => {
  const buttons=[...document.querySelectorAll('[data-module-view]')];
  function setView(view){
    view=view==='list'?'list':'cards';
    document.documentElement.dataset.moduleView=view;
    buttons.forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.moduleView===view)));
    try{localStorage.setItem('neurona-module-view',view);}catch(_){}
  }
  let initial='cards';try{initial=localStorage.getItem('neurona-module-view');}catch(_){}
  setView(initial);
  buttons.forEach(b=>b.addEventListener('click',()=>setView(b.dataset.moduleView)));
})();
