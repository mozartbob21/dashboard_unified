(() => {
  const body=document.body;
  if(!body||body.classList.contains('home-page')) return;
  body.classList.add('unified-module-page');
  let topbar=document.querySelector('header.topbar,.topbar,.top,body > header');
  if(!topbar){
    topbar=document.createElement('header');
    topbar.className='topbar unified-shell-topbar shell-generated';
    const subtitle=document.title.replace(/\s+[—·-]\s+Нейрона.*$/i,'');
    topbar.innerHTML=`<a class="system-shell-brand" href="/"><img class="brand-logo" src="/static/logo.svg" alt="Нейрона ИИ"><span class="system-shell-brand-copy"><strong class="brand-title">Нейрона ИИ</strong><span>${subtitle}</span></span></a><div class="system-shell-actions"></div>`;
    body.insertBefore(topbar,body.firstChild);
  }else topbar.classList.add('unified-shell-topbar');
  if(topbar.tagName==='HEADER'&&topbar.parentElement===body&&!topbar.querySelector('.brand,.system-shell-brand')){
    const label=topbar.querySelector(':scope > span');
    const brand=document.createElement('div');
    brand.className='brand';
    brand.innerHTML=`<a class="brand-link" href="/"><img class="brand-logo" src="/static/logo.svg" alt="Нейрона ИИ"><span class="system-shell-brand-copy"><strong class="brand-title">Нейрона ИИ</strong><span class="brand-subtitle">${label?.textContent||document.title}</span></span></a>`;
    topbar.insertBefore(brand,topbar.firstChild);
    if(label) label.remove();
    const host=document.createElement('div');host.className='system-shell-actions';
    [...topbar.querySelectorAll(':scope > a,:scope > button')].forEach(item=>host.appendChild(item));
    topbar.appendChild(host);
  }
  const brandTitle=topbar.querySelector('.brand-title');
  if(brandTitle) brandTitle.textContent='Нейрона ИИ';
  let actions=topbar.querySelector('.top-actions,.topbar-right,.actions,.nav,.system-shell-actions');
  if(!actions){actions=document.createElement('div');actions.className='system-shell-actions';topbar.appendChild(actions)}
  const internal=[...topbar.querySelectorAll('a[href]')].filter(link=>!link.closest('.brand,.system-shell-brand'));
  let back=internal.find(link=>{
    const href=link.getAttribute('href')||'',text=(link.textContent||'').toLowerCase();
    return href==='/'||text.includes('главн')||text.includes('назад')||text.includes('систем');
  });
  if(!back){back=document.createElement('a');back.href='/';actions.appendChild(back)}
  if((back.getAttribute('href')||'')==='/') back.textContent='← К модулям';
  back.classList.add('shell-control','system-back-button');
  topbar.querySelectorAll(':scope > a,:scope > button').forEach(control=>control.classList.add('shell-control'));
  actions.querySelectorAll(':scope > a,:scope > button,:scope > .theme-dropdown > button').forEach(control=>control.classList.add('shell-control'));
  topbar.dataset.shellReady='true';
})();
