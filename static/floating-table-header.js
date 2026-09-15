(() => {
  'use strict';
  const headers = [...document.querySelectorAll('[data-floating-header]')].map(wrap => {
    const source = wrap.querySelector('table');
    if (!source?.tHead) return null;
    const layer = document.createElement('div');
    layer.className = 'floating-table-header';
    layer.setAttribute('aria-hidden', 'true');
    layer.hidden = true;
    const table = document.createElement('table');
    const head = source.tHead.cloneNode(true);
    head.querySelectorAll('[id]').forEach(node => node.removeAttribute('id'));
    table.append(head);
    layer.append(table);
    document.body.append(layer);
    return {wrap, source, layer, table, head, dirty: true};
  }).filter(Boolean);
  if (!headers.length) return;
  let frame = 0;
  function schedule() {
    if (!frame) frame = requestAnimationFrame(update);
  }
  function update() {
    frame = 0;
    for (const item of headers) {
      const {wrap, source, layer, table, head} = item;
      const rect = source.getBoundingClientRect();
      const original = source.tHead.getBoundingClientRect();
      const viewport = wrap.getBoundingClientRect();
      const visible = original.top < 0 && rect.bottom > 0 && viewport.right > 0 && viewport.left < innerWidth;
      layer.hidden = !visible;
      if (!visible) continue;
      if (item.dirty) {
        const cells = [...source.tHead.querySelectorAll('th,td')];
        head.querySelectorAll('th,td').forEach((cell, i) => {
          cell.style.width = cells[i].getBoundingClientRect().width + 'px';
        });
        table.style.width = rect.width + 'px';
        head.style.height = original.height + 'px';
        item.dirty = false;
      }
      const left = Math.max(0, viewport.left + wrap.clientLeft);
      const right = Math.min(innerWidth, viewport.left + wrap.clientLeft + wrap.clientWidth);
      layer.style.left = left + 'px';
      layer.style.width = Math.max(0, right - left) + 'px';
      layer.style.top = Math.min(0, rect.bottom - original.height) + 'px';
      table.style.marginLeft = (rect.left - left) + 'px';
    }
  }
  const resize = new ResizeObserver(() => {
    headers.forEach(item => item.dirty = true);
    schedule();
  });
  headers.forEach(({wrap, source}) => {resize.observe(wrap); resize.observe(source);});
  document.addEventListener('scroll', schedule, {passive: true, capture: true});
  window.addEventListener('resize', schedule, {passive: true});
  schedule();
})();
