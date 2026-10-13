/* Facts are model-generated in a separate server cache, never chat messages. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.NeuronaWaiting = api;
})(typeof window === 'undefined' ? globalThis : window, function () {
  'use strict';
  const DELAY_MS = 60000, ROTATE_MS = 30000;
  function create(env, host) {
    let items = [], busy = false, dismissed = false, ready = false, index = -1;
    let delay = null, rotation = null, request = null, deadline = null, generation = 0;
    const topic = host.querySelector('[data-fact-topic]');
    const text = host.querySelector('[data-fact-text]');
    const close = host.querySelector('[data-fact-close]');
    function hide() { host.hidden = true; }
    function showNext() {
      if (!busy || !ready || dismissed || env.document.hidden || !items.length) return;
      index = (index + 1) % items.length;
      topic.textContent = items[index].topic;
      text.textContent = items[index].text;
      host.hidden = false;
    }
    async function refresh(warm = false) {
      if (request) return;
      const version = generation;
      const controller = new env.AbortController();
      request = controller;
      deadline = env.setTimeout(() => controller.abort(), 8000);
      try {
        const response = await env.fetch('/aichat/api/waiting-facts' + (warm ? '' : '?warm=false'), {
          credentials:'same-origin', cache:'no-store', signal:controller.signal,
          headers:{Accept:'application/json'}
        });
        if (!response.ok || response.redirected) return;
        const data = await response.json();
        if (generation !== version) return;
        const next = Array.isArray(data.items) ? data.items.filter(item => item &&
          ['ИИ','ЖКХ'].includes(item.topic) && typeof item.text === 'string' &&
          item.text.trim().length > 0 && item.text.length <= 320).slice(0, 8) : [];
        if (next.length) items = next;
        if (host.hidden) showNext();
      } catch (_) { /* Optional facts must never interrupt the main answer. */ }
      finally {
        if (request === controller) {
          env.clearTimeout(deadline);deadline = null;request = null;
        }
      }
    }
    function cancelFetch() {
      generation++;request?.abort();request = null;
      env.clearTimeout(deadline);deadline = null;
    }
    function stop() {
      busy = false;ready = false;
      env.clearTimeout(delay);env.clearInterval(rotation);
      delay = null;rotation = null;cancelFetch();hide();
    }
    function start() {
      stop();busy = true;dismissed = false;index = -1;
      // This only reads/warms a shared batch; it does not submit the user's text.
      refresh();
      delay = env.setTimeout(() => {
        ready = true;showNext();
        if (!items.length) refresh();
        rotation = env.setInterval(() => {
          if (env.document.hidden || dismissed) return;
          if (items.length) showNext();else refresh();
        }, ROTATE_MS);
      }, DELAY_MS);
    }
    function dismiss() {
      dismissed = true;hide();env.clearInterval(rotation);rotation = null;cancelFetch();
    }
    function visible() { if (!env.document.hidden && host.hidden) showNext(); }
    function destroy() {
      stop();close.removeEventListener('click', dismiss);
      env.document.removeEventListener('visibilitychange', visible);
      env.removeEventListener('pagehide', stop);
    }
    close.addEventListener('click', dismiss);
    env.document.addEventListener('visibilitychange', visible);
    env.addEventListener('pagehide', stop);
    hide();
    return {start, stop, warm:() => refresh(true), destroy};
  }
  return {create, DELAY_MS, ROTATE_MS};
});
