// One recent signal from any visible tab is enough for the account to be online.
(() => {
  if (window.top !== window.self) return;
  let busy = false;
  async function ping() {
    if (document.hidden || busy) return;
    busy = true;
    try {
      const response = await fetch('/api/me/presence', {method:'POST', credentials:'same-origin', cache:'no-store', signal:globalThis.AbortSignal?.timeout?.(12000)});
      if (response.status === 401 || response.status === 403) clearInterval(timer);
    } catch (_) { /* Next tick retries after a lost connection. */ }
    finally { busy = false; }
  }
  const timer = setInterval(ping, 30000);
  document.addEventListener('visibilitychange', ping);
  window.addEventListener('online', ping);
  ping();
})();
