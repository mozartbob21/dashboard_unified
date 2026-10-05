/* Shared Leaflet basemap for EDDS and the ministry complaints map.
 * https://yandex.ru/maps-api/docs/tiles-api/request.html
 * Browser caching stays enabled; only the visible map is requested.
 */
(() => {
  'use strict';
  const WAIT = 12000;

  function configuration() {
    let config = {};
    try { config = JSON.parse(document.getElementById('neurona-map-config')?.textContent || '{}'); }
    catch (_) { /* Standalone/older pages retain the default map. */ }
    config = config && typeof config === 'object' ? config : {};
    const key = typeof config.api_key === 'string' ? config.api_key.trim() : '';
    const yandex = config.provider === 'yandex' && !!key;
    return {
      yandex,
      notice: typeof config.notice === 'string' ? config.notice : '',
      url: yandex
        ? 'https://tiles.api-maps.yandex.ru/v1/tiles/?x={x}&y={y}&z={z}&lang=ru_RU&l=map&projection=web_mercator&maptype=map&apikey=' + encodeURIComponent(key)
        : 'https://tile.openstreetmap.org/{z}/{x}/{y}.png',
      attribution: yandex
        ? '© <a href="https://yandex.ru/maps/" target="_blank" rel="noopener noreferrer">Яндекс Карты</a>'
        : '© <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener noreferrer">OpenStreetMap</a> contributors'
    };
  }

  function addYandexLogo(map) {
    const logo = L.control({position: 'bottomleft'});
    logo.onAdd = () => {
      const box = L.DomUtil.create('div', 'neurona-map-logo');
      const link = L.DomUtil.create('a', '', box);
      link.href = 'https://yandex.ru/maps/';
      link.target = '_blank'; link.rel = 'noopener noreferrer';
      link.setAttribute('aria-label', 'Открыть Яндекс Карты');
      const img = L.DomUtil.create('img', '', link);
      img.src = '/static/vendor/yandex-map-logo.svg';
      img.alt = 'Яндекс';
      L.DomEvent.disableClickPropagation(box);
      L.DomEvent.disableScrollPropagation(box);
      return box;
    };
    logo.addTo(map);
    return logo;
  }

  function create(map, {onStatus = () => {}} = {}) {
    const config = configuration();
    const logo = config.yandex ? addYandexLogo(map) : null;
    const layer = L.tileLayer(config.url, {
      maxZoom: 18,
      attribution: config.attribution,
      // Request visible tiles after panning/zooming, without speculative buffers.
      keepBuffer: 0, updateWhenIdle: true, updateWhenZooming: false,
      crossOrigin: 'anonymous',
      // Provider restrictions use Referer. Send only the origin, never page/query data.
      referrerPolicy: 'strict-origin-when-cross-origin'
    });
    let timer = null, ok = 0, errors = 0, note = '', disposed = false;
    const stop = () => { clearTimeout(timer); timer = null; };
    const say = value => {
      const message = [config.notice, value].filter(Boolean).join(' ');
      if (!disposed && message !== note) { note = message; onStatus(message); }
    };
    const fail = () => say(config.yandex
      ? 'Подложка Яндекс Карт недоступна. Проверьте ключ Tiles API, его ограничения и доступ к tiles.api-maps.yandex.ru. Точки и данные сохранены.'
      : 'Подложка OpenStreetMap недоступна. Проверьте доступ к tile.openstreetmap.org. Точки и данные сохранены.');
    const start = () => {
      stop(); ok = 0; errors = 0;
      timer = setTimeout(() => { if (!ok) fail(); }, WAIT);
    };
    const handlers = {
      loading: start,
      tileload: () => { ok++; stop(); say(''); },
      tileerror: () => { errors++; },
      load: () => {
        stop();
        // Leaflet also emits "load" when every tile failed.
        if (ok) say(''); else if (errors) fail();
      },
      remove: stop
    };
    layer.on(handlers);
    say('');
    layer.addTo(map);
    return {
      layer,
      retry() { if (!disposed) { say(''); start(); layer.redraw(); } },
      destroy() {
        disposed = true; stop(); layer.off(handlers); map.removeLayer(layer);
        if (logo) logo.remove();
      }
    };
  }
  window.NeuronaBasemap = Object.freeze({create});
})();
