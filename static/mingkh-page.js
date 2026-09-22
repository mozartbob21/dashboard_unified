var initialDates=JSON.parse(document.getElementById("mingkhDates").textContent);

var preset = 'thucur', inflight = 0;
var lastUpdated = null;      // метка портала на момент последней удачной загрузки
function $(id) { return document.getElementById(id); }

function iso(daysAgo) {
  var d = new Date(); d.setDate(d.getDate() - daysAgo);
  return d.toISOString().slice(0, 10);
}

function query(fresh) {
  var p = new URLSearchParams();
  if (fresh) p.set('fresh', '1');
  p.set('preset', preset);
  if (preset === 'custom') {
    p.set('curr_from', $('curr_from').value);
    p.set('curr_to', $('curr_to').value);
  }
  p.set('cmp', $('cmp').value);
  if ($('cmp').value === 'custom') {
    p.set('prev_from', $('prev_from').value);
    p.set('prev_to', $('prev_to').value);
  }
  return p.toString();
}

/** Сходить на портал за данными. Только при смене периода, по кнопке
 *  и по автообновлению — но НЕ при смене фильтров. */
function loadData(fresh) {
  var my = ++inflight;
  nextAt = null;
  $('status').className = '';
  $('status').textContent = 'Запрашиваю данные с портала…';
  $('out').classList.add('skel');
  fetch('/mingkh/api/dataset?' + query(fresh))
    .then(function (r) {
      // тело ответа несёт причину — без него на экране было голое «500»
      return r.json().then(function (d) {
        if (!r.ok || d.error) throw new Error(d.error || d.detail || ('сервер ответил ' + r.status));
        return d;
      }, function () { throw new Error('сервер ответил ' + r.status); });
    })
    .then(function (d) {
      if (my !== inflight) return;
      if (d.error) throw new Error(d.error);
      $('out').classList.remove('skel');
      $('lastload').textContent = d.updated;
      lastUpdated = d.updated;
      // коротко: период, с чем сравниваем, когда загружено
      var short = function (s) { return esc(s.replace(/^\S+ /, '').replace(/\.\d{4}/g, '')); };
      var msg = short(d.curr_label) + ' &middot; ' + d.counts.curr + ' обращений' +
                ' &middot; пред. ' + short(d.prev_label) + ' (' + d.counts.prev + ')' +
                ' &middot; АППГ (' + d.counts.appg + ')' +
                ' &middot; загружено ' + d.fetched_at.slice(0, 5);
      var bad = Object.keys(d.errors || {});
      if (bad.length) msg += ' &middot; не отдал портал: ' + bad.join(', ');
      $('status').innerHTML = msg + '<span id="calc"></span><span id="calcNote"></span>';
      $('out').hidden = false;
      window.DASH.setData(d);
      scheduleNext();
    })
    .catch(function (e) {
      if (my !== inflight) return;
      $('status').className = 'err';
      $('status').textContent = 'Не удалось получить данные. ' + e.message;
      $('out').classList.remove('skel');
      $('out').hidden = true;
      scheduleNext();
    });
}

function syncDateFields() {
  document.querySelector('.custom-dates').classList.toggle('hidden', preset !== 'custom' && $('cmp').value !== 'custom');
  document.querySelector('.current-dates').hidden = preset !== 'custom';
}

$('period').addEventListener('change', function () {
  preset = $('period').value;
  syncDateFields();
  if (preset !== 'custom' || ($('curr_from').value && $('curr_to').value)) loadData();
});

$('cmp').addEventListener('change', function () {
  var custom = $('cmp').value === 'custom';
  syncDateFields();
  document.querySelectorAll('.cmp-dates').forEach(function (x) {
    x.classList.toggle('hidden', !custom);
  });
  if (!custom || ($('prev_from').value && $('prev_to').value)) loadData();
});

['curr_from', 'curr_to', 'prev_from', 'prev_to'].forEach(function (id) {
  $(id).addEventListener('change', loadData);
});

/* ---- автообновление ---- */
var nextAt = null;

function parseHours() {
  return ($('hours').value || '').split(',')
    .map(function (s) { return parseInt(s.trim(), 10); })
    .filter(function (h) { return h >= 0 && h <= 23; })
    .sort(function (a, b) { return a - b; });
}

function scheduleNext() {
  var mode = $('auto').value;
  if (mode === 'off') { nextAt = null; tick(); return; }
  var now = new Date();
  if (mode === 'at') {
    var hs = parseHours();
    if (!hs.length) { nextAt = null; tick(); return; }
    var next = null;
    for (var i = 0; i < hs.length; i++) {
      var c = new Date(now); c.setHours(hs[i], 0, 0, 0);
      if (c > now) { next = c; break; }
    }
    if (!next) {
      next = new Date(now);
      next.setDate(next.getDate() + 1);
      next.setHours(hs[0], 0, 0, 0);
    }
    nextAt = next;
  } else {
    nextAt = new Date(now.getTime() + parseInt(mode, 10) * 3600 * 1000);
  }
  tick();
}

function pad(n) { return (n < 10 ? '0' : '') + n; }

/** Обновляем только если портал реально выложил новые данные.
 *  Метка берётся отдельным дешёвым запросом, а не полной выгрузкой. */
function autoRefresh() {
  nextAt = null;
  fetch('/mingkh/api/updated')
    .then(function (r) { if (!r.ok) throw new Error('Источник недоступен'); return r.json(); })
    .then(function (d) {
      if (d.updated && d.updated === lastUpdated) {
        $('timer').dataset.skipped = (+($('timer').dataset.skipped || 0)) + 1;
        if ($('calcNote')) $('calcNote').textContent = ' · источник не обновлялся с ' + d.updated +
          ', пропустили обновление';
        scheduleNext();
        return;
      }
      loadData(true);
    })
    .catch(function () { loadData(true); });   // не смогли проверить — обновляем как обычно
}

function tick() {
  if (!nextAt) { $('timer').textContent = 'автообновление выключено'; return; }
  var left = Math.round((nextAt - new Date()) / 1000);
  if (left <= 0) { autoRefresh(); return; }
  var h = Math.floor(left / 3600), m = Math.floor(left % 3600 / 60), s = left % 60;
  $('timer').textContent = 'обновление через ' + pad(h) + ':' + pad(m) + ':' + pad(s) +
    ' (в ' + pad(nextAt.getHours()) + ':' + pad(nextAt.getMinutes()) + ')';
}

setInterval(tick, 1000);

$('auto').addEventListener('change', function () {
  document.querySelector('.at-hours').classList.toggle('hidden', $('auto').value !== 'at');
  save(); scheduleNext();
});
$('hours').addEventListener('change', function () { save(); scheduleNext(); });
$('refresh').addEventListener('click', function () { loadData(true); });

function save() {
  try {
    localStorage.setItem('mingkh_auto', $('auto').value);
    localStorage.setItem('mingkh_hours', $('hours').value);
  } catch (e) { /* приватный режим — не запоминаем */ }
}

try {
  var a = localStorage.getItem('mingkh_auto');
  if (a) $('auto').value = a;
  var hh = localStorage.getItem('mingkh_hours');
  if (hh) $('hours').value = hh;
} catch (e) { /* нет доступа к хранилищу */ }
document.querySelector('.at-hours').classList.toggle('hidden', $('auto').value !== 'at');

$('curr_from').value = initialDates[0];
$('curr_to').value = initialDates[1];
$('prev_from').value = initialDates[2];
$('prev_to').value = initialDates[3];
loadData();
