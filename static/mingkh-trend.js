'use strict';

/* ---------- график динамики ---------- */

// С чем сравнивать: тот же период год назад или предыдущий период
var trendCompare = 'appg';
var MONTHS = ['янв', 'фев', 'мар', 'апр', 'май', 'июн',
              'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
var WDAYS = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
var DAY = 86400000;

function pad2(n) { return (n < 10 ? '0' : '') + n; }

/** 'ГГГГ-ММ-ДД' <-> миллисекунды UTC: без часовых поясов и перевода часов. */
function parseDay(s) {
  var p = s.split('-');
  return Date.UTC(+p[0], +p[1] - 1, +p[2]);
}
function fmtDay(ms) {
  var d = new Date(ms);
  return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
}

/** Та же дата год спустя. 29 февраля превращается в 28-е. */
function plusYear(s) {
  var p = s.split('-'), y = +p[0] + 1, m = +p[1];
  var last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + pad2(m) + '-' + pad2(Math.min(+p[2], last));
}

/** Шаг графика по длине периода: до двух месяцев — дни, до двух лет — месяцы. */
function trendGrain(from, to) {
  var days = (parseDay(to) - parseDay(from)) / DAY + 1;
  if (days <= 62) return 'day';
  if (days <= 731) return 'month';
  return 'year';
}

function bucketKey(s, grain) {
  return grain === 'day' ? s : grain === 'month' ? s.slice(0, 7) : s.slice(0, 4);
}

function bucketLabel(key, grain, i, n) {
  if (grain === 'year') return key;
  if (grain === 'month') {
    var m = +key.slice(5, 7);
    // год подписываем на первом столбике и на январе — дальше и так ясно
    return MONTHS[m - 1] + (i === 0 || m === 1 ? ' ' + key.slice(0, 4) : '');
  }
  var d = new Date(parseDay(key));
  var dm = pad2(d.getUTCDate()) + '.' + pad2(d.getUTCMonth() + 1);
  return n <= 14 ? WDAYS[d.getUTCDay()] + ' ' + dm : dm;
}

/** Ряды графика с учётом всех фильтров. Сравнение сдвигается на даты текущего
 *  периода: год назад — на год, предыдущий период — на разницу их начал. */
function trendSeries() {
  var b = data.bounds, grain = trendGrain(b.curr[0], b.curr[1]);
  var keys = [], seen = {};
  for (var t = parseDay(b.curr[0]); t <= parseDay(b.curr[1]); t += DAY) {
    var k = bucketKey(fmtDay(t), grain);
    if (!seen[k]) { seen[k] = true; keys.push(k); }
  }

  var cmpPeriod = trendCompare === 'appg' ? 'appg' : 'prev';
  var move;
  if (cmpPeriod === 'appg') {
    move = plusYear;
  } else {
    var off = (parseDay(b.curr[0]) - parseDay(b.prev[0])) / DAY;
    move = function (s) { return fmtDay(parseDay(s) + off * DAY); };
  }

  var di = DIMS.indexOf('date');
  var missing = {}, outside = {};
  function count(period, shift) {
    var out = {}, rows = data.rows[period] || [];
    missing[period] = 0;
    outside[period] = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!matches(r)) continue;
      var s = (data.dims.date || [])[r[di]];
      if (!s) { missing[period]++; continue; }
      if (shift) s = shift(s);
      // Ограничиваем и неполные месяцы точными границами текущего окна.
      if (s < b.curr[0] || s > b.curr[1]) { outside[period]++; continue; }
      var key = bucketKey(s, grain);
      out[key] = (out[key] || 0) + 1;
    }
    return out;
  }
  return {
    grain: grain, keys: keys, cmpPeriod: cmpPeriod,
    cur: count('curr'), cmp: count(cmpPeriod, move),
    missing: missing, outside: outside,
    // итоги считаем по строкам, а не по столбикам: если окно сравнения задано
    // вручную и длиннее текущего, часть его дней на график не ляжет
    curTotal: total('curr'), cmpTotal: total(cmpPeriod)
  };
}

/** Круглый шаг сетки: 1, 2, 5, 10, 20, 50… */
function niceStep(v) {
  if (v <= 1) return 1;
  var p = Math.pow(10, Math.floor(Math.log(v) / Math.LN10)), m = v / p;
  return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p;
}

/** Столбик со скруглённой вершиной; основание прямое — стоит на оси. */
function barPath(x, y, w, h) {
  if (h <= 0) return '';
  var r = Math.min(3, w / 2, h);
  return 'M' + x + ',' + (y + h) + 'V' + (y + r) +
    'Q' + x + ',' + y + ' ' + (x + r) + ',' + y +
    'H' + (x + w - r) + 'Q' + (x + w) + ',' + y + ' ' + (x + w) + ',' + (y + r) +
    'V' + (y + h) + 'Z';
}

/** Подпись изменения для SVG: цвет по смыслу, как в таблицах. */
function deltaSvg(cur, prev) {
  if (!prev) return { text: cur ? 'новые' : '', cls: 't-flat' };
  var d = cur - prev;
  if (!d) return { text: '0', cls: 't-flat' };
  return {
    text: (d > 0 ? '▲' : '▼') + fmt(Math.abs(d)) + ' (' + pctText(d / prev * 100) + '%)',
    cls: d > 0 ? 't-bad' : 't-good'
  };
}

function renderTrend() {
  var box = $('trend');
  if (!box) return;
  if (!data || !data.bounds) { box.innerHTML = ''; return; }
  var s = trendSeries();
  var n = s.keys.length;
  var cmpName = s.cmpPeriod === 'appg' ? 'год назад' : 'предыдущий период';
  var grainName = { day: 'по дням', month: 'по месяцам', year: 'по годам' }[s.grain];
  var dTotal = deltaHtml(s.curTotal, s.cmpTotal);
  var notes = [];
  if (s.missing.curr || s.missing[s.cmpPeriod]) {
    notes.push('Без даты: текущий период — ' + fmt(s.missing.curr) +
      ', ' + cmpName + ' — ' + fmt(s.missing[s.cmpPeriod]) +
      '. Эти обращения входят в итоги, но не распределены по столбцам.');
  }
  if (s.outside.curr || s.outside[s.cmpPeriod]) {
    notes.push('За границами графика: текущий период — ' + fmt(s.outside.curr) +
      ', ' + cmpName + ' — ' + fmt(s.outside[s.cmpPeriod]) +
      '. Итоги включают весь выбранный период.');
  }

  box.innerHTML =
    '<section class="card trendcard"><div class="trendhead">' +
    '<h2>Динамика обращений</h2>' +
    '<div class="seg" role="group" aria-label="С чем сравнивать">' +
    '<button type="button" class="pre' + (trendCompare === 'appg' ? ' on' : '') +
    '" aria-pressed="' + (trendCompare === 'appg') + '" data-cmp="appg">год назад</button>' +
    '<button type="button" class="pre' + (trendCompare === 'prev' ? ' on' : '') +
    '" aria-pressed="' + (trendCompare === 'prev') + '" data-cmp="prev">предыдущий период</button></div></div>' +
    '<div class="trendsum">' +
    '<span class="lgi"><span class="sw sw-cur"></span>поступило <b>' + fmt(s.curTotal) + '</b></span>' +
    '<span class="lgi"><span class="sw sw-cmp"></span>' + cmpName + ' <b>' + fmt(s.cmpTotal) + '</b></span>' +
    '<span class="lgi">динамика ' + dTotal + '</span>' +
    '<span class="lgi muted">' + grainName + '</span></div>' +
    '<div class="trendplot"><div class="trendtip" hidden></div></div>' +
    (notes.length ? '<p class="trend-note">' + esc(notes.join(' ')) + '</p>' : '') + '</section>';

  box.querySelectorAll('[data-cmp]').forEach(function (b) {
    b.addEventListener('click', function () {
      trendCompare = b.dataset.cmp;
      try { localStorage.setItem('mingkh_trendcmp', trendCompare); } catch (e) { /* нет хранилища */ }
      renderTrend();
    });
  });

  var plot = box.querySelector('.trendplot');
  var W = Math.max(plot.clientWidth, 300);
  var slot = (W - 48) / Math.max(n, 1);
  var rich = slot >= 52;                  // места хватает на подписи под каждым столбиком
  // сверху место под две подписи друг над другом — на случай, если они сойдутся
  var PAD_L = 40, PAD_R = 8, PAD_T = rich ? 32 : 10, PAD_B = rich ? 44 : 26;
  var H = 250, plotW = W - PAD_L - PAD_R, plotH = H - PAD_T - PAD_B;
  slot = plotW / Math.max(n, 1);

  var max = 0;
  s.keys.forEach(function (k) { max = Math.max(max, s.cur[k] || 0, s.cmp[k] || 0); });
  var step = niceStep(max / 4), top = Math.max(step, Math.ceil(max / step) * step);
  function y(v) { return PAD_T + plotH - v / top * plotH; }

  var svg = '';
  for (var v = 0; v <= top + 1e-9; v += step) {           // сетка — едва заметная
    svg += '<line class="trend-grid" x1="' + PAD_L + '" x2="' + (W - PAD_R) + '" y1="' + y(v) +
      '" y2="' + y(v) + '"/><text class="axlab axis-y" x="' + (PAD_L - 6) + '" y="' + (y(v) + 4) +
      '" text-anchor="end">' + fmt(v) + '</text>';
  }

  var group = Math.min(slot * 0.74, 64), bw = (group - 2) / 2;
  var labelEvery = Math.max(1, Math.ceil(46 / slot));
  s.keys.forEach(function (k, i) {
    var c = s.cur[k] || 0, p = s.cmp[k] || 0;
    var x0 = PAD_L + i * slot + (slot - group) / 2;
    svg += '<path class="bar-cmp" d="' + barPath(x0, y(p), bw, y(0) - y(p)) + '"/>';
    svg += '<path class="bar-cur" d="' + barPath(x0 + bw + 2, y(c), bw, y(0) - y(c)) + '"/>';
    var cx = PAD_L + i * slot + slot / 2;
    if (rich) {
      var xc = x0 + bw + 2 + bw / 2, yc = y(c) - 5;
      var xp = x0 + bw / 2, yp = y(p) - 5;
      // ширину подписи прикидываем по символам: цифры шрифта ~6 px
      var wc = fmt(c).length * 6, wp = fmt(p).length * 6;
      var clashX = Math.abs(xc - xp) < (wc + wp) / 2 + 2;
      if (c && p && clashX && Math.abs(yc - yp) < 12) {
        yp = Math.min(yp, yc) - 12;       // подписи сошлись — «год назад» уходит выше
      }
      if (p) {
        svg += '<text class="val val-cmp" x="' + xp + '" y="' + yp +
          '" text-anchor="middle">' + fmt(p) + '</text>';
      }
      if (c) {
        svg += '<text class="val" x="' + xc + '" y="' + yc +
          '" text-anchor="middle">' + fmt(c) + '</text>';
      }
    }
    if (i % labelEvery === 0) {
      svg += '<text class="axlab" x="' + cx + '" y="' + (H - PAD_B + 15) +
        '" text-anchor="middle">' + esc(bucketLabel(k, s.grain, i, n)) + '</text>';
    }
    if (rich) {
      var ds = deltaSvg(c, p);
      svg += '<text class="dlt ' + ds.cls + '" x="' + cx + '" y="' + (H - PAD_B + 32) +
        '" text-anchor="middle">' + esc(ds.text) + '</text>';
    }
    // невидимая зона наведения на всю высоту — попасть мышью легко
    svg += '<rect class="hit" x="' + (PAD_L + i * slot) + '" y="' + PAD_T + '" width="' + slot +
      '" height="' + plotH + '" data-i="' + i + '"><title>' +
      esc(bucketLabel(k, s.grain, 0, 99) + ': поступило ' + fmt(c) + ', ' + cmpName + ' ' + fmt(p)) +
      '</title></rect>';
  });
  svg += '<line class="base" x1="' + PAD_L + '" x2="' + (W - PAD_R) + '" y1="' + y(0) +
    '" y2="' + y(0) + '"/>';

  plot.insertAdjacentHTML('beforeend',
    '<svg viewBox="0 0 ' + W + ' ' + H + '" width="' + W + '" height="' + H +
    '" role="img" aria-label="Обращения ' + grainName + ': текущий период и ' + cmpName + '">' +
    svg + '</svg>');

  var tip = plot.querySelector('.trendtip');
  plot.querySelectorAll('.hit').forEach(function (h) {
    h.addEventListener('mouseenter', function () {
      var i = +h.dataset.i, k = s.keys[i], c = s.cur[k] || 0, p = s.cmp[k] || 0;
      tip.innerHTML = '<b>' + esc(bucketLabel(k, s.grain, 0, 99)) + '</b><br>' +
        'поступило: <b>' + fmt(c) + '</b><br>' + cmpName + ': ' + fmt(p) + '<br>' +
        'динамика: ' + deltaHtml(c, p);
      tip.hidden = false;
      var left = PAD_L + i * slot + slot / 2;
      tip.style.left = Math.min(Math.max(left - 80, 0), W - 170) + 'px';
      h.classList.add('on');
    });
    h.addEventListener('mouseleave', function () {
      tip.hidden = true;
      h.classList.remove('on');
    });
  });
}

// Ширина графика зависит от окна: при изменении размера перерисовываем только его
var trendResize = null;
window.addEventListener('resize', function () {
  clearTimeout(trendResize);
  trendResize = setTimeout(function () { if (data) renderTrend(); }, 150);
});
