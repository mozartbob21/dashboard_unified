/* Дашборд обращений МИНЖКХ — вся фильтрация и отрисовка на стороне браузера.
 *
 * Сервер отдаёт детальные строки обращений за три периода (текущий, предыдущий,
 * АППГ) один раз на выбранный период. Дальше переключение любого фильтра —
 * пересчёт в памяти, без обращения к порталу.
 */
'use strict';

// порядок обязан совпадать с DIMS в dataset.py
var SERVER_DIMS = ['omsu', 'source', 'direction', 'theme', 'subtopic', 'fact', 'executor'];
// popgroup считается здесь же из населения — на портал за ним ходить не нужно
var DIMS = SERVER_DIMS.concat(['popgroup']);
var DIM_TITLE = {
  omsu: 'Муниципалитет', source: 'Источник', direction: 'Направление',
  theme: 'Тема', subtopic: 'Подтема', fact: 'Факт', executor: 'Исполнитель',
  popgroup: 'Группа населения'
};
var DIM_ALL = {
  omsu: 'Все муниципалитеты', source: 'Все источники', direction: 'Все направления',
  theme: 'Все темы', subtopic: 'Все подтемы', fact: 'Все факты',
  executor: 'Все исполнители', popgroup: 'Любая численность'
};
// фильтры на панели; муниципалитет выбирается и отсюда, и кликом по карте
var FILTER_DIMS = ['omsu', 'source', 'direction', 'theme', 'subtopic', 'fact',
                   'executor', 'popgroup'];

// Значения, которые не показываем в разрезах: к работе МИНЖКХ они не относятся.
// В общий итог обращения при этом входят — цифра «Обращений» остаётся полной.
var HIDDEN = {
  direction: ['Вне компетенции Ведомств МО']
};

function isHidden(dim, code) {
  var list = HIDDEN[dim];
  return !!list && list.indexOf(rawLabel(dim, code)) !== -1;
}

// Группы по численности — как на портале: до 100 тыс., 100-200 тыс., свыше 200 тыс.
var POP_GROUPS = ['до 100 тыс.', '100-200 тыс.', 'больше 200 тыс.', 'без данных'];

function popGroupCode(population) {
  if (!population) return 3;
  if (population < 100000) return 0;
  if (population <= 200000) return 1;
  return 2;
}

/* Быстрые наборы фильтров. Значения задаются точным названием из данных;
   если названия на портале поменяются, кнопка просто не найдёт совпадений —
   это видно сразу, в отличие от молчаливого неверного подсчёта. */
var QUICK = [
  {
    key: 'water', title: 'Вода', cls: 'q-blue',
    hint: 'Направление «Водоснабжение/Водоотведение» — это темы «Качество воды» и «Водоотведение»',
    set: { direction: ['Водоснабжение/Водоотведение'] }
  },
  {
    key: 'kapremont', title: 'Капитальный ремонт', cls: 'q-green',
    hint: 'Подтема «Фонд капитального ремонта»',
    set: { subtopic: ['Фонд капитального ремонта'] }
  },
  {
    key: 'fkr', title: 'ФКР', cls: 'q-orange',
    hint: 'Исполнитель — Фонд капитального ремонта общего имущества многоквартирных домов',
    set: { executor: ['Фонд капитального ремонта общего имущества многоквартирных домов'] }
  }
];

var SEQ_STEPS = 7;                 // столько ступеней в красной шкале карты
var THEME_STEPS = 4;               // цветов для тем

var data = null;                   // текущий датасет
var sel = {};                      // выбранные значения: dim -> Set индексов
var mapSort = 'deltaAbs';          // по умолчанию — динамика в штуках
var notMunicipal = {};

/* ---------- утилиты ---------- */

function $(id) { return document.getElementById(id); }

function esc(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

function fmt(v, digits) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  digits = digits || 0;
  var s = Math.abs(v).toFixed(digits);
  var parts = s.split('.');
  var whole = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return (v < 0 ? '-' : '') + whole + (parts[1] ? ',' + parts[1] : '');
}

function plural(n, one, few, many) {
  var t = n % 10, h = n % 100;
  if (t === 1 && h !== 11) return one;
  if (t >= 2 && t <= 4 && (h < 12 || h > 14)) return few;
  return many;
}

/** Процент без дробной части. Ненулевое изменение не должно превращаться
 *  в «0%», поэтому всё, что меньше процента, показываем как «<1%». */
function pctText(v) {
  var a = Math.abs(v);
  if (a > 0 && Math.round(a) === 0) return '<1';
  return fmt(a, 0);
}

function deltaHtml(cur, prev, plain) {
  var flat = plain ? 'd-plain' : 'd-flat';
  if (cur === null || prev === null || !prev) {
    // на плитке места мало — коротко, полная формулировка есть в подсказке
    return '<span class="' + flat + '">' + (plain ? 'нет сравнения' : '—') + '</span>';
  }
  var d = cur - prev, pct = d / prev * 100;
  if (Math.abs(pct) < 0.05) return '<span class="' + flat + '">без изменений</span>';
  var arrow = d > 0 ? '▲' : '▼';
  if (plain) {
    return '<span class="d-plain">' + arrow + ' ' + fmt(Math.abs(d)) +
           ' (' + pctText(pct) + '%)</span>';
  }
  // рост числа обращений — плохо (красный), падение — хорошо (зелёный)
  return '<span class="' + (d > 0 ? 'd-bad' : 'd-good') + '">' + arrow + ' ' +
         fmt(Math.abs(d)) + ' (' + pctText(pct) + '%)</span>';
}

function numcell(value, text) {
  var v = (value === null || value === undefined || isNaN(value)) ? '' : value;
  return '<td class="numcell" data-v="' + v + '">' + text + '</td>';
}

/* ---------- фильтрация и агрегация ---------- */

/** Подходит ли строка под фильтр. skipDim — измерение, которое не учитываем
 *  (нужно, чтобы в разрезе по источникам сам фильтр источника не схлопывал
 *  диаграмму в один сектор). */
function matches(row, skipDim) {
  for (var i = 0; i < DIMS.length; i++) {
    var d = DIMS[i];
    if (d === skipDim) continue;
    var s = sel[d];
    if (s && s.size && !s.has(row[i])) return false;
  }
  return true;
}

/** Сколько строк проходит фильтр. */
function total(period, skipDim) {
  var rows = data.rows[period] || [], n = 0;
  for (var i = 0; i < rows.length; i++) if (matches(rows[i], skipDim)) n++;
  return n;
}

/** Разбивка по измерению: индекс значения -> количество. */
function groupBy(period, dim, skipSelf) {
  var col = DIMS.indexOf(dim), rows = data.rows[period] || [], out = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!matches(r, skipSelf ? dim : null)) continue;
    out[r[col]] = (out[r[col]] || 0) + 1;
  }
  return out;
}

/** Название значения как оно пришло с портала. */
function rawLabel(dim, code) { return data.dims[dim][code]; }

// Переименования для показа. Слева — как на портале, справа — как надо нам.
// Сопоставление с данными и фильтрами идёт по исходным названиям.
var RENAME = {
  direction: {
    'Общественные территории': 'Содержание/обустройство объектов водоотведения',
    // «Двор» = тема МКД: 93% из них — Фонд капитального ремонта,
    // остальное (мусор в МКД, УК, платёжки) тоже осталось под этим названием
    'Двор': 'Капитальный ремонт'
  },
  theme: {
    'МКД': 'Капитальный ремонт'
  }
};

/** Название для показа — с учётом переименований. */
function label(dim, code) {
  var raw = data.dims[dim][code];
  var map = RENAME[dim];
  return (map && map[raw]) || raw;
}

/* ---------- блоки ---------- */

/** Охват и интенсивность по муниципалитетам за один период. */
function coverage(period) {
  var by = groupBy(period, 'omsu');
  var count = 0, shareSum = 0, shareN = 0;
  Object.keys(by).forEach(function (code) {
    var name = rawLabel('omsu', code);
    if (notMunicipal[name]) return;
    var pop = data.population[name];
    count++;
    if (pop) { shareSum += by[code] / pop * 10000; shareN++; }
  });
  return { count: count, avgShare: shareN ? shareSum / shareN : null };
}

function renderKpi() {
  var cur = total('curr'), prev = total('prev'), appg = total('appg');
  var cc = coverage('curr'), cp = coverage('prev'), ca = coverage('appg');

  var cards = [
    ['Обращений', fmt(cur),
     '<span class="kpi-sub">пред. ' + deltaHtml(cur, prev) +
     ' &middot; АППГ ' + deltaHtml(cur, appg) + '</span>'],
    ['Муниципалитетов', fmt(cc.count),
     '<span class="kpi-sub">пред. ' + deltaHtml(cc.count, cp.count) +
     ' &middot; АППГ ' + deltaHtml(cc.count, ca.count) + '</span>'],
    ['На 10 тыс. жителей', fmt(cc.avgShare, 1),
     '<span class="kpi-sub">пред. ' + deltaHtml(cc.avgShare, cp.avgShare) +
     ' &middot; АППГ ' + deltaHtml(cc.avgShare, ca.avgShare) + '</span>']
  ];
  $('kpi').innerHTML = '<div class="kpis">' + cards.map(function (c) {
    return '<div class="kpi"><div class="kpi-t">' + esc(c[0]) + '</div>' +
           '<div class="kpi-v">' + c[1] + '</div>' + c[2] + '</div>';
  }).join('') + '</div>';
}

function renderSources() {
  var g = groupBy('curr', 'source', true);
  var gp = groupBy('prev', 'source', true);
  var items = Object.keys(g).map(function (code) {
    return { code: +code, label: label('source', code), v: g[code], prev: gp[code] || 0 };
  }).filter(function (x) { return x.v > 0; });
  if (!items.length) { $('sources').innerHTML = ''; return; }
  items.sort(function (a, b) { return b.v - a.v; });
  var sum = items.reduce(function (a, x) { return a + x.v; }, 0) || 1;
  var chosen = sel.source;

  var rows = items.map(function (x) {
    // без цветных квадратиков: в этой таблице нет полос, цвет ничего не кодировал
    var on = chosen && chosen.size && chosen.has(x.code);
    var share = x.v / sum * 100;
    return '<tr class="clickrow' + (on ? ' on' : '') + '" data-dim="source" data-code="' +
      x.code + '" tabindex="0" role="button" aria-pressed="' + (on ? 'true' : 'false') +
      '" title="Нажмите, чтобы отфильтровать по источнику">' +
      '<th scope="row">' + esc(x.label) + '</th>' +
      numcell(x.v, fmt(x.v)) +
      numcell(share, pctText(share) + '%') +
      numcell(x.prev ? (x.v - x.prev) / x.prev * 100 : null, deltaHtml(x.v, x.prev)) +
      '</tr>';
  }).join('');

  $('sources').innerHTML =
    '<section class="card"><h2>Источники обращений</h2>' +
    '<table class="plain sortable"><thead><tr>' +
    '<th scope="col" data-sort="text">Источник</th>' +
    '<th scope="col" class="numcell" data-sort="num" aria-sort="descending">Обращений</th>' +
    '<th scope="col" class="numcell" data-sort="num">Доля</th>' +
    '<th scope="col" class="numcell" data-sort="num">К пред. периоду</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table></section>';
}


/** Общая таблица-разрез для направлений и тем. */
function breakdown(dim, containerId, title, colorPrefix, steps, expandDim) {
  var c = groupBy('curr', dim, true);
  var p = groupBy('prev', dim, true);
  var a = groupBy('appg', dim, true);
  var items = Object.keys(c).filter(function (code) {
    return !isHidden(dim, +code);
  }).map(function (code) {
    return {
      code: +code, label: label(dim, code),
      curr: c[code] || 0, prev: p[code] || 0, appg: a[code] || 0
    };
  });
  if (!items.length) { $(containerId).innerHTML = ''; return; }
  items.sort(function (x, y) { return y.curr - x.curr; });
  var mx = Math.max.apply(null, items.map(function (x) {
    return Math.max(x.curr, x.prev);
  })) || 1;
  var chosen = sel[dim];

  var rows = items.map(function (x, i) {
    var on = chosen && chosen.size && chosen.has(x.code);
    var step = i % steps;
    var swatch = colorPrefix ? '<span class="sw ' + colorPrefix + step + '"></span>' : '';
    var caret = '';
    if (expandDim) {
      var xid = 'x' + (++expandSeq);
      expandScopes[xid] = [[dim, x.code]];
      caret = '<button type="button" class="caretbtn" data-scope="' + xid +
        '" data-child="' + expandDim + '" aria-expanded="false" ' +
        'aria-label="Показать ' + esc(DIM_TITLE[expandDim]) + '" ' +
        'title="Показать ' + esc(DIM_TITLE[expandDim]) + '"><span class="caret"></span></button>';
    }
    var barCls = colorPrefix ? colorPrefix + step : 'b-curr';
    return '<tr class="clickrow' + (on ? ' on' : '') + '" data-dim="' + dim +
      '" data-code="' + x.code + '" tabindex="0" role="button" aria-pressed="' +
      (on ? 'true' : 'false') + '" title="Нажмите, чтобы отфильтровать">' +
      '<th scope="row">' + caret + swatch + esc(x.label) + '</th>' +
      '<td class="barcell">' +
      '<div class="bar ' + barCls + '" style="width:' + (x.curr / mx * 100).toFixed(2) + '%"></div>' +
      '<div class="bar b-prev" style="width:' + (x.prev / mx * 100).toFixed(2) + '%"></div></td>' +
      numcell(x.curr, fmt(x.curr)) +
      numcell(x.prev ? (x.curr - x.prev) / x.prev * 100 : null, deltaHtml(x.curr, x.prev)) +
      numcell(x.appg ? (x.curr - x.appg) / x.appg * 100 : null, deltaHtml(x.curr, x.appg)) +
      '</tr>';
  }).join('');

  $(containerId).innerHTML =
    '<section class="card"><h2>' + esc(title) + '</h2>' +
    '<table class="bartable sortable"><thead><tr>' +
    '<th scope="col" data-sort="text">' + esc(DIM_TITLE[dim]) + '</th>' +
    '<th scope="col">Сравнение</th>' +
    '<th scope="col" class="numcell" data-sort="num" aria-sort="descending">Текущий</th>' +
    '<th scope="col" class="numcell" data-sort="num">К пред. периоду</th>' +
    '<th scope="col" class="numcell" data-sort="num">К АППГ</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table></section>';

  if (expandDim) bindBreakdownExpand($(containerId));
}

/** Раскрытие строки разреза во вложенный разрез (тема -> факты).
 *  Каретка раскрывает, клик по остальной строке по-прежнему фильтрует. */
/* Иерархия раскрытия, как на портале: направление -> темы -> подтемы -> факты.
   Ключ — измерение строки, значение — что показать внутри неё. */
var CHILD_DIM = {
  direction: 'theme',
  theme: 'subtopic',
  subtopic: 'fact'
};

// Ограничения раскрытых уровней: id кнопки -> [[измерение, код], ...]
var expandScopes = {};
var expandSeq = 0;

/** Разрез childDim при заданных ограничениях. Фильтр по самому childDim
 *  игнорируем — иначе в раскрытии осталась бы одна строка. */
function scopedBreakdown(scope, childDim) {
  var ci = DIMS.indexOf(childDim);
  var idx = scope.map(function (s) { return [DIMS.indexOf(s[0]), s[1]]; });
  function count(period) {
    var rows = data.rows[period] || [], out = {};
    outer: for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      for (var j = 0; j < idx.length; j++) {
        if (r[idx[j][0]] !== idx[j][1]) continue outer;
      }
      if (!matches(r, childDim)) continue;
      out[r[ci]] = (out[r[ci]] || 0) + 1;
    }
    return out;
  }
  var c = count('curr'), p = count('prev');
  return Object.keys(c).filter(function (code) {
    return !isHidden(childDim, +code);
  }).map(function (code) {
    return { code: +code, label: label(childDim, code), curr: c[code], prev: p[code] || 0 };
  }).sort(function (a, b) { return b.curr - a.curr; });
}

/** Таблица одного уровня раскрытия. Строки фильтруют, каретка уводит глубже. */
function childTable(scope, childDim) {
  var list = scopedBreakdown(scope, childDim);
  if (!list.length) return '<p class="sub" style="margin:0">Нет данных.</p>';
  var deeper = CHILD_DIM[childDim];
  var chosen = sel[childDim];

  return '<table class="plain inner"><thead><tr>' +
    '<th scope="col">' + esc(DIM_TITLE[childDim]) + '</th>' +
    '<th scope="col" class="numcell">Текущий</th>' +
    '<th scope="col" class="numcell">Пред. период</th>' +
    '<th scope="col" class="numcell">Динамика</th></tr></thead><tbody>' +
    list.map(function (s) {
      var on = chosen && chosen.size && chosen.has(s.code);
      var caret = '';
      if (deeper) {
        var id = 'x' + (++expandSeq);
        expandScopes[id] = scope.concat([[childDim, s.code]]);
        caret = '<button type="button" class="caretbtn" data-scope="' + id +
          '" data-child="' + deeper + '" aria-expanded="false" ' +
          'aria-label="Показать ' + esc(DIM_TITLE[deeper]) + '" ' +
          'title="Показать ' + esc(DIM_TITLE[deeper]) + '"><span class="caret"></span></button>';
      }
      return '<tr class="clickrow' + (on ? ' on' : '') + '" data-dim="' + childDim +
        '" data-code="' + s.code + '" tabindex="0" role="button" aria-pressed="' +
        (on ? 'true' : 'false') + '" title="Нажмите, чтобы отфильтровать">' +
        '<th scope="row">' + caret + esc(s.label) + '</th>' +
        numcell(s.curr, fmt(s.curr)) + numcell(s.prev, fmt(s.prev)) +
        numcell(s.prev ? (s.curr - s.prev) / s.prev * 100 : null,
                deltaHtml(s.curr, s.prev)) + '</tr>';
    }).join('') + '</tbody></table>';
}

/** Навешивает раскрытие на все каретки внутри контейнера, на любом уровне. */
function bindBreakdownExpand(root) {
  (root || document).querySelectorAll('.caretbtn[data-scope]').forEach(function (btn) {
    if (btn.dataset.bound) return;
    btn.dataset.bound = '1';
    btn.addEventListener('click', function (e) {
      e.stopPropagation();               // не даём сработать фильтру строки
      var tr = btn.closest('tr');
      var next = tr.nextElementSibling;
      if (next && next.classList.contains('detailrow')) {
        next.remove();
        btn.setAttribute('aria-expanded', 'false');
        btn.classList.remove('open');
        return;
      }
      var row = document.createElement('tr');
      row.className = 'detailrow';
      row.innerHTML = '<td colspan="' + tr.cells.length + '">' +
        childTable(expandScopes[btn.dataset.scope], btn.dataset.child) + '</td>';
      tr.parentNode.insertBefore(row, tr.nextSibling);
      btn.setAttribute('aria-expanded', 'true');
      btn.classList.add('open');
      bindBreakdownExpand(row);          // вложенные уровни
      bindPicks(row);
      bindSort(row);
    });
  });
}

function renderHeatmap() {
  var c = groupBy('curr', 'omsu', true);
  var p = groupBy('prev', 'omsu', true);
  var items = Object.keys(c).map(function (code) {
    var name = rawLabel('omsu', code);
    var pop = data.population[name] || 0;
    return {
      code: +code, title: name, amount: c[code] || 0, prev: p[code] || 0,
      pop: pop, share: pop ? c[code] / pop * 10000 : 0
    };
  }).filter(function (x) { return !notMunicipal[x.title]; });
  if (!items.length) { $('heatmap').innerHTML = ''; return; }

  // Цвет считается по той же величине, по которой сортируем: иначе порядок
  // плиток спорил бы с их окраской.
  var byDelta = mapSort === 'delta' || mapSort === 'deltaAbs';
  var metric = byDelta ? mapSort : mapSort;

  items.forEach(function (x) {
    // в штуках изменение считается всегда, даже если в прошлом периоде был ноль;
    // в процентах при нулевой базе доля не определена
    x.deltaAbs = x.amount - x.prev;
    x.delta = x.prev ? (x.amount - x.prev) / x.prev * 100 : null;
  });

  var cls, legendCaption, legendSteps;
  if (byDelta) {
    // Два порога по модулю изменения: слабое / заметное / сильное.
    var bands;
    if (mapSort === 'delta') {
      // Проценты: пороги фиксированные, чтобы один и тот же цвет означал
      // одно и то же при любых фильтрах.
      bands = [20, 50];
      legendCaption = 'динамика в %: падение';
    } else {
      // Штуки: разброс зависит от длины периода, поэтому пороги берём
      // из самих данных — по квантилям модуля ненулевых изменений.
      var mags = items.map(function (x) { return x.deltaAbs; })
                      .filter(function (v) { return v; })
                      .map(Math.abs).sort(function (a, b) { return a - b; });
      function q(p) {
        if (!mags.length) return 1;
        return mags[Math.min(mags.length - 1, Math.floor(mags.length * p))];
      }
      var a1 = Math.max(q(0.50), 1), a2 = Math.max(q(0.80), a1 + 1);
      bands = [a1, a2];
      legendCaption = 'динамика в шт.: падение';
    }
    // Серый — только когда динамики нет совсем. Всё остальное красится.
    cls = function (x) {
      var v = x[metric];
      if (v === null || v === 0) return 'dv3';
      var m = Math.abs(v);
      var lvl = m >= bands[1] ? 2 : (m >= bands[0] ? 1 : 0);
      return v < 0 ? 'dv' + (2 - lvl) : 'dv' + (4 + lvl);
    };
    legendSteps = 'dv';
  } else {
    var vals = items.map(function (x) { return x[metric]; })
                    .sort(function (a, b) { return a - b; });
    var cuts = [];
    for (var i = 0; i < SEQ_STEPS; i++) {
      cuts.push(vals[Math.min(vals.length - 1,
                Math.ceil(vals.length * (i + 1) / SEQ_STEPS) - 1)]);
    }
    cls = function (x) {
      for (var i = 0; i < cuts.length; i++) if (x[metric] <= cuts[i]) return 's' + i;
      return 's' + (SEQ_STEPS - 1);
    };
    legendCaption = metric === 'amount' ? 'обращений: меньше' : 'на 10 тыс. жителей: меньше';
    legendSteps = 's';
  }
  items.sort(function (a, b) {
    if (byDelta) {
      // без базы сравнения — в конец, остальные по убыванию роста
      if (a[metric] === null && b[metric] === null) return 0;
      if (a[metric] === null) return 1;
      if (b[metric] === null) return -1;
    }
    return b[metric] - a[metric];
  });
  var chosen = sel.omsu;

  var cells = items.map(function (x) {
    var on = chosen && chosen.size && chosen.has(x.code);
    var tip = esc(x.title) + '\n' + fmt(x.amount) + ' ' +
      plural(x.amount, 'обращение', 'обращения', 'обращений') + '\n' +
      fmt(x.share, 1) + ' на 10 тыс. жителей\nнаселение ' + fmt(x.pop) +
      '\nпрошлый период: ' + fmt(x.prev);
    return '<div class="cell ' + cls(x) + (on ? ' picked' : '') +
      '" tabindex="0" role="button" aria-pressed="' + (on ? 'true' : 'false') +
      '" data-dim="omsu" data-code="' + x.code + '" data-tip="' + tip +
      '" data-amount="' + x.amount + '" data-share="' + x.share +
      '" data-title="' + esc(x.title) + '">' +
      '<div class="cell-t">' + esc(x.title) + '</div>' +
      '<div class="cell-v">' + fmt(x.amount) + '</div>' +
      '<div class="cell-d">' + deltaHtml(x.amount, x.prev, true) + '</div>' +
      '<div class="cell-share">' + fmt(x.share, 1) + ' на 10 тыс.</div></div>';
  }).join('');

  var legend = '';
  for (var k = 0; k < SEQ_STEPS; k++) legend += '<span class="lg ' + legendSteps + k + '"></span>';

  var rows = items.map(function (x) {
    var xid = 'x' + (++expandSeq);
    expandScopes[xid] = [['omsu', x.code]];
    var caret = '<button type="button" class="caretbtn" data-scope="' + xid +
      '" data-child="theme" aria-expanded="false" aria-label="Показать темы" ' +
      'title="Показать темы"><span class="caret"></span></button>';
    return '<tr>' +
      '<th scope="row">' + caret + esc(x.title) + '</th>' +
      numcell(x.amount, fmt(x.amount)) + numcell(x.prev, fmt(x.prev)) +
      numcell(x.prev ? (x.amount - x.prev) / x.prev * 100 : null,
              deltaHtml(x.amount, x.prev)) +
      numcell(x.share, fmt(x.share, 1)) + numcell(x.pop, fmt(x.pop)) + '</tr>';
  }).join('');

  $('heatmap').innerHTML =
    '<section class="card"><h2>Тепловая карта муниципалитетов</h2>' +
    '<div class="maptools"><span class="lg-t">Сортировать:</span>' +
    '<button type="button" class="pre mapsort' + (mapSort === 'deltaAbs' ? ' on' : '') +
    '" data-by="deltaAbs">динамика, шт.</button>' +
    '<button type="button" class="pre mapsort' + (mapSort === 'delta' ? ' on' : '') +
    '" data-by="delta">динамика, %</button>' +
    '<button type="button" class="pre mapsort' + (mapSort === 'share' ? ' on' : '') +
    '" data-by="share">по коэффициенту</button>' +
    '<button type="button" class="pre mapsort' + (mapSort === 'amount' ? ' on' : '') +
    '" data-by="amount">по количеству</button>' +
    '<span class="legend"><span class="lg-t">' + legendCaption + '</span>' +
    legend + '<span class="lg-t">' + (byDelta ? 'рост' : 'больше') +
    '</span></span></div>' +
    '<div class="grid">' + cells + '</div>' +
    '<details open><summary>Показать таблицей (' + items.length + ')</summary>' +
    '<table class="plain sortable"><thead><tr><th scope="col" data-sort="text">Муниципалитет</th>' +
    '<th scope="col" class="numcell" data-sort="num">Обращений</th>' +
    '<th scope="col" class="numcell" data-sort="num">Пред. период</th>' +
    '<th scope="col" class="numcell" data-sort="num">Динамика</th>' +
    '<th scope="col" class="numcell" data-sort="num" aria-sort="descending">На 10 тыс.</th>' +
    '<th scope="col" class="numcell" data-sort="num">Население</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table></details></section>';

  bindMapSortButtons();
  bindBreakdownExpand($('heatmap'));
}

/* ---------- фильтры на панели ---------- */

function buildFilters() {
  var box = $('filters');
  box.innerHTML = FILTER_DIMS.map(function (dim) {
    var g = groupBy('curr', dim, true);
    var codes = Object.keys(g).map(Number).filter(function (code) {
      return !isHidden(dim, code);
    });
    if (dim === 'omsu') {
      // муниципалитетов много — по алфавиту их искать проще, чем по объёму
      codes.sort(function (a, b) {
        return label(dim, a).localeCompare(label(dim, b), 'ru');
      });
    } else {
      codes.sort(function (a, b) { return g[b] - g[a]; });
    }
    var opts = '<label class="opt"><input type="checkbox" data-dim="' + dim +
      '" value="" checked><span>' + esc(DIM_ALL[dim]) + '</span></label>' +
      codes.map(function (code) {
        return '<label class="opt"><input type="checkbox" data-dim="' + dim +
          '" value="' + code + '"><span>' + esc(label(dim, code)) +
          ' <span class="lgv">' + fmt(g[code]) + '</span></span></label>';
      }).join('');
    // подпись не выводим: на кнопке и так написано «Все источники» и т.п.
    return '<div class="fld ms">' +
      '<button type="button" class="msbtn" id="btn_' + dim + '" aria-expanded="false" ' +
      'aria-label="' + esc(DIM_TITLE[dim]) + '" title="' + esc(DIM_TITLE[dim]) + '">' +
      esc(DIM_ALL[dim]) + '</button>' +
      '<div class="mspanel" id="pan_' + dim + '" hidden>' + opts + '</div></div>';
  }).join('');

  FILTER_DIMS.forEach(function (dim) {
    $('btn_' + dim).addEventListener('click', function (e) {
      e.stopPropagation();
      var pan = $('pan_' + dim), open = !pan.hidden;
      document.querySelectorAll('.mspanel').forEach(function (x) { x.hidden = true; });
      pan.hidden = open;
      $('btn_' + dim).setAttribute('aria-expanded', String(!open));
    });
    $('pan_' + dim).addEventListener('click', function (e) { e.stopPropagation(); });
    $('pan_' + dim).querySelectorAll('input').forEach(function (cb) {
      cb.addEventListener('change', function () {
        if (cb.value === '') {
          $('pan_' + dim).querySelectorAll('input').forEach(function (x) {
            if (x !== cb) x.checked = false;
          });
          sel[dim] = new Set();
        } else {
          $('pan_' + dim).querySelector('input[value=""]').checked = false;
          var s = sel[dim] || new Set();
          if (cb.checked) s.add(+cb.value); else s.delete(+cb.value);
          sel[dim] = s;
        }
        normalize(dim);
        rerender();
      });
    });
  });
  syncFilterLabels();
}

/** Быстрые наборы: кнопки «Вода», «Капитальный ремонт», «ФКР».
 *  Значения ищем по названию в справочнике текущего датасета. */
function applyQuick(preset) {
  var missing = [];
  DIMS.forEach(function (d) { sel[d] = new Set(); });
  Object.keys(preset.set).forEach(function (dim) {
    preset.set[dim].forEach(function (name) {
      var code = data.dims[dim].indexOf(name);
      if (code === -1) missing.push(name);
      else sel[dim].add(code);
    });
  });
  DIMS.forEach(function (d) { normalize(d); });
  syncChecks();
  rerender();
  return missing;
}

function quickActive(preset) {
  // набор считается включённым, если выбрано ровно то, что он задаёт
  var dims = Object.keys(preset.set);
  for (var i = 0; i < DIMS.length; i++) {
    var d = DIMS[i], s = sel[d] || new Set();
    if (dims.indexOf(d) === -1) { if (s.size) return false; continue; }
    var want = preset.set[d].map(function (n) { return data.dims[d].indexOf(n); })
                            .filter(function (c) { return c !== -1; });
    if (s.size !== want.length) return false;
    for (var j = 0; j < want.length; j++) if (!s.has(want[j])) return false;
  }
  return true;
}

function renderQuick() {
  var box = $('quick');
  if (!box) return;
  var any = DIMS.some(function (d) { return sel[d] && sel[d].size; });
  box.innerHTML = '<span class="glab">Быстрый выбор</span>' +
    QUICK.map(function (q) {
      return '<button type="button" class="pre quick ' + q.cls +
        (quickActive(q) ? ' on' : '') +
        '" data-quick="' + q.key + '" title="' + esc(q.hint) + '">' + esc(q.title) + '</button>';
    }).join('') +
    (any ? '<button type="button" class="pre" id="quickclear">Сбросить фильтры</button>' : '');

  box.querySelectorAll('[data-quick]').forEach(function (b) {
    b.addEventListener('click', function () {
      var q = QUICK.filter(function (x) { return x.key === b.dataset.quick; })[0];
      if (quickActive(q)) {                 // повторное нажатие снимает набор
        DIMS.forEach(function (d) { sel[d] = new Set(); normalize(d); });
        syncChecks();
        rerender();
        return;
      }
      var missing = applyQuick(q);
      if (missing.length) {
        $('status').className = 'err';
        $('status').textContent = 'В данных за этот период нет: ' + missing.join(', ') +
          '. Возможно, на портале изменились названия.';
      }
    });
  });
  var clr = $('quickclear');
  if (clr) {
    clr.addEventListener('click', function () {
      DIMS.forEach(function (d) { sel[d] = new Set(); normalize(d); });
      syncChecks();
      rerender();
    });
  }
}

/** Привести галочки на панели в соответствие с sel. */
function syncChecks() {
  FILTER_DIMS.forEach(function (dim) {
    var pan = $('pan_' + dim);
    if (!pan) return;
    var s = sel[dim] || new Set();
    pan.querySelectorAll('input').forEach(function (cb) {
      cb.checked = cb.value === '' ? s.size === 0 : s.has(+cb.value);
    });
  });
}

/** Пустой выбор равнозначен «все» — возвращаем галочку на «Все …». */
function normalize(dim) {
  var pan = $('pan_' + dim);
  if (!pan) return;
  if (!sel[dim] || !sel[dim].size) {
    sel[dim] = new Set();
    pan.querySelector('input[value=""]').checked = true;
  }
}

function syncFilterLabels() {
  FILTER_DIMS.forEach(function (dim) {
    var btn = $('btn_' + dim), s = sel[dim];
    if (!btn) return;
    if (!s || !s.size) {
      btn.textContent = DIM_ALL[dim];
      btn.classList.remove('active');
    } else if (s.size === 1) {
      btn.textContent = label(dim, s.values().next().value);
      btn.classList.add('active');
    } else {
      btn.textContent = 'Выбрано: ' + s.size;
      btn.classList.add('active');
    }
  });
  // выбранные муниципалитеты — отдельной строкой, их выбирают кликом по карте
  var s = sel.omsu, box = $('omsupick');
  if (!s || !s.size) { box.innerHTML = ''; return; }
  box.innerHTML = '<span class="lg-t">Муниципалитеты:</span> ' +
    Array.from(s).map(function (code) {
      return '<button type="button" class="chip" data-dim="omsu" data-code="' + code +
        '">' + esc(label('omsu', code)) + ' ×</button>';
    }).join('') +
    ' <button type="button" class="chip clear" id="omsuclear">сбросить</button>';
  $('omsuclear').addEventListener('click', function () {
    sel.omsu = new Set(); rerender();
  });
}

/* ---------- взаимодействие ---------- */

function toggle(dim, code) {
  var s = sel[dim] || new Set();
  if (s.has(code)) s.delete(code); else s.add(code);
  sel[dim] = s;
  normalize(dim);
  syncChecks();
  rerender();
}

function bindPicks(root) {
  (root || document).querySelectorAll('[data-dim][data-code]').forEach(function (el) {
    if (el.dataset.bound) return;
    el.dataset.bound = '1';
    function go(e) { e.preventDefault(); toggle(el.dataset.dim, +el.dataset.code); }
    el.addEventListener('click', go);
    el.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') go(e);
    });
  });
}

/** Кнопки порядка плиток. Смена порядка меняет и шкалу цвета,
 *  поэтому карта перерисовывается целиком, а не переставляется. */
function bindMapSortButtons() {
  document.querySelectorAll('#heatmap .mapsort').forEach(function (b) {
    b.addEventListener('click', function () {
      mapSort = b.dataset.by;
      try { localStorage.setItem('mingkh_mapsort', mapSort); } catch (e) { /* нет хранилища */ }
      renderHeatmap();
      bindPicks();
      bindSort();
    });
  });
}

/** Полная перерисовка из уже загруженных данных — без сети. */
function rerender() {
  if (!data) return;
  var t0 = performance.now();
  expandScopes = {};           // старые раскрытия ссылаются на снесённые узлы
  renderKpi();
  renderSources();
  breakdown('direction', 'direction', 'Направления: период к периоду',
            null, 1, CHILD_DIM.direction);
  breakdown('theme', 'themes', 'Темы', 't', THEME_STEPS, CHILD_DIM.theme);
  renderHeatmap();
  syncFilterLabels();
  renderQuick();
  bindPicks();
  bindSort();
  $('calc').textContent = 'пересчёт ' + Math.round(performance.now() - t0) + ' мс';
}

window.DASH = {
  setData: function (d) {
    data = d;
    notMunicipal = {};
    (d.not_municipal || []).forEach(function (n) { notMunicipal[n] = true; });

    // дописываем каждой строке группу населения её муниципалитета
    d.dims.popgroup = POP_GROUPS;
    var byOmsuCode = d.dims.omsu.map(function (name) {
      return popGroupCode(d.population[name]);   // ключ — исходное название
    });
    Object.keys(d.rows).forEach(function (period) {
      d.rows[period].forEach(function (r) {
        if (r.length === SERVER_DIMS.length) r.push(byOmsuCode[r[0]]);
      });
    });
    DIMS.forEach(function (dim) { if (!sel[dim]) sel[dim] = new Set(); });
    // выбранные значения могли исчезнуть: коды привязаны к конкретному датасету
    DIMS.forEach(function (dim) { sel[dim] = new Set(); });
    buildFilters();
    try {
      var s = localStorage.getItem('mingkh_mapsort');
      // 'title' остался у тех, кто открывал прошлую версию — больше не поддерживаем
      if (['share', 'amount', 'delta', 'deltaAbs'].indexOf(s) !== -1) mapSort = s;
    } catch (e) { /* нет хранилища */ }
    rerender();
  },
  rerender: rerender
};

document.addEventListener('click', function () {
  document.querySelectorAll('.mspanel').forEach(function (x) { x.hidden = true; });
});
