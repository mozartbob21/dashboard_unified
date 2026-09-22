
// Сортировка таблиц по клику на заголовок.
// Ключ берётся из data-v, а при его отсутствии — из текста ячейки.
function bindSort(root) {
  (root || document).querySelectorAll('table.sortable').forEach(function (table) {
    if (table.dataset.sortBound) return;
    table.dataset.sortBound = '1';
    var heads = Array.prototype.slice.call(table.tHead.rows[0].cells);
    heads.forEach(function (th, col) {
      var type = th.dataset.sort;
      if (!type) return;
      th.tabIndex = 0;
      th.classList.add('sortable-h');
      function apply() {
        // первый клик по числовой колонке — по убыванию, по текстовой — по возрастанию
        var cur = th.getAttribute('aria-sort');
        var desc = cur ? cur !== 'descending' : type === 'num';
        heads.forEach(function (h) { h.removeAttribute('aria-sort'); });
        th.setAttribute('aria-sort', desc ? 'descending' : 'ascending');

        var body = table.tBodies[0];
        // раскрытые подстроки схлопываем: иначе сортировка их перемешает
        body.querySelectorAll('tr.detailrow').forEach(function (r) { r.remove(); });
        body.querySelectorAll('.caretbtn').forEach(function (b) {
          b.classList.remove('open');
          b.setAttribute('aria-expanded', 'false');
        });
        var rows = Array.prototype.slice.call(body.rows);
        rows.sort(function (a, b) {
          var x = a.cells[col], y = b.cells[col];
          if (type === 'num') {
            var xv = x.dataset.v, yv = y.dataset.v;
            // пустые значения всегда внизу, в любом направлении
            var xe = xv === undefined || xv === '';
            var ye = yv === undefined || yv === '';
            if (xe && ye) return 0;
            if (xe) return 1;
            if (ye) return -1;
            return desc ? yv - xv : xv - yv;
          }
          var xs = x.textContent.trim(), ys = y.textContent.trim();
          return desc ? ys.localeCompare(xs, 'ru') : xs.localeCompare(ys, 'ru');
        });
        rows.forEach(function (r) { body.appendChild(r); });
      }
      th.addEventListener('click', apply);
      th.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); apply(); }
      });
    });
  });
}

// Переключатель порядка плиток тепловой карты.
// Цвет плитки всегда означает коэффициент на 10 тыс. — меняется только порядок.
function bindMapSort(root) {
  (root || document).querySelectorAll('.maptools').forEach(function (tools) {
    var grid = tools.parentNode.querySelector('.grid');
    if (!grid) return;
    tools.querySelectorAll('.mapsort').forEach(function (btn) {
      btn.addEventListener('click', function () {
        tools.querySelectorAll('.mapsort').forEach(function (b) { b.classList.remove('on'); });
        btn.classList.add('on');
        var by = btn.dataset.by;
        var cells = Array.prototype.slice.call(grid.children);
        cells.sort(function (a, b) {
          if (by === 'title') {
            return a.dataset.title.localeCompare(b.dataset.title, 'ru');
          }
          return b.dataset[by] - a.dataset[by];   // числовые — по убыванию
        });
        cells.forEach(function (c) { grid.appendChild(c); });
        try { localStorage.setItem('ecur_mapsort', by); } catch (e) { /* не запоминаем */ }
      });
    });
    var saved = null;
    try { saved = localStorage.getItem('ecur_mapsort'); } catch (e) { /* нет хранилища */ }
    if (saved) {
      var b = tools.querySelector('.mapsort[data-by="' + saved + '"]');
      if (b) b.click();
    }
  });
}
