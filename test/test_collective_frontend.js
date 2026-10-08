// Synthetic browser harness for state, exports and races; no browser or network required.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const tests = require('node:test');

function harness(realExcel = false) {
  const elements = new Map(), values = new Map(), pending = [], downloads = [];
  class Element {
    constructor() { this.value = ''; this.dataset = {}; this.hidden = true; this.children = new Map(); this.classList = { toggle() {} }; this.listeners = {}; }
    querySelector(key) { if (!this.children.has(key)) this.children.set(key, new Element()); return this.children.get(key); }
    querySelectorAll() { return []; }
    addEventListener(key, callback) { this.listeners[key] = callback; }
    setAttribute() {} focus() {} click() {} remove() {} appendChild() {} showModal() {} close() {}
  }
  const document = { querySelectorAll: () => [], addEventListener() {}, createElement: () => new Element(), body: new Element(),
    querySelector(key) { if (!elements.has(key)) elements.set(key, new Element()); return elements.get(key); },
    getElementById(id) { return this.querySelector('#' + id); },
  };
  document.querySelector('.collective-dashboard').dataset.owner = 'alice';
  document.getElementById('periodSel').value = 'd7';
  let XLSX = { utils: { book_new: () => [], aoa_to_sheet: data => ({ data }), book_append_sheet(wb, sheet, name) { wb.push({ sheet, name }); } }, writeFile: wb => downloads.push(wb) };
  if (realExcel) {
    const library = require('../static/vendor/xlsx.full.min.js');
    XLSX = {...library, writeFile(wb) {
      const bytes = library.write(wb, {type: 'buffer', bookType: 'xlsx'});
      downloads.push(library.read(bytes, {type: 'buffer'}));
    }};
  }
  const context = { document, localStorage: { getItem: k => values.get(k) ?? null, setItem: (k, v) => values.set(k, v), removeItem: k => values.delete(k) },
    fetch: (...args) => new Promise((resolve, reject) => pending.push({ args, resolve, reject })),
    AbortController, setTimeout: () => {}, Blob, URL: { createObjectURL: () => 'blob:synthetic', revokeObjectURL() {} },
    window: { XLSX }, XLSX, alert: () => {}, console,
  };
  vm.createContext(context);
  let source = fs.readFileSync(path.join(__dirname, '../static/collective.js'), 'utf8');
  source = source.replace(/\}\)\(\);\s*$/, 'globalThis.testApi = { classify, keyOf, store, load, state, baseFiltered, exportAll, exportPptx, refresh, getRows: () => ROWS, getPeriod: () => loadedPeriod };\n})();');
  vm.runInContext(source, context);
  return { api: context.testApi, document, pending, values, downloads };
}
const head = ['ОМСУ', 'Аннотация/краткое содержание', 'Подписи (фактическое число подписей)', 'Дата поступления обращения', 'Повтор', 'Статус', 'Номер обращения ЕЦУР, МСЭД'];
const rows = [['Первый', 'Нет воды', '101', '01.01.2026', 'Да', 'В работе', 'a'], ['Второй', 'Капитальный ремонт кровли', '100', '01.01.2026', '', 'В работе', 'b'], ['Первый', 'Освещение двора', '90', '01.01.2026', '', 'Не учитывается', 'c']];
const reply = (data, ok = true) => ({ ok, json: async () => data });
const flush = () => new Promise(resolve => setImmediate(resolve));

 tests('theme rules, manual corrections and user-local storage', () => {
  const h = harness(), a = h.api;
  assert.equal(a.classify({ 'Аннотация/краткое содержание': 'Капитальный ремонт водопровода' }), 'kr');
  assert.equal(a.classify({ 'Аннотация/краткое содержание': 'Нет газа' }), 'ii');
  assert.equal(a.classify({ 'Аннотация/краткое содержание': 'Покос газона' }), 'pr');
  const key = a.keyOf({ 'Номер обращения ЕЦУР, МСЭД': 'a' });
  h.values.set('neurona.collective.bob.' + key, 'pr');
  a.store.set(key, '__proto__');
  a.load(head, rows, 'Synthetic');
  assert.equal(a.getRows()[0]._theme, 'ii');
  a.store.set(key, 'pr'); a.load(head, rows, 'Synthetic');
  assert.equal(a.getRows()[0]._theme, 'pr');
  assert.notEqual(a.keyOf({ 'Аннотация/краткое содержание': 'Без номера А' }), a.keyOf({ 'Аннотация/краткое содержание': 'Без номера Б' }));
});

 tests('filters and Excel include the same rows and manual themes', () => {
  const h = harness(), a = h.api;
  a.load(head, rows, 'Synthetic');
  a.state.big = true;
  assert.equal(a.baseFiltered().length, 1); // exactly 100 is excluded
  a.state.big = false; a.state.hideNc = true;
  assert.equal(a.baseFiltered().length, 2);
  a.state.rep = true;
  assert.equal(a.baseFiltered().length, 1);
  a.getRows()[0]._theme = 'pr'; a.state.theme = 'pr';
  a.exportAll();
  const output = h.downloads[0];
  assert.deepEqual(Array.from(output, x => x.name), ['Обращения', 'По ОМСУ']);
  assert.equal(output[0].sheet.data.length, 2);
  assert.equal(output[0].sheet.data[1][0], 'Прочее');
  assert.equal(output[1].sheet.data[1][5], 101);
});

 tests('failed refresh and request races keep the actual loaded period for PPTX', async () => {
  const h = harness(), a = h.api;
  h.pending[0].resolve(reply({ head, rows, start: '2026-01-01', end: '2026-01-02', meta: 'Synthetic', updated: 'now' }));
  await flush();
  const refresh = a.refresh();
  h.pending[1].resolve(reply({ error: 'Synthetic failure' }, false)); await refresh;
  assert.deepEqual(Array.from(a.getPeriod()), ['2026-01-01', '2026-01-02']);
  assert.match(h.document.getElementById('err').textContent, /ранее загруженные/);
  const first = a.refresh(), second = a.refresh();
  h.pending[3].resolve(reply({ head, rows: rows.slice(0, 1), start: '2026-02-01', end: '2026-02-02', meta: 'New', updated: 'now' })); await second;
  h.pending[2].resolve(reply({ head, rows, start: '2026-01-01', end: '2026-01-02', meta: 'Stale', updated: 'before' })); await first;
  assert.deepEqual(Array.from(a.getPeriod()), ['2026-02-01', '2026-02-02']);
  assert.equal(a.getRows().length, 1);
  a.getRows()[0]._theme = 'kr';
  const exporting = a.exportPptx();
  const request = h.pending[4];
  assert.equal(request.args[0], '/mingkh/collective/api/pptx');
  const body = JSON.parse(request.args[1].body);
  assert.equal(body.start, '2026-02-01');
  assert.equal(body.end, '2026-02-02');
  assert.equal(body.rows[0]._theme, 'kr');
  request.resolve({ ok: true, headers: { get: () => '' }, blob: async () => new Blob() });
  await exporting;
});

tests('real XLSX roundtrip preserves selected data and treats portal formulas as text', () => {
  const h = harness(true), a = h.api;
  const formulaText = '=HYPERLINK("https://example.invalid", "portal text")';
  const source = rows.map(row => row.slice());
  source[0][1] = formulaText;
  a.load(head, source, 'Synthetic');
  a.state.big = true;
  a.getRows()[0]._theme = 'pr';
  a.exportAll();
  const workbook = h.downloads[0];
  assert.equal(workbook.SheetNames.length, 5);
  const sheet = workbook.Sheets['Обращения'];
  assert.equal(sheet['!ref'], 'A1:H2');
  assert.equal(sheet.A2.v, 'Прочее');
  assert.equal(sheet.C2.v, formulaText);
  assert.equal(sheet.C2.t, 's');
  assert.equal(sheet.C2.f, undefined);
  assert.equal(sheet.D2.v, 101);
  assert.equal(workbook.Sheets['По ОМСУ'].F2.v, 101);
});


tests('repeat whitespace is normalized consistently in filter, KPI, badges and export', () => {
  const h = harness(), a = h.api;
  const source = rows.map(row => row.slice());
  source[0][4] = ' Да ';
  source[1][4] = '\tда\u00a0';
  source[2][4] = 'Нет';
  a.load(head, source, 'Synthetic');
  assert.match(h.document.getElementById('kpis').innerHTML, /Повторных<\/div><div class="v">2<\/div>/);
  assert.equal((h.document.getElementById('list').innerHTML.match(/>Повтор<\/span>/g) || []).length, 2);
  a.state.rep = true;
  assert.equal(a.baseFiltered().length, 2);
  a.exportAll();
  assert.equal(h.downloads[0][0].sheet.data.length, 3);
});
