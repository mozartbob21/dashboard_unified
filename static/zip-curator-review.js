// Review state belongs to the registry/item, never to its current page number.
const zipReviewViews = new WeakMap();
const zipReviewEditors = new Map();
const zipReviewSaves = new Map();
const zipApproving = new Set();
let zipEditQueue = Promise.resolve();
const zipAttr = value => esc(value).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const zipInlineId = value => encodeURIComponent(value).replace(/'/g, '%27');
function zipReviewView(record) {
  if (!zipReviewViews.has(record)) zipReviewViews.set(record, {open: false, page: 0, size: 50});
  return zipReviewViews.get(record);
}
function zipRequired(value) {
  return typeof value === 'string' && value.trim() && !['—', '–', '-', 'нет', 'не указано', 'не указан', 'не определен', 'не определён', 'муниципалитет не определён', 'муниципалитет не определен'].includes(value.trim().toLowerCase());
}
function zipApprovalIssues(record) {
  const issues = [], add = (field, message, item_index) => issues.push({field, message, ...(item_index === undefined ? {} : {item_index})});
  if (!zipRequired(record.rso)) add('rso', 'Не указана организация РСО.');
  if (!zipRequired(record.okrug)) add('okrug', 'Укажите ОМСУ (муниципалитет) организации.');
  const items = Array.isArray(record.items) ? record.items : [];
  if (!items.length) add('items', 'В реестре нет позиций.');
  items.forEach((item, index) => {
    if (!zipRequired(item.name)) add('name', 'Не указано наименование позиции.', index);
    if (!zipRequired(item.cat) || !D.categories.includes(item.cat)) add('cat', 'Выберите категорию из справочника.', index);
    else {
      const groups = D.groupsByCat[item.cat] || [];
      if ((groups.length && !groups.includes(item.grp)) || (item.grp && !groups.includes(item.grp))) add('grp', 'Выберите группу из справочника.', index);
    }
    if (!item.unit || item.unitUnknown || normUnit(item.unit).unknown) add('unit', 'Укажите единицу измерения.', index);
    if (typeof item.qty !== 'number' || !Number.isFinite(item.qty) || item.qty < 0) add('qty', 'Укажите количество: число не меньше нуля.', index);
  });
  // Unknown server checks are retained; the known checks above stay current after editing OMSU.
  for (const issue of record.approval_issues || []) {
    if (issue && !['rso', 'okrug', 'municipality', 'items', 'name', 'cat', 'category', 'grp', 'group', 'unit', 'qty'].includes(issue.field)) issues.push(issue);
  }
  return issues;
}
function zipIssueSummary(issues) {
  const groups = new Map();
  for (const issue of issues) {
    const text = String(issue.message || 'Заполните обязательные данные.');
    if (!groups.has(text)) groups.set(text, {message: text, count: 0, first: issue.item_index});
    if (Number.isInteger(issue.item_index)) groups.get(text).count++;
  }
  return [...groups.values()];
}
function zipApprovalMarkup(record, pi) {
  const issues = zipApprovalIssues(record);
  if (!issues.length) return '<span class="review-ready">Все обязательные сведения заполнены.</span>';
  return '<strong>Согласование пока недоступно</strong><ul>' + zipIssueSummary(issues).map(issue =>
    `<li>${esc(issue.message)}${issue.count ? ` <span class="mini">Позиций: ${issue.count}.</span>` : ''}${Number.isInteger(issue.first) ? ` <button class="review-issue-link" onclick="showReviewIssue(${pi},${issue.first})">К позиции ${issue.first + 1}</button>` : ''}</li>`).join('') + '</ul>';
}
function updateApprovalStatus(pi) {
  const record = PENDING[pi]; if (!record) return;
  const issues = zipApprovalIssues(record), button = $('#approve' + pi), note = $('#approvalNotes' + pi);
  if (button) {
    button.disabled = issues.length > 0 || zipApproving.has(record) || [...zipReviewEditors.keys()].some(item => record.items.includes(item));
    button.textContent = zipApproving.has(record) ? 'Согласовываем…' : 'Согласовать';
  }
  if (note) {
    note.classList.toggle('is-ready', !issues.length);
    note.innerHTML = zipApprovalMarkup(record, pi);
  }
}
function zipEditorValues(editor) {
  if (!editor.row || !editor.row.isConnected) return editor.values;
  const values = {...editor.values};
  for (const field of ['cat', 'grp', 'unit', 'qty']) {
    const node = editor.row.querySelector('[data-edit="' + field + '"]');
    if (node) values[field] = node.value;
  }
  editor.values = values;
  return values;
}
function zipCaptureEditors() { for (const editor of zipReviewEditors.values()) zipEditorValues(editor); }
function zipUnitSelect(unit) {
  const units = [...new Set([...(D.units || []), 'пог. м'])];
  return `<select data-edit="unit" aria-label="Единица измерения"><option value="">Выберите единицу</option>${units.map(value => `<option value="${zipAttr(value)}" ${value === unit ? 'selected' : ''}>${esc(value === 'пог. м' ? 'пог. м — погонный метр' : value)}</option>`).join('')}</select>`;
}
function grpSel(pi, ii, cat, grp) {
  const groups = cat ? (D.groupsByCat[cat] || []) : [];
  return `<select data-edit="grp" aria-label="Группа позиции ${ii + 1}"><option value="">${groups.length ? 'Выберите группу' : 'Нет групп'}</option>${groups.map(value => `<option value="${zipAttr(value)}" ${value === grp ? 'selected' : ''}>${esc(value)}</option>`).join('')}</select>`;
}
function zipRenderEditor(pi, ii, row) {
  const item = PENDING[pi]?.items[ii], editor = zipReviewEditors.get(item); if (!editor) return;
  const values = editor.values;
  editor.row = row;
  row.classList.add('review-editing');
  row.querySelector('.jc').innerHTML = `<select data-edit="cat" aria-label="Категория позиции ${ii + 1}" onchange="onCat(${pi},${ii},this)"><option value="">Выберите категорию</option>${D.categories.map(value => `<option value="${zipAttr(value)}" ${value === values.cat ? 'selected' : ''}>${esc(value)}</option>`).join('')}</select>`;
  row.querySelector('.jg').innerHTML = grpSel(pi, ii, values.cat, values.grp);
  row.querySelector('.ju').innerHTML = zipUnitSelect(values.unit);
  row.querySelector('.jq').innerHTML = `<input type="text" data-edit="qty" inputmode="decimal" aria-label="Количество позиции ${ii + 1}" value="${zipAttr(values.qty ?? '')}" placeholder="0">`;
  row.querySelector('.ja').innerHTML = `<div class="review-row-actions"><button class="btn sm primary" onclick="saveReviewRow(${pi},${ii})">Сохранить</button><button class="btn sm ghost" onclick="cancelReviewRow(${pi},${ii})">Отмена</button><button class="btn sm ghost" onclick="reviewCategoryEditor(${pi},${ii})">+ Категория / группа</button></div><div class="review-save-note" role="status"></div>`;
  if (editor.error) row.querySelector('.review-save-note').textContent = editor.error;
}
function editRow(btn, pi, ii) {
  const record = PENDING[pi], item = record?.items[ii]; if (!item || zipApproving.has(record)) return;
  if (!zipReviewEditors.has(item)) zipReviewEditors.set(item, {record, values: {cat: item.cat || '', grp: item.grp || '', unit: item.unit || '', qty: item.qty ?? ''}, row: null, error: ''});
  zipRenderEditor(pi, ii, btn.closest('tr'));
  updateApprovalStatus(pi);
  zipReviewEditors.get(item).row.querySelector('[data-edit="unit"]').focus();
}
function onCat(pi, ii, select) {
  const editor = zipReviewEditors.get(PENDING[pi]?.items[ii]); if (!editor) return;
  zipEditorValues(editor); editor.values.cat = select.value; editor.values.grp = '';
  select.closest('tr').querySelector('.jg').innerHTML = grpSel(pi, ii, select.value, '');
}
function onGrp() {} // Group is saved with the entire row, so intermediate selections cannot race.
function persistEdit(pi, ii, values) {
  const next = zipEditQueue.then(() => zipPersistEdit(pi, ii, values));
  zipEditQueue = next.catch(() => false);
  return next;
}
async function zipPersistEdit(pi, ii, values) {
  const record = PENDING[pi], item = record?.items[ii]; if (!item) return false;
  const patch = values || {cat: item.cat, grp: item.grp, unit: item.unit, qty: item.qty};
  try {
    const response = await fetch('/zip_curator/api/edit', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({pi, ii, ...patch, revision: record.revision})});
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
    if (PENDING[pi] !== record || record.items[ii] !== item) throw new Error('Очередь изменилась. Обновите страницу, чтобы проверить сохранение.');
    const saved = data.state?.pending?.[pi];
    if (!saved?.items?.[ii]) throw new Error('Сервер не подтвердил сохранение позиции. Повторите попытку.');
    Object.assign(item, saved.items[ii]); record.revision = saved.revision; record.approval_issues = saved.approval_issues || [];
    if (item.cat) D.dict[norm(item.name)] = [item.cat, item.grp || '', item.water ? 1 : 0];
    return true;
  } catch (error) {
    const editor = zipReviewEditors.get(item);
    if (editor) editor.error = 'Правка не сохранена. ' + error.message;
    toast('Не удалось сохранить правку: ' + error.message, 6000);
    return false;
  }
}
async function saveReviewRow(pi, ii) {
  const record = PENDING[pi], item = record?.items[ii], editor = zipReviewEditors.get(item);
  if (!editor) return true;
  if (zipReviewSaves.has(item)) return zipReviewSaves.get(item);
  const values = {...zipEditorValues(editor)};
  const quantity = typeof values.qty === 'string' ? values.qty.trim().replace(',', '.') : values.qty;
  if (quantity !== '' && quantity !== null && (!Number.isFinite(Number(quantity)) || Number(quantity) < 0)) {
    editor.error = 'Введите количество: число не меньше нуля.';
    editor.row.querySelector('.review-save-note').textContent = editor.error;
    editor.row.querySelector('[data-edit="qty"]').focus();
    return false;
  }
  values.qty = quantity === '' || quantity === null ? null : Number(quantity);
  const controls = editor.row ? [...editor.row.querySelectorAll('button, input, select')] : [];
  controls.forEach(control => control.disabled = true);
  const note = editor.row?.querySelector('.review-save-note'); if (note) note.textContent = 'Сохраняем…';
  const saving = (async () => {
    const saved = await persistEdit(pi, ii, values);
    if (saved) { zipReviewEditors.delete(item); renderPendingPage(pi); updateApprovalStatus(pi); }
    else { controls.forEach(control => control.disabled = false); if (note) note.textContent = editor.error; }
    return saved;
  })();
  zipReviewSaves.set(item, saving);
  try { return await saving; } finally { zipReviewSaves.delete(item); }
}
function cancelReviewRow(pi, ii) {
  const item = PENDING[pi]?.items[ii]; if (zipReviewSaves.has(item)) return;
  zipReviewEditors.delete(item); renderPendingPage(pi); updateApprovalStatus(pi);
}
async function flushReviewEdits(record) {
  zipCaptureEditors();
  for (const [item, editor] of [...zipReviewEditors]) {
    if (record && editor.record !== record) continue;
    const pi = PENDING.indexOf(editor.record), ii = editor.record.items.indexOf(item);
    if (pi < 0 || ii < 0) { toast('Очередь изменилась. Проверьте сохранение правок перед продолжением.', 6000); return false; }
    if (!await saveReviewRow(pi, ii)) return false;
  }
  return true;
}
async function reviewCategoryEditor(pi, ii) {
  if (!await flushReviewEdits(PENDING[pi])) return;
  openCategoryEditor(pi, ii);
}
async function approve(idx) {
  const record = PENDING[idx]; if (!record || zipApproving.has(record)) return;
  if (!await flushReviewEdits()) return;
  if (zipApproving.has(record) || PENDING[idx] !== record) return;
  const issues = zipApprovalIssues(record);
  if (issues.length) { updateApprovalStatus(idx); toast('Сначала заполните обязательные сведения. Причины указаны в карточке.', 5000); return; }
  zipApproving.add(record); updateApprovalStatus(idx);
  try {
    const response = await fetch('/zip_curator/api/approve', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({idx, revision: record.revision})});
    const data = await response.json();
    if (!response.ok || !data.ok) {
      if (Array.isArray(data.issues)) record.approval_issues = data.issues;
      const note = $('#approvalNotes' + idx);
      if (note && Array.isArray(data.issues)) { note.classList.remove('is-ready'); note.textContent = 'Согласование недоступно: ' + data.issues.map(issue => issue.message).join(' '); }
      throw new Error(data.error || ('HTTP ' + response.status));
    }
    applyServerState(data.state || {});
    try { const channel = new BroadcastChannel('zip-curator-updates'); channel.postMessage({type: 'published'}); channel.close(); } catch (_) {}
    toast('Согласовано и опубликовано в ЗИП: ' + record.rso, 4000);
  } catch (error) { toast('Не удалось согласовать: ' + error.message, 6000); }
  finally { zipApproving.delete(record); const button = $('#approve' + idx); if (button && PENDING[idx] === record) { button.disabled = zipApprovalIssues(record).length > 0; button.textContent = 'Согласовать'; } }
}
function zipReviewPager(pi, record, view, count) {
  const pages = Math.max(1, Math.ceil(count / view.size)), start = count ? view.page * view.size + 1 : 0, end = Math.min(count, (view.page + 1) * view.size);
  return `<nav class="review-pagination" aria-label="Страницы позиций ${zipAttr(record.rso)}"><span class="review-range" role="status">Позиции ${start}–${end} из ${count}</span><div class="review-page-controls"><label>На странице <select aria-label="Количество позиций на странице" onchange="changeReviewPage(${pi},0,Number(this.value))">${[50,100,200].map(size => `<option ${size === view.size ? 'selected' : ''}>${size}</option>`).join('')}</select></label><button class="btn sm" ${view.page === 0 ? 'disabled' : ''} onclick="changeReviewPage(${pi},${view.page - 1})" aria-label="Предыдущая страница">← Назад</button><label>Страница <input type="number" min="1" max="${pages}" value="${view.page + 1}" aria-label="Номер страницы" onchange="changeReviewPage(${pi},Number(this.value)-1)"> <span>из ${pages}</span></label><button class="btn sm" ${view.page + 1 >= pages ? 'disabled' : ''} onclick="changeReviewPage(${pi},${view.page + 1})" aria-label="Следующая страница">Далее →</button></div></nav>`;
}
function renderPendingPage(pi) {
  const record = PENDING[pi], host = $('#pv' + pi); if (!record || !host) return;
  const view = zipReviewView(record); if (!view.open) { host.innerHTML = ''; return; }
  zipCaptureEditors();
  const count = record.items.length, pages = Math.max(1, Math.ceil(count / view.size));
  view.page = Math.max(0, Math.min(pages - 1, view.page));
  const start = view.page * view.size;
  const rows = record.items.slice(start, start + view.size).map((item, offset) => {
    const ii = start + offset;
    const via = item.via === 'override' ? '<span class="badge b-blue">правил куратор</span>' : item.via === 'match' ? '<span class="badge b-green">матч</span>' : item.via === 'keyword' ? '<span class="badge b-gray">словарь</span>' : '<span class="badge b-amber">не классиф.</span>';
    return `<tr data-item-index="${ii}"><td class="review-number">${ii + 1}</td><td class="review-name">${item.water ? '<span class="badge b-blue">вода</span> ' : ''}${esc(item.name)}</td><td class="jc">${item.cat ? esc(item.cat) : '<span class="badge b-amber">—</span>'}</td><td class="jg">${item.grp ? esc(item.grp) : '<span class="mini">—</span>'}</td><td class="num jq">${item.qty == null ? '<span class="badge b-rose">нет</span>' : fmt(item.qty)}</td><td class="ju">${!item.unit || normUnit(item.unit).unknown ? '<span class="badge b-rose">Укажите</span>' : esc(item.unit)}</td><td>${via}</td><td class="ja"><button class="btn sm ghost" onclick="editRow(this,${pi},${ii})">Исправить</button></td></tr>`;
  }).join('');
  const pager = zipReviewPager(pi, record, view, count);
  host.innerHTML = pager + `<div class="review-table-wrap" tabindex="0" role="region" aria-label="Позиции реестра"><table class="review-table"><thead><tr><th>№</th><th>Наименование</th><th>Категория</th><th>Группа</th><th>Кол-во</th><th>Ед.</th><th>Источник</th><th>Действия</th></tr></thead><tbody>${rows}</tbody></table></div>` + (count > view.size ? pager : '');
  for (let ii = start; ii < Math.min(count, start + view.size); ii++) {
    if (zipReviewEditors.has(record.items[ii])) zipRenderEditor(pi, ii, host.querySelector('[data-item-index="' + ii + '"]'));
  }
  const toggle = $('#reviewToggle' + pi); if (toggle) { toggle.textContent = 'Скрыть позиции'; toggle.setAttribute('aria-expanded', 'true'); }
}
async function changeReviewPage(pi, page, size) {
  const record = PENDING[pi]; if (!record) return;
  if (!await flushReviewEdits(record)) { renderPendingPage(pi); return; }
  const view = zipReviewView(record);
  if ([50,100,200].includes(size)) view.size = size;
  view.page = Number.isInteger(page) ? page : view.page;
  view.open = true; renderPendingPage(pi);
  const host = $('#pv' + pi); host.querySelector('.review-table-wrap')?.focus({preventScroll: true});
  host.scrollIntoView({behavior: 'smooth', block: 'start'});
}
async function showReviewIssue(pi, ii) {
  const record = PENDING[pi]; if (!record) return;
  const view = zipReviewView(record); await changeReviewPage(pi, Math.floor(ii / view.size));
}
async function toggleView(pi) {
  const record = PENDING[pi]; if (!record) return;
  const view = zipReviewView(record);
  if (view.open && !await flushReviewEdits(record)) return;
  view.open = !view.open; renderPendingPage(pi);
  const toggle = $('#reviewToggle' + pi); if (toggle) { toggle.textContent = view.open ? 'Скрыть позиции' : 'Посмотреть'; toggle.setAttribute('aria-expanded', String(view.open)); }
}
window.addEventListener('beforeunload', event => { if (zipReviewEditors.size || zipReviewSaves.size) { event.preventDefault(); event.returnValue = ''; } });
