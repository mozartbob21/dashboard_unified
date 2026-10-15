// Category management uses the server dictionary, including curator additions.
let categoryTarget = null;
let categorySaving = false;

let dictionaryRevision = 0;
let dictionaryReadId = 0;
let dictionaryImporting = false;
let dictionaryDownloading = false;

function applyCuratorDictionary(dictionary) {
  if (!dictionary || typeof dictionary !== 'object' ||
      !dictionary.dict || typeof dictionary.dict !== 'object' || Array.isArray(dictionary.dict) ||
      !Array.isArray(dictionary.categories) || !dictionary.groupsByCat ||
      typeof dictionary.groupsByCat !== 'object' || Array.isArray(dictionary.groupsByCat)) {
    throw new Error('Сервер вернул неполный справочник. Повторите загрузку.');
  }
  Object.assign(D, dictionary);
  dictionaryRevision++;
  dictDirty = false;
  render();
}

function dictionaryErrorMessage(value, fallback) {
  return typeof value === 'string' && value.trim() && value.length <= 500 && /[А-Яа-яЁё]/.test(value) && !/https?:\/\//i.test(value)
    ? value.trim() : fallback;
}

async function refreshCuratorDictionary() {
  const revision = dictionaryRevision, readId = ++dictionaryReadId;
  try {
    const response = await fetch('/zip_curator/api/dictionary', {cache: 'no-store'});
    if (!response.ok) throw new Error('Не удалось загрузить справочник с сервера');
    const dictionary = await response.json();
    // An earlier GET must never replace a dictionary saved by a later import.
    if (revision !== dictionaryRevision || readId !== dictionaryReadId || dictionaryImporting) return false;
    applyCuratorDictionary(dictionary);
    return true;
  } catch (error) {
    if (revision !== dictionaryRevision || readId !== dictionaryReadId) return false;
    throw error;
  }
}

function setDictionaryTransferBusy(busy) {
  for (const id of ['btnLoadDict', 'btnSaveDict', 'btnCategories', 'fiDict']) {
    document.getElementById(id).disabled = busy;
  }
}

async function importCuratorDictionary(file) {
  if (!file || dictionaryImporting) return false;
  if (typeof flushReviewEdits === 'function' && !await flushReviewEdits()) return false;
  if (dictionaryDownloading || categorySaving) {
    toast('Дождитесь завершения текущей операции со справочником.', 4000);
    return false;
  }
  const button = document.getElementById('btnLoadDict'), label = button.textContent;
  dictionaryImporting = true;
  dictionaryRevision++;
  setDictionaryTransferBusy(true);
  button.textContent = 'Сохраняем справочник…';
  button.setAttribute('aria-busy', 'true');
  try {
    const form = new FormData();
    form.append('file', file);
    const response = await fetch('/zip_curator/api/dictionary/import', {method: 'POST', body: form});
    let data;
    try { data = await response.json(); } catch (error) { data = null; }
    if (!response.ok || !data || data.ok !== true) {
      throw new Error(dictionaryErrorMessage(data && data.error, 'Не удалось сохранить справочник на сервере. Повторите загрузку.'));
    }
    applyCuratorDictionary(data.dictionary);
    if (typeof loadServerState === 'function') await loadServerState();
    const imported = Number.isSafeInteger(data.imported) && data.imported >= 0 ? data.imported : 0;
    toast('Справочник сохранён на сервере: ' + imported.toLocaleString('ru-RU') + ' записей.', 5000);
    return true;
  } catch (error) {
    toast(dictionaryErrorMessage(error.message, 'Не удалось подтвердить сохранение справочника. Проверьте соединение и повторите загрузку.'), 6000);
    return false;
  } finally {
    dictionaryImporting = false;
    setDictionaryTransferBusy(false);
    button.textContent = label;
    button.removeAttribute('aria-busy');
    document.getElementById('fiDict').value = '';
  }
}

async function downloadCuratorDictionary() {
  if (dictionaryDownloading) return false;
  if (dictionaryImporting || categorySaving) {
    toast('Дождитесь сохранения справочника на сервере.', 4000);
    return false;
  }
  const button = document.getElementById('btnSaveDict'), label = button.textContent;
  dictionaryDownloading = true;
  setDictionaryTransferBusy(true);
  button.textContent = 'Подготавливаем файл…';
  button.setAttribute('aria-busy', 'true');
  try {
    if (!await refreshCuratorDictionary()) throw new Error('Справочник изменился во время подготовки файла. Повторите скачивание.');
    const rows = [['Наименование (норм.)', 'Категория', 'Группа', 'Вода']];
    Object.entries(D.dict).forEach(([name, item]) => rows.push([name, item[0] || '', item[1] || '', item[2] ? 'да' : '']));
    const sheet = XLSX.utils.aoa_to_sheet(rows), workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Справочник');
    XLSX.writeFile(workbook, 'справочник.xlsx');
    toast('Справочник с сервера скачан');
    return true;
  } catch (error) {
    toast(dictionaryErrorMessage(error.message, 'Не удалось скачать справочник. Повторите попытку.'), 6000);
    return false;
  } finally {
    dictionaryDownloading = false;
    setDictionaryTransferBusy(false);
    button.textContent = label;
    button.removeAttribute('aria-busy');
  }
}

function categoryOptions() {
  const options = document.getElementById('categoryOptions');
  options.replaceChildren(...D.categories.map(name => {
    const option = document.createElement('option');
    option.value = name;
    return option;
  }));
  categoryGroupOptions();
}

function categoryGroupOptions() {
  const name = document.getElementById('categoryName').value.trim().toLocaleLowerCase('ru');
  const category = D.categories.find(value => value.toLocaleLowerCase('ru') === name);
  const groups = D.groupsByCat[category] || [];
  document.getElementById('categoryGroupOptions').replaceChildren(...groups.map(name => {
    const option = document.createElement('option');
    option.value = name;
    return option;
  }));
}

async function openCategoryEditor(pi = null, ii = null) {
  if (categorySaving || dictionaryImporting || dictionaryDownloading) return;
  const dialog = document.getElementById('categoryDialog');
  const item = pi === null ? null : PENDING[pi]?.items[ii];
  categoryTarget = item ? {pi, ii, item} : null;
  document.getElementById('categoryForm').reset();
  document.getElementById('categoryName').value = item?.cat || '';
  document.getElementById('categoryGroup').value = item?.grp || '';
  document.getElementById('categoryContext').textContent = item ? 'Для позиции: ' + item.name : '';
  document.getElementById('categorySave').textContent = item ? 'Сохранить и назначить' : 'Сохранить';
  document.getElementById('categoryError').textContent = '';
  dialog.showModal();
  document.getElementById('categorySave').disabled = true;
  try {
    await refreshCuratorDictionary();
    categoryOptions();
    document.getElementById('categorySave').disabled = false;
    document.getElementById('categoryName').focus();
  } catch (error) {
    document.getElementById('categoryError').textContent = error.message + '. Закройте окно и попробуйте снова.';
  }
}

document.getElementById('btnCategories').onclick = () => openCategoryEditor();
document.getElementById('categoryName').addEventListener('input', categoryGroupOptions);
document.getElementById('categoryCancel').onclick = () => {
  if (!categorySaving) document.getElementById('categoryDialog').close();
};
document.getElementById('categoryDialog').addEventListener('cancel', event => {
  if (categorySaving) event.preventDefault();
});
document.getElementById('categoryForm').addEventListener('submit', async event => {
  event.preventDefault();
  if (categorySaving) return;
  categorySaving = true;
  dictionaryRevision++;
  const save = document.getElementById('categorySave');
  save.disabled = true;
  document.getElementById('categoryError').textContent = '';
  try {
    const response = await fetch('/zip_curator/api/categories', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({category: document.getElementById('categoryName').value,
                            group: document.getElementById('categoryGroup').value})
    });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || 'Не удалось сохранить категорию');
    applyCuratorDictionary(data.dictionary);
    if (categoryTarget) {
      const {pi, ii, item} = categoryTarget;
      if (PENDING[pi]?.items[ii] !== item) throw new Error('Очередь изменилась. Откройте позицию повторно; категория уже сохранена.');
      const previous = {...item};
      item.cat = data.category;
      item.grp = data.group || null;
      if (!await persistEdit(pi, ii)) {
        Object.assign(item, previous);
        throw new Error('Категория сохранена, но назначение позиции не удалось. Повторите сохранение.');
      }
      render();
      zipReviewView(PENDING[pi]).open = true;
      renderPendingPage(pi);
    }
    document.getElementById('categoryDialog').close();
    toast(categoryTarget ? 'Категория назначена. Позиция сохранена в справочнике.' :
      data.created ? 'Сохранено в справочник' : 'Такая категория и группа уже есть в справочнике');
  } catch (error) {
    document.getElementById('categoryError').textContent = error.message;
  } finally {
    categorySaving = false;
    save.disabled = false;
  }
});

refreshCuratorDictionary().catch(error => toast(dictionaryErrorMessage(error.message, 'Не удалось загрузить справочник с сервера. Повторите попытку.'), 5000));
