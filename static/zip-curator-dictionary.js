// Category management uses the server dictionary, including curator additions.
let categoryTarget = null;
let categorySaving = false;

async function refreshCuratorDictionary() {
  const response = await fetch('/zip_curator/api/dictionary', {cache: 'no-store'});
  if (!response.ok) throw new Error('Не удалось загрузить справочник');
  Object.assign(D, await response.json());
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
  if (categorySaving) return;
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
    Object.assign(D, data.dictionary);
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
      toggleView(pi);
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

refreshCuratorDictionary().catch(error => toast(error.message, 5000));
