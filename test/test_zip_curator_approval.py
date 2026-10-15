"""Curator approval requires complete, current data and persists explicit unit edits."""
import copy
import json

import pytest

from services.zip_curator import core, dictionary


@pytest.fixture
def store(monkeypatch, tmp_path):
    for key, value in {
        'DATA': tmp_path,
        'STATE_FILE': tmp_path / 'state.json',
        'PUBLISHED_FILE': tmp_path / 'published.json',
        'INPUT_DIR': tmp_path / 'input',
        'FOLDER_A_DIR': tmp_path / 'input/folder_a',
        'FOLDER_A_MANIFEST_FILE': tmp_path / 'manifest.json',
        'MUNICIPALITY_OVERRIDES_FILE': tmp_path / 'municipalities.json',
    }.items():
        monkeypatch.setattr(core, key, value)
    monkeypatch.setattr(dictionary, 'CUSTOM_FILE', tmp_path / 'dictionary.json')
    dictionary._merged.cache_clear()
    yield tmp_path
    dictionary._merged.cache_clear()


def item(**extra):
    return {'name': 'Труба 219 ППУ', 'nn': 'труба 219 ппу', 'cat': 'Трубы',
            'grp': 'Труба электросварная', 'qty': 90, 'unit': 'пог. м',
            'unitRaw': 'пог. м', 'unitUnknown': False, 'water': False, **extra}


def registry(**extra):
    return {'rso': 'Тестовое РСО', 'okrug': 'Тестовый округ', 'date': '',
            'items': [item()], **extra}


def save(record):
    core.save_state({'pending': [record], 'clean': {}, 'excluded_rso': {}})


@pytest.mark.parametrize('municipality', ['', ' ', None, '—', '-', 'не определён'])
def test_no_approval_without_municipality_leaves_both_files_unchanged(store, municipality):
    save(registry(okrug=municipality))
    before = core.STATE_FILE.read_bytes()
    with pytest.raises(core.ApprovalValidationError) as failure:
        core.approve(0)
    assert failure.value.issues[0]['field'] == 'okrug'
    assert 'ОМСУ' in str(failure.value)
    assert core.STATE_FILE.read_bytes() == before
    assert not core.PUBLISHED_FILE.exists()
    assert len(core.load_state()['pending']) == 1


def test_mapping_resolves_missing_municipality_and_allows_approval(store):
    save(registry(okrug=''))
    core.save_municipality_override('Тестовое РСО', 'Королёв')
    assert core.load_state()['pending'][0]['approval_issues'] == []
    assert core.approve(0)
    assert core.load_state()['pending'] == []
    assert core.load_state()['clean']['тестовое рсо']['okrug'] == 'Королёв'
    assert json.loads(core.PUBLISHED_FILE.read_text())['rows'][1][1] == 'Королёв'


@pytest.mark.parametrize('patch, field', [
    ({'name': ''}, 'name'), ({'cat': None}, 'cat'), ({'cat': 'Поддельная категория'}, 'cat'),
    ({'grp': ''}, 'grp'), ({'grp': 'Поддельная группа'}, 'grp'),
    ({'unit': None, 'unitRaw': '', 'unitUnknown': True}, 'unit'),
    ({'unit': 'неизвестно', 'unitRaw': 'неизвестно', 'unitUnknown': True}, 'unit'),
    ({'qty': None}, 'qty'), ({'qty': ''}, 'qty'), ({'qty': -1}, 'qty'),
    ({'qty': True}, 'qty'), ({'qty': float('nan')}, 'qty'), ({'qty': float('inf')}, 'qty'),
])
def test_missing_and_invalid_row_fields_block_approval(store, patch, field):
    save(registry(items=[item(**patch)]))
    with pytest.raises(core.ApprovalValidationError) as failure:
        core.approve(0)
    assert any(issue['field'] == field and issue['item_index'] == 0 for issue in failure.value.issues)
    assert core.load_state()['clean'] == {}
    assert not core.PUBLISHED_FILE.exists()


def test_empty_registry_is_blocked(store):
    save(registry(items=[]))
    with pytest.raises(core.ApprovalValidationError) as failure:
        core.approve(0)
    assert any(issue['field'] == 'items' for issue in failure.value.issues)


def test_zero_quantity_is_valid_and_never_replaced_by_empty(store):
    save(registry(items=[item(qty=0)]))
    assert core.approve(0)
    assert core.load_state()['clean']['тестовое рсо']['items'][0]['qty'] == 0


def test_category_without_groups_can_be_approved(store):
    dictionary.add_category('Новая категория')
    save(registry(items=[item(cat='Новая категория', grp=None)]))
    assert core.approve(0)


def test_explicit_running_metre_and_quantity_edits_survive_reload_and_publication(store):
    save(registry(items=[item(unit=None, unitRaw='', unitUnknown=True, qty=None)]))
    assert core.edit_item(0, 0, 'Трубы', 'Труба электросварная', unit='погонный метр', qty='90,5')
    current = core.load_state()['pending'][0]
    assert current['approval_issues'] == []
    assert current['items'][0]['unit'] == current['items'][0]['unitRaw'] == 'пог. м'
    assert current['items'][0]['qty'] == 90.5
    assert core.approve(0, expected_revision=current['revision'])
    rows = json.loads(core.PUBLISHED_FILE.read_text())['rows']
    assert rows[1][5:7] == [90.5, 'пог. м']
    core.STATE_FILE.unlink()
    assert core.load_state()['clean']['тестовое рсо']['items'][0]['unit'] == 'пог. м'


def test_old_edit_contract_preserves_unit_and_quantity(store):
    save(registry())
    assert core.edit_item(0, 0, 'Трубы', 'Труба электросварная')
    updated = core.load_state()['pending'][0]['items'][0]
    assert (updated['unit'], updated['qty']) == ('пог. м', 90)


@pytest.mark.parametrize('changes', [
    {'unit': 'своя единица'}, {'unit': ['шт']}, {'qty': -1}, {'qty': True},
    {'qty': 'NaN'}, {'qty': 'Infinity'}, {'qty': {}}, {'qty': 'ошибка'},
])
def test_invalid_edit_is_atomic_and_does_not_teach_dictionary(store, changes):
    save(registry())
    before = core.STATE_FILE.read_bytes()
    with pytest.raises(ValueError):
        core.edit_item(0, 0, 'Трубы', 'Труба электросварная', **changes)
    assert core.STATE_FILE.read_bytes() == before
    assert not dictionary.CUSTOM_FILE.exists()


def test_explicit_clearing_unit_or_qty_is_a_draft_not_approval(store):
    save(registry())
    assert core.edit_item(0, 0, 'Трубы', 'Труба электросварная', unit=None, qty=None)
    current = core.load_state()['pending'][0]
    assert current['items'][0]['unit'] is None
    assert current['items'][0]['unitRaw'] == ''
    assert {x['field'] for x in current['approval_issues']} == {'unit', 'qty'}
    with pytest.raises(core.ApprovalValidationError):
        core.approve(0)


def test_same_name_in_different_units_is_never_summed(store):
    save(registry(items=[item(qty=5), item(qty=7), item(qty=2, unit='шт', unitRaw='шт')]))
    assert core.approve(0)
    items = core.load_state()['clean']['тестовое рсо']['items']
    assert {(x['unit'], x['qty']) for x in items} == {('пог. м', 12), ('шт', 2)}


def test_stale_tab_cannot_edit_or_approve_another_record(store):
    core.save_state({'pending': [registry(), registry(rso='Второе РСО')], 'clean': {}})
    revision = core.load_state()['pending'][0]['revision']
    assert core.approve(0, expected_revision=revision)
    with pytest.raises(core.CuratorConflictError):
        core.approve(0, expected_revision=revision)
    with pytest.raises(core.CuratorConflictError):
        core.edit_item(0, 0, 'Трубы', 'Труба электросварная', qty=1000, expected_revision=revision)
    assert core.load_state()['pending'][0]['items'][0]['qty'] == 90
    assert list(core.load_state()['clean']) == ['тестовое рсо']


def test_updated_row_changes_revision_and_rejects_old_edit(store):
    save(registry())
    original = core.load_state()['pending'][0]['revision']
    core.edit_item(0, 0, 'Трубы', 'Труба электросварная', qty=12, expected_revision=original)
    current = core.load_state()['pending'][0]['revision']
    assert current != original
    with pytest.raises(core.CuratorConflictError):
        core.edit_item(0, 0, 'Трубы', 'Труба электросварная', qty=99, expected_revision=original)
    assert core.load_state()['pending'][0]['items'][0]['qty'] == 12


def test_derived_status_is_recalculated_instead_of_persisted_or_trusted(store):
    save(registry(okrug='', approval_issues=[], revision='fake'))
    saved = json.loads(core.STATE_FILE.read_text())['pending'][0]
    assert 'approval_issues' not in saved and 'revision' not in saved
    loaded = core.load_state()['pending'][0]
    assert loaded['approval_issues'][0]['field'] == 'okrug'
    assert loaded['revision'] != 'fake'


def test_mutation_during_source_scan_returns_conflict(store):
    save(registry())
    with core._folder_sync_lock():
        with pytest.raises(core.CuratorConflictError):
            core.approve(0)
        with pytest.raises(core.CuratorConflictError):
            core.edit_item(0, 0, 'Трубы', 'Труба электросварная', qty=2)
    assert core.load_state()['pending'][0]['items'][0]['qty'] == 90


def test_failed_publication_keeps_pending_for_retry(store, monkeypatch):
    save(registry())
    before = core.STATE_FILE.read_bytes()
    def fail(*args, **kwargs):
        raise OSError('synthetic full disk')
    monkeypatch.setattr(core, 'publish_clean', fail)
    with pytest.raises(OSError):
        core.approve(0)
    assert core.STATE_FILE.read_bytes() == before
    assert not core.PUBLISHED_FILE.exists()


def test_fixing_approved_municipality_updates_published_rows_and_preserves_metadata(store):
    approved = registry(okrug='', loaded_at='2026-10-15T12:00:00', source_created_at='2026-10-14T10:00:00')
    other = registry(rso='Другое РСО')
    core.save_state({'pending': [], 'clean': {'тестовое рсо': approved, 'другое рсо': other}})
    core.publish_clean()
    core.save_municipality_override('Тестовое РСО', 'Королёв')
    published = json.loads(core.PUBLISHED_FILE.read_text())
    assert {row[0]: row[1] for row in published['rows'][1:]} == {'Тестовое РСО': 'Королёв', 'Другое РСО': 'Тестовый округ'}
    assert published['rso_metadata']['тестовое рсо']['loaded_at'] == '2026-10-15T12:00:00'
    core.STATE_FILE.unlink()
    assert core.load_state()['clean']['тестовое рсо']['okrug'] == 'Королёв'


def test_direct_publication_cannot_bypass_completeness_checks(store):
    rows = core.export_rows_from_state({'clean': {'test': registry(okrug='')}})
    with pytest.raises(core.ApprovalValidationError):
        core.publish_rows(rows)
    assert not core.PUBLISHED_FILE.exists()
    assert not core.STATE_FILE.exists()


def test_direct_publication_rejects_empty_names_instead_of_silently_dropping_rows(store):
    rows = core.export_rows_from_state({'clean': {'test': registry(items=[item(name='')])}})
    with pytest.raises(core.ApprovalValidationError):
        core.publish_rows(rows)
    assert not core.PUBLISHED_FILE.exists()


def test_direct_publication_retains_unchanged_legacy_and_validates_new_records(store):
    legacy = registry(rso='Старое РСО', okrug='', items=[item(cat=None, grp=None)])
    core.save_state({'pending': [], 'clean': {'старое рсо': legacy}})
    core.publish_clean()
    fresh = registry()
    rows = core.export_rows_from_state({'clean': {'старое рсо': legacy, 'тестовое рсо': fresh}})
    core.publish_rows(rows)
    assert set(core.load_state()['clean']) == {'старое рсо', 'тестовое рсо'}


def test_new_validation_failure_never_replaces_readable_state_with_empty_queue(store, monkeypatch):
    save(registry())
    before = core.STATE_FILE.read_bytes()
    def unexpected(*args, **kwargs):
        raise RuntimeError('synthetic validation bug')
    monkeypatch.setattr(core, '_with_approval_status', unexpected)
    with pytest.raises(RuntimeError):
        core.load_state()
    assert core.STATE_FILE.read_bytes() == before


def test_concurrent_approval_publishes_only_once_without_consuming_next_rso(store):
    from concurrent.futures import ThreadPoolExecutor
    from threading import Barrier
    core.save_state({'pending': [registry(), registry(rso='Следующее РСО')], 'clean': {}})
    revision = core.load_state()['pending'][0]['revision']
    start = Barrier(2)
    def approve_from_tab():
        start.wait()
        try:
            return core.approve(0, expected_revision=revision)
        except core.CuratorConflictError:
            return 'conflict'
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = [pool.submit(approve_from_tab) for _ in range(2)]
        results = [future.result() for future in futures]
    assert results.count(True) == 1
    assert results.count('conflict') == 1
    state = core.load_state()
    assert [record['rso'] for record in state['pending']] == ['Следующее РСО']
    assert list(state['clean']) == ['тестовое рсо']
