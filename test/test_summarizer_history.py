import importlib
import json
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import httpx
import pytest
from fastapi.testclient import TestClient

from services.summarizer import storage, engine


@pytest.fixture
def archive(tmp_path, monkeypatch):
    monkeypatch.setattr(storage, 'DATA', tmp_path / 'summarizer')
    monkeypatch.setattr(storage, 'FILE', tmp_path / 'summarizer/reports.json')
    return storage


def result(text='Первоначальный результат', backend='algo'):
    return {'ok': True, 'report_text': text, 'backend': backend, 'stats': {}}


def test_new_history_preserves_entire_source_and_result(archive):
    source = 'Длинная исходная переписка\n' * 1500
    report = archive.create_report(source, result(), 'Иван')
    saved = archive.get_report(report['id'])
    assert saved['source'] == source and len(source) > 20000
    assert saved['versions'][0]['source'] == source
    assert saved['versions'][0]['result']['report_text'] == 'Первоначальный результат'
    assert saved['source_complete'] and not saved.get('source_may_be_truncated')
    assert archive.FILE.exists()


def test_pagination_previews_do_not_ship_full_archive(archive):
    for number in range(25):
        archive.create_report(f'Исходник {number}\n' * 40, result(f'Результат {number}' * 40), 'Иван')
    page = archive.history_page(20, 0)
    assert len(page['items']) == 20 and page['total'] == 25 and page['has_more']
    assert page['items'][0]['source_preview'].startswith('Исходник 24')
    assert all('source' not in item and 'result' not in item and 'versions' not in item for item in page['items'])
    second = archive.history_page(20, 20)
    assert len(second['items']) == 5 and not second['has_more']
    assert not set(item['id'] for item in page['items']) & set(item['id'] for item in second['items'])


def test_regeneration_persists_result_comment_and_both_versions(archive):
    report = archive.create_report('Оригинальный текст', result(), 'Иван')
    rid = report['id']
    archive.approve(rid, 'Анна')
    archive.reject(rid, 'Борис', 'Уточнить срок')
    request = 'Оригинальный текст\nКомментарий для корректировки: Уточнить срок'
    archive.regenerate(rid, result('Уточнённый результат', 'qwen'), 'Борис', request, 'Уточнить срок')
    saved = archive.get_report(rid)
    assert saved['source'] == 'Оригинальный текст'
    assert saved['result']['report_text'] == 'Уточнённый результат'
    assert saved['status'] == 'pending' and saved['approvals'] == {} and saved['revision_comment'] == ''
    assert len(saved['versions']) == 2
    assert saved['versions'][0]['result']['report_text'] == 'Первоначальный результат'
    assert saved['versions'][1]['source'] == request and saved['versions'][1]['comment'] == 'Уточнить срок'
    assert saved['rejects'][0]['user'] == 'Борис'


def test_old_reports_are_read_without_rewriting_or_inventing_source(archive):
    archive.DATA.mkdir()
    legacy = [{'id': 'old', 'source': 'а' * 20000, 'result': result(), 'author': 'Иван', 'created_at': '01.10.2026 10:00', 'status': 'pending'}]
    archive.FILE.write_text(json.dumps(legacy))
    original = archive.FILE.read_bytes()
    saved = archive.get_report('old')
    assert len(saved['versions']) == 1 and saved['source_may_be_truncated']
    assert archive.FILE.read_bytes() == original
    archive.regenerate('old', result('Новая версия'), 'Анна', saved['source'])
    assert len(archive.get_report('old')['versions']) == 2
    assert archive.get_report('old')['source_may_be_truncated']


def test_threaded_reports_are_not_lost(archive):
    with ThreadPoolExecutor(max_workers=6) as pool:
        list(pool.map(lambda n: archive.create_report(str(n), result(), 'Иван'), range(30)))
    assert archive.history_page()['total'] == 30


def test_damaged_archive_is_not_silently_overwritten(archive):
    archive.DATA.mkdir(); archive.FILE.write_text('damaged')
    with pytest.raises(ValueError):
        archive.create_report('Новый текст', result(), 'Иван')
    assert archive.FILE.read_text() == 'damaged'


def test_two_distinct_approvals_are_still_required(archive):
    rid = archive.create_report('Текст', result(), 'Иван')['id']
    archive.approve(rid, 'Иван'); archive.approve(rid, 'Иван')
    assert archive.get_report(rid)['status'] == 'pending'
    assert archive.approve(rid, 'Анна')['status'] == 'approved'


@pytest.fixture
def client(archive, tmp_path, monkeypatch):
    import utils.db as db
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'history.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr('services.auth.activity.track_response', lambda *a: None)
    with patch('services.scheduler.start'):
        app = importlib.import_module('app')
    user = {'username': 'tester', 'role': 'user', 'modules': ['summarizer']}
    monkeypatch.setattr(app, 'get_user_from_token', lambda *a: user)
    monkeypatch.setattr('services.auth.security.get_user_from_token', lambda *a: user)
    with TestClient(app.app) as c:
        c.cookies.set('access_token', 'test-only')
        yield c, app, user


def test_complete_user_path_create_reload_read_revise_and_regenerate(client):
    c, app, user = client
    source = 'Переписка с исходными данными. ' * 2000
    with patch.object(app.sum_engine, 'summarize', return_value=result('Готовый отчёт', 'qwen')):
        created = c.post('/summarizer/api/summary', json={'text': source, 'backend': 'qwen'}).json()
    rid = created['report']['id']
    listing = c.get('/summarizer/api/reports').json()
    assert listing['items'][0]['id'] == rid and listing['total'] == 1
    response = c.get('/summarizer/api/reports/' + rid)
    assert response.headers['cache-control'] == 'no-store'
    assert response.json()['report']['source'] == source.strip()
    c.post('/summarizer/api/reject', json={'id': rid, 'comment': 'Уточни факты'})
    with patch.object(app.sum_engine, 'summarize', return_value=result('Исправленный отчёт', 'qwen')) as summarize:
        assert c.post('/summarizer/api/regenerate', json={'id': rid}).json()['ok']
        assert 'Уточни факты' in summarize.call_args.args[0]
        assert summarize.call_args.kwargs['backend'] == 'qwen'
    saved = c.get('/summarizer/api/reports/' + rid).json()['report']
    assert saved['result']['report_text'] == 'Исправленный отчёт'
    assert saved['versions'][0]['result']['report_text'] == 'Готовый отчёт'
    assert saved['versions'][1]['result']['report_text'] == 'Исправленный отчёт'
    assert c.get('/summarizer/api/reports').json()['items'][0]['version_count'] == 2


def test_history_endpoints_require_summarizer_permission(client):
    c, app, user = client
    rid = storage.create_report('Private input', result(), 'Иван')['id']
    user['modules'] = []
    for path in ('/summarizer/api/reports', '/summarizer/api/reports/' + rid):
        response = c.get(path)
        assert response.status_code == 403 and 'Private input' not in response.text


def test_new_revision_during_ai_generation_is_preserved(client):
    c, app, user = client
    rid = storage.create_report('Достаточно длинный исходный текст для проверки пересборки.', result(), 'Иван')['id']
    storage.reject(rid, 'Анна', 'Первый комментарий')
    def slow_generation(*args, **kwargs):
        storage.reject(rid, 'Борис', 'Новый комментарий во время генерации')
        return result('Устаревший результат')
    with patch.object(app.sum_engine, 'summarize', side_effect=slow_generation):
        response = c.post('/summarizer/api/regenerate', json={'id': rid})
    assert response.status_code == 409
    saved = c.get('/summarizer/api/reports/' + rid).json()['report']
    assert saved['revision_comment'] == 'Новый комментарий во время генерации'
    assert saved['status'] == 'revision'
    assert saved['result']['report_text'] == 'Первоначальный результат'
    assert len(saved['versions']) == 1
    assert response.json()['report'] == saved


def test_qwen_request_persists_actual_answer_in_history(client, monkeypatch):
    c, app, user = client
    monkeypatch.setenv('QWEN_API_BASE', 'https://aiplatform.mosreg.ru/api/user-models/v1')
    monkeypatch.setenv('QWEN_API_KEY', 'synthetic-summary-key')
    monkeypatch.setenv('QWEN_MODEL', 'synthetic-qwen')
    monkeypatch.delenv('SUMMARIZER_QWEN_API_KEY', raising=False)
    monkeypatch.delenv('AI_CA_BUNDLE', raising=False)
    source = 'Синтетическая переписка для проверки сумматора: работы завершатся в 15:00.'
    calls = []
    def post(url, **kwargs):
        calls.append(url)
        assert url == 'https://aiplatform.mosreg.ru/api/user-models/v1/chat/completions'
        assert kwargs['headers']['Authorization'] == 'Bearer synthetic-summary-key'
        assert kwargs['json']['messages'][-1]['content'] == source
        data = {'choices': [{'message': {'content': 'Работы завершатся в 15:00.'}}]}
        return httpx.Response(200, json=data, request=httpx.Request('POST', url))
    monkeypatch.setattr(httpx, 'post', post)
    created = c.post('/summarizer/api/summary', json={'text': source, 'backend': 'qwen'})
    assert created.status_code == 200
    rid = created.json()['report']['id']
    saved = c.get('/summarizer/api/reports/' + rid).json()['report']
    assert saved['source'] == source
    assert saved['result']['backend'] == 'qwen'
    assert saved['result']['report_text'] == 'Работы завершатся в 15:00.'
    assert saved['versions'][0]['result'] == saved['result']
    assert c.get('/summarizer/api/reports').json()['items'][0]['backend'] == 'qwen'
    assert len(calls) == 1


def test_missing_report_and_invalid_size(client):
    c, app, user = client
    assert c.get('/summarizer/api/reports/missing').status_code == 404
    with patch.object(app.sum_engine, 'summarize') as summarize:
        response = c.post('/summarizer/api/summary', json={'text': 'a' * 200001})
        assert response.status_code == 400
        summarize.assert_not_called()
    assert c.get('/summarizer/api/reports?offset=-1&limit=999').status_code == 200


def test_summarizer_fallback_is_labeled_and_never_logs_key(caplog, monkeypatch):
    url = 'https://aiplatform.mosreg.ru/api/user-models/v1/chat/completions'
    request = httpx.Request('POST', url)
    response = httpx.Response(401, json={'error': 'SECRET_KEY'}, request=request)
    failure = httpx.HTTPStatusError('SECRET_KEY', request=request, response=response)
    with patch.object(engine, '_qwen_chat', side_effect=failure):
        answer = engine.summarize('Текст оперативного сообщения об аварии на водопроводе. ' * 3, backend='qwen')
    assert answer['ok'] and answer['backend'] == 'algo'
    assert answer['ai_error_code'] == 'AI_AUTH_FAILED' and answer['requested_backend'] == 'qwen'
    assert 'обычным алгоритмом' in answer['warning']
    assert 'SECRET_KEY' not in caplog.text + json.dumps(answer)


def test_delete_one_report_preserves_others_and_removes_all_versions(client):
    c, app, user = client
    kept = storage.create_report('Запись, которую нужно оставить', result(), 'Анна')
    removed = storage.create_report('Удаляемый уникальный исходник', result('Уникальная старая версия'), 'Борис')
    rid = removed['id']
    storage.regenerate(rid, result('Уникальная новая версия'), 'Борис', 'Удаляемая доработка')
    assert c.delete('/summarizer/api/reports/' + rid).status_code == 200
    assert c.get('/summarizer/api/reports/' + rid).status_code == 404
    assert c.delete('/summarizer/api/reports/' + rid).status_code == 404
    listing = c.get('/summarizer/api/reports').json()
    assert listing['total'] == 1 and listing['items'][0]['id'] == kept['id']
    assert storage.get_report(kept['id'])['source'] == kept['source']
    on_disk = storage.FILE.read_text()
    assert rid not in on_disk and 'Уникальная' not in on_disk and 'Удаляем' not in on_disk


def test_clear_history_includes_all_pages_and_is_persistent(client):
    c, app, user = client
    for index in range(25):
        storage.create_report('Исходник ' + str(index), result(), 'Анна')
    assert c.get('/summarizer/api/reports').json()['has_more']
    response = c.delete('/summarizer/api/reports')
    assert response.json() == {'ok': True, 'deleted': 25}
    assert c.get('/summarizer/api/reports').json()['total'] == 0
    assert json.loads(storage.FILE.read_text()) == []
    assert c.delete('/summarizer/api/reports').json()['deleted'] == 0


def test_delete_requires_module_access_and_same_origin(client):
    c, app, user = client
    rid = storage.create_report('Сохранённый текст', result(), 'Анна')['id']
    for path in ('/summarizer/api/reports', '/summarizer/api/reports/' + rid):
        assert c.delete(path, headers={'origin': 'https://foreign.invalid'}).status_code == 403
    user['modules'] = []
    for path in ('/summarizer/api/reports', '/summarizer/api/reports/' + rid):
        assert c.delete(path).status_code == 403
    assert storage.get_report(rid) is not None


@pytest.mark.parametrize('clear_all', [False, True])
def test_regeneration_cannot_restore_a_deleted_report(client, clear_all):
    c, app, user = client
    rid = storage.create_report('Достаточно длинный исходный текст для проверки пересборки.', result(), 'Иван')['id']
    def generation_during_delete(*args, **kwargs):
        if clear_all:
            storage.clear_reports()
        else:
            storage.delete_report(rid)
        return result('Этот ответ уже не нужно сохранять')
    with patch.object(app.sum_engine, 'summarize', side_effect=generation_during_delete):
        response = c.post('/summarizer/api/regenerate', json={'id': rid})
    assert response.status_code == 404
    assert storage.get_report(rid) is None
    assert storage.history_page()['total'] == 0


def test_delete_does_not_overwrite_a_damaged_archive(archive):
    archive.DATA.mkdir(); archive.FILE.write_text('damaged')
    with pytest.raises(ValueError):
        archive.clear_reports()
    with pytest.raises(ValueError):
        archive.delete_report('missing')
    assert archive.FILE.read_text() == 'damaged'
