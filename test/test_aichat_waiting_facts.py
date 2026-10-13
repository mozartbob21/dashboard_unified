"""Offline tests for generic waiting facts; no real model or user data is used."""
import json
import threading
from datetime import datetime, timezone
from types import SimpleNamespace

import pytest

from services.aichat import waiting_facts as facts
from services.summarizer import engine as provider


def fact_items():
    return [
        {'topic': 'ИИ', 'text': 'Языковые модели обрабатывают текст как последовательность небольших фрагментов.'},
        {'topic': 'ИИ', 'text': 'Нейронные сети обучаются находить закономерности на примерах из данных.'},
        {'topic': 'ИИ', 'text': 'Качество исходных данных влияет на результат обучения модели.'},
        {'topic': 'ЖКХ', 'text': 'Теплоизоляция помогает уменьшить потери тепла через стены здания.'},
        {'topic': 'ЖКХ', 'text': 'Вентиляция обеспечивает обмен воздуха между помещением и улицей.'},
        {'topic': 'ЖКХ', 'text': 'Водопровод и канализация выполняют разные задачи в инженерной системе дома.'},
    ]


@pytest.fixture
def state(monkeypatch, tmp_path):
    clock = SimpleNamespace(wall=1770000000.0, monotonic=1000.0)
    monkeypatch.setattr(facts, 'time', SimpleNamespace(time=lambda: clock.wall,
                                                    monotonic=lambda: clock.monotonic))
    monkeypatch.setattr(facts, 'DATA_DIR', tmp_path)
    for name, value in {'_cache': None, '_loaded': False, '_enabled': False,
                        '_preparing': False, '_active_answers': 0, '_retry_at': 0.0}.items():
        monkeypatch.setattr(facts, name, value)
    workers, calls = [], []

    class DeferredWorker:
        def __init__(self, *, target, name, daemon):
            assert daemon is True
            self.target = target

        def start(self):
            workers.append(self)

    monkeypatch.setattr(facts, 'threading', SimpleNamespace(Thread=DeferredWorker))

    def model(messages, **kwargs):
        calls.append({'messages': messages, **kwargs})
        return json.dumps({'items': fact_items()}, ensure_ascii=False)

    monkeypatch.setattr(provider, '_qwen_chat', model)
    return SimpleNamespace(clock=clock, workers=workers, calls=calls, root=tmp_path)


def cache_file(state, *, age=0, items=None):
    generated_at = datetime.fromtimestamp(state.clock.wall - age, timezone.utc).isoformat(timespec='seconds')
    value = {'version': 1, 'generated_at': generated_at, 'items': items or fact_items()}
    path = state.root / 'waiting_facts.json'
    path.write_text(json.dumps(value, ensure_ascii=False), encoding='utf-8')
    return path, value


def test_get_is_nonblocking_single_flight_and_uses_only_generic_model_prompt(state):
    assert facts.get_facts() == {'items': [], 'generated_at': None, 'preparing': True}
    for _ in range(5):
        assert facts.get_facts()['preparing'] is True
    assert len(state.workers) == 1 and state.calls == []
    state.workers[0].target()
    result = facts.get_facts()
    assert result['items'] == fact_items() and result['preparing'] is False
    call, = state.calls
    assert [message['role'] for message in call['messages']] == ['system', 'user']
    assert call['messages'][-1]['content'] == facts._PROMPT
    assert set(call) == {'messages', 'max_tokens', 'timeout'}
    assert call['max_tokens'] == 1200
    assert call['timeout'].read == 45 and call['timeout'].connect == 10
    assert call['timeout'].write == 10 and call['timeout'].pool == 5
    assert result['generated_at'].endswith('+00:00')
    saved = json.loads((state.root / 'waiting_facts.json').read_text())
    assert saved == {'version': 1, 'items': fact_items(), 'generated_at': result['generated_at']}
    assert list(state.root.glob('.waiting-facts-*')) == []


def test_valid_disk_cache_is_reused_and_public_results_are_copies(state):
    _, saved = cache_file(state)
    result = facts.get_facts()
    assert result == {'items': fact_items(), 'generated_at': saved['generated_at'], 'preparing': False}
    result['items'][0]['text'] = 'Caller mutation must not poison the cache'
    assert facts.get_facts()['items'] == fact_items()
    assert state.workers == [] and state.calls == []


def test_stale_cache_remains_available_while_refreshed(state):
    _, saved = cache_file(state, age=facts._TTL_SECONDS)
    result = facts.get_facts()
    assert result['items'] == fact_items() and result['preparing'] is True
    assert result['generated_at'] == saved['generated_at']
    state.workers[0].target()
    assert facts.get_facts()['generated_at'] != saved['generated_at']
    assert facts.get_facts()['preparing'] is False


def test_failed_refresh_preserves_cache_and_waits_for_cooldown(state, monkeypatch):
    path, saved = cache_file(state, age=facts._TTL_SECONDS + 1)
    previous = path.read_bytes()

    def fail(*args, **kwargs):
        raise RuntimeError('PRIVATE_PROVIDER_RESPONSE_AND_KEY')

    monkeypatch.setattr(provider, '_qwen_chat', fail)
    facts.get_facts()
    state.workers[0].target()
    result = facts.get_facts()
    assert result['items'] == fact_items() and result['generated_at'] == saved['generated_at']
    assert result['preparing'] is False and path.read_bytes() == previous
    state.clock.monotonic += 899
    facts.get_facts()
    assert len(state.workers) == 1
    state.clock.monotonic += 1
    assert facts.get_facts()['preparing'] is True
    assert len(state.workers) == 2


def test_disk_failure_keeps_existing_file_and_valid_cache(state, monkeypatch):
    path, _ = cache_file(state, age=facts._TTL_SECONDS + 1)
    previous = path.read_bytes()
    monkeypatch.setattr(facts.os, 'replace', lambda *args: (_ for _ in ()).throw(OSError('disk unavailable')))
    facts.get_facts()
    state.workers[0].target()
    assert facts.get_facts()['items'] == fact_items()
    assert path.read_bytes() == previous
    assert list(state.root.glob('.waiting-facts-*')) == []
    assert len(state.workers) == 1


def test_answers_defer_refresh_until_last_answer_finishes(state):
    facts.begin_answer()
    facts.begin_answer()
    assert facts.get_facts() == {'items': [], 'generated_at': None, 'preparing': False}
    facts.end_answer()
    assert state.workers == []
    facts.end_answer()
    assert len(state.workers) == 1
    state.workers[0].target()
    assert facts.get_facts()['items'] == fact_items()


def test_answers_do_not_activate_feature_without_a_fact_request(state):
    facts.begin_answer()
    facts.end_answer()
    assert state.workers == [] and state.calls == []


def test_read_only_lookup_does_not_activate_generation_or_post_answer_warming(state):
    assert facts.get_facts(warm=False) == {'items': [], 'generated_at': None, 'preparing': False}
    facts.begin_answer()
    facts.end_answer()
    assert state.workers == [] and state.calls == []
    assert facts.get_facts()['preparing'] is True


def test_read_only_lookup_can_show_old_facts_without_triggering_refresh(state):
    _, saved = cache_file(state, age=facts._TTL_SECONDS + 1)
    assert facts.get_facts(warm=False) == {'items': fact_items(), 'generated_at': saved['generated_at'],
                                          'preparing': False}
    assert state.workers == [] and state.calls == []


def test_cancelled_answer_can_release_without_warming(state):
    facts.begin_answer()
    facts.get_facts()
    facts.end_answer(warm=False)
    assert state.workers == [] and facts._active_answers == 0


def test_answer_starting_before_worker_defers_network_until_idle(state):
    facts.get_facts()
    facts.begin_answer()
    state.workers[0].target()
    assert state.calls == []
    assert facts.get_facts()['preparing'] is False
    facts.end_answer()
    assert len(state.workers) == 2
    state.workers[1].target()
    assert len(state.calls) == 1


@pytest.mark.parametrize('text', [
    'Коротко', 'Факт ' + 'я' * 321,
    'Подробности доступны по ссылке https://example.org/topic',
    'Подробности доступны по ссылке example.org/topic',
    '<b>Факт с недопустимой разметкой внутри</b>',
    '[Подробнее](example) о факте с форматированной ссылкой.',
    'Сегодня применяется временная мера в городском хозяйстве.',
    'Текущая доля выросла до 20 процентов за последний год.',
])
def test_invalid_fact_response_never_becomes_fallback_content(state, monkeypatch, text):
    items = fact_items()
    items[0]['text'] = text
    monkeypatch.setattr(provider, '_qwen_chat', lambda *args, **kwargs: json.dumps({'items': items}))
    facts.get_facts()
    state.workers[0].target()
    assert facts.get_facts() == {'items': [], 'generated_at': None, 'preparing': False}
    assert not (state.root / 'waiting_facts.json').exists()
    assert len(state.workers) == 1


@pytest.mark.parametrize('kind', ['duplicate', 'single-topic', 'too-few', 'extra-field', 'malformed-json'])
def test_invalid_shape_and_topic_balance_are_rejected(state, monkeypatch, kind):
    items = fact_items()
    if kind == 'duplicate':
        items[1] = dict(items[0])
    elif kind == 'single-topic':
        items = [{**item, 'topic': 'ИИ'} for item in items]
    elif kind == 'too-few':
        items = items[:3]
    elif kind == 'extra-field':
        items[0]['secret'] = 'unexpected'
    raw = '{bad json' if kind == 'malformed-json' else json.dumps({'items': items})
    monkeypatch.setattr(provider, '_qwen_chat', lambda *args, **kwargs: raw)
    facts.get_facts()
    state.workers[0].target()
    assert facts.get_facts()['items'] == []


def test_corrupt_disk_cache_is_ignored_and_replaced_only_after_valid_generation(state):
    path = state.root / 'waiting_facts.json'
    path.write_text('{invalid cache', encoding='utf-8')
    assert facts.get_facts()['items'] == []
    state.workers[0].target()
    assert facts.get_facts()['items'] == fact_items()
    assert json.loads(path.read_text())['items'] == fact_items()


def test_real_worker_returns_immediately_and_never_overlaps(state, monkeypatch):
    started, released, finished = threading.Event(), threading.Event(), threading.Event()
    actual_threads = []

    class Worker(threading.Thread):
        def __init__(self, **kwargs):
            super().__init__(**kwargs)
            actual_threads.append(self)

    def blocked_model(*args, **kwargs):
        started.set()
        assert released.wait(3), 'Test did not release the model'
        finished.set()
        return json.dumps({'items': fact_items()})

    monkeypatch.setattr(facts, 'threading', SimpleNamespace(Thread=Worker))
    monkeypatch.setattr(provider, '_qwen_chat', blocked_model)
    try:
        assert facts.get_facts()['preparing'] is True
        assert started.wait(3)
        assert facts.get_facts()['items'] == []
        assert len(actual_threads) == 1 and not finished.is_set()
    finally:
        released.set()
        for worker in actual_threads:
            worker.join(3)
    assert facts.get_facts()['items'] == fact_items()
