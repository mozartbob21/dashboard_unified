"""Synthetic HTTP at both edges; real PDF parsing, temp files and SQLite."""
from contextlib import contextmanager
import json
import ssl

import fitz
import httpx
import pytest

from services.max_bot.client import MaxClient
from services.max_bot.config import Settings
from services.max_bot.presentation_source import DOCUMENT_PATH, RedmineSource
from services.max_bot.presentations import PresentationService
from test.test_max_presentations import env, subscribe, advance, states


@pytest.fixture
def pipeline(env):
    with fitz.open() as document:
        document.new_page()
        document.new_page()
        pdf = document.tobytes()
    scenario = {'new': False, 'pdf': pdf, 'source_status': 200, 'upload_error': False,
                'message': {'body': {'mid': 'synthetic-pdf-message'}}, 'sent_bodies': [], 'uploaded': []}

    def source_response(request):
        if request.url.path == DOCUMENT_PATH:
            link = ('<p><a class="icon-attachment" href="/attachments/download/123/weekly.pdf">'
                    'Недельный отчёт.pdf</a><span class="author">04.11.2026 19:35</span></p>') if scenario['new'] else ''
            return httpx.Response(200, text='<div class="attachments">' + link + '</div>')
        assert request.url.path == '/attachments/download/123/weekly.pdf'
        return httpx.Response(scenario['source_status'], content=scenario['pdf'], headers={'content-type': 'application/pdf'})

    @contextmanager
    def source_factory():
        with RedmineSource('synthetic-user', 'synthetic-password', transport=httpx.MockTransport(source_response)) as source:
            yield source

    def max_response(request):
        if request.url.path == '/uploads':
            return httpx.Response(200, json={'url': 'https://fu.oneme.ru/api/upload.do?sig=synthetic'})
        if request.url.host == 'fu.oneme.ru':
            if scenario['upload_error']:
                raise ssl.SSLCertVerificationError(1, 'SYNTHETIC_PRIVATE_DIAGNOSTIC')
            scenario['uploaded'].append(request.read())
            assert 'authorization' not in request.headers
            return httpx.Response(200, json={'token': 'synthetic-upload-token'})
        assert request.url.path == '/messages'
        scenario['sent_bodies'].append(json.loads(request.content))
        return httpx.Response(200, json={'message': scenario['message']})

    client = MaxClient('synthetic-bot-token', ca_bundle=Settings().ca_file,
                       transport=httpx.MockTransport(max_response), sleeper=lambda _: None)
    env.service = PresentationService(env.store, source_factory=source_factory, client=client,
        clock=lambda: env.now, waiter=lambda _: False)
    subscribe(env)
    assert env.service.tick()['scanned']
    scenario['new'] = True
    scenario['env'] = env
    yield scenario
    client.close()


def test_pdf_source_download_upload_message_confirmation_and_persistent_status(pipeline):
    scenario, env = pipeline, pipeline['env']
    assert advance(env)['sent'] == 1
    assert states(env) == [('123', 1, 'sent')]
    assert len(scenario['uploaded']) == 1 and scenario['pdf'] in scenario['uploaded'][0]
    body = scenario['sent_bodies'][0]
    assert 'PDF: 2 стр.' in body['text']
    assert '04.11.2026 19:35' in body['text']
    assert body['attachments'] == [{'type': 'file', 'payload': {'token': 'synthetic-upload-token'}}]
    # A restarted/status-only instance obtains confirmed delivery from SQLite.
    restored = PresentationService(env.store, clock=lambda: env.now)
    status = restored.status(1)
    assert status['sent'] == 1 and status['last_error'] is None
    assert status['confirmed'] == [{'name': 'Недельный отчёт.pdf', 'sent_at': env.now}]
    assert restored.status(2)['confirmed'] == []
    assert advance(env)['sent'] == 0 and len(scenario['sent_bodies']) == 1
    assert not list(env.store.db_path.parent.glob('max-presentation-*'))


def test_screenshot_upload_certificate_error_does_not_claim_pdf_delivered(pipeline, caplog):
    scenario, env = pipeline, pipeline['env']
    scenario['upload_error'] = True
    assert advance(env)['sent'] == 0
    status = env.service.status(1)
    assert status['error_code'] == 'MAX_UPLOAD_TLS_CERTIFICATE_FAILED'
    assert states(env) == [('123', 1, 'failed')]
    assert status['sent'] == 0 and status['confirmed'] == []
    assert not scenario['sent_bodies']
    assert 'SYNTHETIC_PRIVATE' not in caplog.text + repr(status)
    # No blind replay of a previously failed delivery. Explicit existing retry
    # resumes the saved personal delivery after the certificate issue is fixed.
    scenario['upload_error'] = False
    assert advance(env)['sent'] == 0
    env.service.retry_delivery(1, '101', '201', status['issues'][0]['id'])
    assert env.service.tick()['sent'] == 1
    assert env.service.status(1)['error_code'] is None
    assert len(scenario['sent_bodies']) == 1


@pytest.mark.parametrize('message', [{}, {'body': {}}, {'body': {'mid': ''}}, {'body': {'mid': None}}])
def test_upload_token_alone_is_not_confirmed_delivery(pipeline, message):
    scenario, env = pipeline, pipeline['env']
    scenario['message'] = message
    assert advance(env)['sent'] == 0
    assert states(env) == [('123', 1, 'uncertain')]
    assert env.service.status(1)['confirmed'] == []
    assert env.service.status(1)['error_code'] == 'MAX_RESPONSE_INVALID'
    scenario['message'] = {'body': {'mid': 'later'}}
    assert advance(env)['sent'] == 0 and len(scenario['sent_bodies']) == 1


def test_pdf_source_server_failure_is_not_misreported_as_tls_or_success(pipeline):
    scenario, env = pipeline, pipeline['env']
    scenario['source_status'] = 502
    assert advance(env)['sent'] == 0
    assert env.service.status(1)['error_code'] == 'SOURCE_HTTP_FAILED'
    assert env.service.status(1)['confirmed'] == []
    assert not scenario['uploaded'] and not scenario['sent_bodies']
    scenario['source_status'] = 200
    assert advance(env, 301)['sent'] == 1


def test_pdf_extension_does_not_make_an_error_payload_a_valid_pdf(pipeline):
    scenario, env = pipeline, pipeline['env']
    scenario['pdf'] = b'{"error":"SYNTHETIC_PRIVATE_SOURCE_ERROR"}'
    assert advance(env)['sent'] == 0
    status = env.service.status(1)
    assert states(env) == [('123', 1, 'failed')]
    assert status['error_code'] == 'PRESENTATION_INVALID_PDF'
    assert status['subscribed'] and status['sent'] == 0
    assert not scenario['uploaded'] and not scenario['sent_bodies']
    assert 'SYNTHETIC_PRIVATE' not in repr(status)
