"""Private dialog and polling path against synthetic data, without a real MAX transport."""
import json
from unittest.mock import Mock
import sqlite3

from services.max_bot.polling import DevPoller
from services.max_bot.client import _body
from test.test_max_bot_handler import handler, event, callback, link
from test.test_max_bot_handler_extra import payload_by_label


def test_menu_prompts_for_place_and_reads_current_contacts_without_inventory(handler):
    link(handler)
    handler.stock.published_path.unlink()
    menu=handler.handle(event('/menu'))
    prompt=handler.handle(callback(payload_by_label(menu,'Контакты РСО')))
    assert 'Напишите ОМСУ' in prompt['body']['text']
    assert 'Телефон:' not in prompt['body']['text']
    response=handler.handle(event('Власиха'))
    assert '+7 000 111-22-33' in response['body']['text']
    assert 'PRIVATE_NOTE' not in str(response) and 'example.invalid' not in str(response)
    assert _body(response['body'])==response['body']
    # A successful one-shot directory selection does not trap later stock queries.
    assert 'остатки' in handler.handle(event('Муфта D100'))['body']['text'].lower()


def test_paging_works_without_stocks_and_buttons_recheck_access(handler):
    link(handler)
    data={'items':[{'municipality':'Власиха','rso':'РСО Тест','person':f'Инженер {n:02d}',
                    'phone':f'1111{n:03d}'} for n in range(12)]}
    handler.stock.contacts_path.write_text(json.dumps(data))
    first=handler.handle(event('/contacts Власиха'))
    assert 'Найдено: 12.' in first['body']['text']
    assert 'Инженер 04' in first['body']['text'] and 'Инженер 05' not in first['body']['text']
    token=payload_by_label(first,'Следующие контакты →')
    assert 'Власиха' not in token
    second=handler.handle(callback(token))
    assert 'Инженер 05' in second['body']['text'] and 'Инженер 10' not in second['body']['text']
    with sqlite3.connect(handler.store.accounts_path) as conn:
        conn.execute("UPDATE users SET modules='[]' WHERE id=1")
    rejected=handler.handle(callback(token))
    assert 'Инженер' not in str(rejected) and 'Телефон:' not in str(rejected)


def test_contact_prompt_is_private_and_cleared_by_menu(handler):
    link(handler)
    link(handler,user=2,uid=777,chat=888)
    handler.handle(event('/contacts'))
    outsider=handler.handle(event('Власиха',uid=777,chat=888))
    assert 'Телефон:' not in str(outsider)
    assert handler.handle(event('/contacts Власиха',kind='group')) is None
    handler.handle(event('/menu'))
    assert 'Муфта D100' in handler.handle(event('муфта'))['body']['text']


def test_natural_message_through_poller_without_real_max(handler):
    link(handler)
    handler.stock.published_path.unlink()
    transport=Mock()
    transport.get_me.return_value={'user_id':8811}
    transport.get_updates.return_value={'updates':[event('контакты Власихи',mid='find-contact-mid')],'marker':'1'}
    worker=DevPoller(handler.store,transport,handler)
    try:
        result=worker.poll_once()
    finally:
        worker.close()
    assert result['failed']==0
    transport.send.assert_called_once()
    uid, body=transport.send.call_args.args
    assert uid==123456789 and '+7 000 111-22-33' in body['text']
    assert _body(body)==body


def test_revocation_during_lookup_prevents_phone_reply(handler,monkeypatch):
    link(handler)
    original=handler.stock.search_contacts
    def search(*args,**kwargs):
        result=original(*args,**kwargs)
        handler.store.unlink_max(123456789,321)
        return result
    monkeypatch.setattr(handler.stock,'search_contacts',search)
    response=handler.handle(event('контакты Власихи'))
    assert '+7 000' not in str(response) and 'Подключить MAX' in response['body']['text']

def test_item_does_not_guess_other_organization_and_offers_explicit_city_search(handler):
    from test.test_max_bot_handler import button
    link(handler)
    handler.stock.contacts_path.write_text(json.dumps({'items':[{'rso':'Другая организация','municipality':'Власиха',
                'person':'Дежурный','phone':'333-555'}]}))
    card=handler.handle(callback(button(handler.handle(event('муфта')))))
    assert '333-555' not in str(card)
    opened=handler.handle(callback(payload_by_label(card,'Искать в справочнике по ОМСУ')))
    assert '333-555' in opened['body']['text'] and 'Другая организация' in opened['body']['text']

def test_bare_directory_request_asks_for_criteria(handler):
    link(handler)
    question=handler.handle(event('нужен телефон ответственного'))
    assert 'Напишите ОМСУ' in question['body']['text'] and '+7 000' not in str(question)
    assert '+7 000' in handler.handle(event('Власиха'))['body']['text']
