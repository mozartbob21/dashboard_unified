import json
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from services.water_dashboard import insights as ai, builder


def source(value=10,previous=12,*,mid='overdue_tasks',unit='шт.'):
    result={'ok':True,'metric_schema':1,'updated_at':'2026-10-26T12:00:00','data_date':'2026-10',
            'metrics':[{'id':mid,'label':'untrusted label secret','value':value,'unit':unit}],
            'tables':[{'rows':[['Secret person','Secret municipality']]}], 'text':'private raw text','url':'https://private.invalid'}
    if previous is not None:
        result['previous_observation']={'metric_schema':1,'updated_at':'2026-10-25T12:00:00','data_date':'2026-10',
             'metrics':[{'id':mid,'value':previous,'unit':unit}]}
    return result


def reply(facts):
    return json.dumps([{'metric_id':f['metric_id'],'text':'Для интерпретации показателя следует уточнить полноту сведений и сопоставимость условий наблюдения.'} for f in facts],ensure_ascii=False)


@pytest.fixture
def service(tmp_path):
    now=[1800000000.0]; work=[]; calls=[]
    def model(facts): calls.append(facts); return reply(facts)
    instance=ai.InsightService(tmp_path/'jobs.sqlite',clock=lambda:now[0],model=model,submit=lambda *args:work.append(args))
    return instance,now,work,calls


def complete(work):
    function,*args=work.pop(0); function(*args)


def test_get_never_starts_post_single_flight_and_restart_uses_durable_cache(service):
    instance,now,work,calls=service; data=source()
    assert instance.request('tasks',data)['status']=='unavailable' and not work
    a=instance.request('tasks',data,start=True)
    assert a['status']=='pending' and a['source_updated_at']==data['updated_at']
    assert instance.request('tasks',data,start=True)==a and len(work)==1
    complete(work)
    b=instance.request('tasks',data)
    assert b['status']=='ready' and b['items'][0]['trend']=='improved'
    assert '10 шт.' in b['items'][0]['text'] and '12 шт.' in b['items'][0]['text']
    reopened=ai.InsightService(instance.path,clock=lambda:now[0],model=lambda *_:pytest.fail('no repeated model'))
    assert reopened.request('tasks',data,start=True)==b
    payload=json.dumps(calls,ensure_ascii=False)
    assert not any(secret in payload for secret in ('Secret','private','untrusted'))


def test_concurrent_requests_launch_exactly_one_job(service):
    instance,_,work,_=service
    with ThreadPoolExecutor(max_workers=8) as pool:
        results=list(pool.map(lambda _:instance.request('tasks',source(),start=True),range(16)))
    assert {r['status'] for r in results}=={'pending'} and len(work)==1


def test_global_capacity_and_expired_worker_cannot_overwrite_new_job(service):
    instance,now,work,_=service
    instance.request('tasks',source(),start=True)
    second=source(20);instance.request('tasks',second,start=True)
    assert instance.request('tasks',source(30),start=True)['status']=='unavailable'
    old=work.pop(0)
    now[0]+=ai.LEASE_SECONDS+1
    assert instance.request('tasks',source())['status']=='unavailable'
    assert instance.request('tasks',source(),start=True)['status']=='pending'
    old[0](*old[1:])
    assert instance.request('tasks',source())['status']=='pending'


@pytest.mark.parametrize('before,after,trend',[(12,10,'improved'),(10,12,'worsened'),(0,0,'stable'),(None,0,'unknown')])
def test_trend_direction_zero_and_missing(before,after,trend):
    assert ai.facts_for('tasks',source(after,before))[0]['trend']==trend


def test_no_comparison_for_changed_period_unit_or_invalid_values_and_neutral_meetings():
    data=source();data['data_date']='2026-11'
    assert not ai.facts_for('tasks',data)[0]['comparable']
    data=source();data['previous_observation']['metrics'][0]['unit']='%'
    assert ai.facts_for('tasks',data)[0]['trend']=='unknown'
    for value in (True,float('nan'),float('inf'),-2,'17'):
        assert ai.facts_for('tasks',source(value))[0]['value'] is None
    facts=ai.facts_for('meetings',source(12,10,mid='meetings'))
    assert next(f for f in facts if f['metric_id']=='meetings')['trend']=='unknown'
    data=source(90,80,mid='attendance_pct',unit='%')
    data['tables']=[{'headers':['ОМСУ','Явка, %'],'rows':[['Test',90]]}]
    data['previous_observation']['comparison_scopes']=ai.comparison_scopes(data)
    facts=ai.facts_for('meetings',data)
    assert next(f for f in facts if f['metric_id']=='attendance_pct')['trend']=='improved'


def test_changed_snapshot_get_does_not_return_previous_revision(service):
    instance,_,work,_=service
    first=instance.request('tasks',source(),start=True);complete(work)
    other=source(9);other['updated_at']='2026-10-26T13:00:00'
    result=instance.request('tasks',other)
    assert result['status']=='unavailable' and result['revision']!=first['revision']


@pytest.mark.parametrize('raw', ['not json','[]','[{"metric_id":"overdue_tasks","text":"Стало лучше на 100% благодаря ремонту"}]',
    '[{"metric_id":"overdue_tasks","text":"<script>alert(1)</script>"}]'])
def test_invalid_or_invented_model_output_is_truthfully_unavailable(service,raw):
    instance,_,work,_=service;instance.model=lambda _:raw
    instance.request('tasks',source(),start=True);complete(work)
    result=instance.request('tasks',source())
    assert result['status']=='unavailable' and result['items']==[]
    assert '100%' not in str(result) and '<script>' not in str(result)


def test_model_error_does_not_leak_and_retry_is_bounded(service):
    instance,now,work,_=service
    def fail(_): raise RuntimeError('SECRET-TOKEN AND PRIVATE URL')
    instance.model=fail;instance.request('tasks',source(),start=True);complete(work)
    assert 'SECRET' not in str(instance.request('tasks',source(),start=True))
    assert not work
    now[0]+=61
    assert instance.request('tasks',source(),start=True)['status']=='pending' and len(work)==1


def test_failed_refresh_does_not_replace_prior_observation(tmp_path,monkeypatch):
    monkeypatch.setattr(builder,'SNAPSHOT_FILE',tmp_path/'snapshot.json')
    def collection(n): return {'tasks':{'data_date':'2026-10','tables':[{'headers':['ОМСУ','Просроченные задачи'],'rows':[['Test',str(n)]]}]}}
    one=builder.build_snapshot(collection(12),['tasks'])
    assert 'previous_observation' not in one['sources']['tasks']
    two=builder.build_snapshot(collection(10),['tasks'])
    prior=two['sources']['tasks']['previous_observation']
    assert prior['metrics'][0]['value']==12 and 'tables' not in prior
    failed=builder.build_snapshot({'tasks':{'error':'Failure'}},['tasks'])
    assert failed['sources']['tasks']['previous_observation']==prior
    three=builder.build_snapshot(collection(8),['tasks'])
    assert three['sources']['tasks']['previous_observation']['metrics'][0]['value']==10
    assert ai.facts_for('tasks',three['sources']['tasks'])[0]['trend']=='improved'


def test_authenticated_api_rechecks_permissions_and_protects_get_and_post(tmp_path,monkeypatch,service):
    from routers import water_insights as router
    from services.auth import security
    from test.test_zip_cloud_integration import isolated_db
    current={'id':1,'is_active':True,'modules':['water-dashboard']}
    monkeypatch.setattr(router,'get_user_from_token',lambda token:dict(current) if token=='signed-session' else None)
    monkeypatch.setattr(router,'service',lambda:service[0])
    path=tmp_path/'snapshot.json';path.write_text(json.dumps({'sources':{'tasks':source()}}))
    monkeypatch.setattr(router,'SNAPSHOT_FILE',path)
    app=FastAPI();app.include_router(router.router)
    with TestClient(app) as client:
        route='/water-dashboard/insights/tasks'
        assert client.get(route).status_code==401 and client.post(route).status_code==401
        client.cookies.set('access_token','signed-session')
        assert client.get(route).json()['status']=='unavailable'
        assert client.post(route,headers={'Origin':'https://attacker.invalid'}).status_code==403
        assert client.post(route).json()['status']=='pending'
        current['modules']=[]
        assert client.get(route).status_code==403 and client.post(route).status_code==403
        current['is_active']=False
        assert client.get(route).status_code==401


def test_model_adapter_only_receives_aggregate_facts(monkeypatch):
    from services.summarizer import engine
    captured=[]
    def call(messages,**kwargs): captured.append((messages,kwargs));return reply(ai.facts_for('tasks',source()))
    monkeypatch.setattr(engine,'_qwen_chat',call)
    ai.model_commentary(ai.facts_for('tasks',source()))
    messages,options=captured[0]
    assert options['module_id']=='water-dashboard'
    assert not any(word in json.dumps(messages) for word in ['Secret','private.invalid','untrusted label'])


def test_all_six_stages_have_neutral_facts_and_week_rollover_is_unknown():
    data=source(4,6,mid='stage_1_pending',unit='РСО / подразделений')
    facts=ai.facts_for('gkh_kontur',data)
    assert len([f for f in facts if f['metric_id'].startswith('stage_')])==6
    assert next(f for f in facts if f['metric_id']=='stage_1_pending')['trend']=='unknown'
    data=source(90,80,mid='electronic_documents_pct',unit='%')
    data['previous_observation']['updated_at']='2026-10-25T12:00:00'
    assert next(f for f in ai.facts_for('edo_rso',data) if f['metric_id']=='electronic_documents_pct')['trend']=='unknown'


def test_attendance_different_population_is_not_compared():
    data=source(90,80,mid='attendance_pct',unit='%')
    data['tables']=[{'headers':['ОМСУ','Явка, %'],'rows':[['Old',90]]}]
    data['previous_observation']['comparison_scopes']=ai.comparison_scopes(data)
    data['tables'][0]['rows'].append(['New',70])
    assert ai.facts_for('meetings',data)[0]['trend']=='unknown'


def test_cache_is_bounded_and_latest_result_is_retained(service,monkeypatch):
    instance,now,work,_=service
    monkeypatch.setattr(ai,'MAX_CACHE_ROWS',3)
    for value in range(8):
        now[0]+=1
        instance.request('tasks',source(value),start=True)
        complete(work)
    conn=instance._connect()
    try:
        assert conn.execute('SELECT count(*) FROM insight_jobs').fetchone()[0]==3
    finally:
        conn.close()
    assert instance.request('tasks',source(7))['status']=='ready'
    assert instance.request('tasks',source(0))['status']=='unavailable'


def test_genuine_metric_specific_model_commentary_is_preserved():
    facts=ai.facts_for('tasks',source())
    prose='Для контроля просроченных задач стоит уточнить сроки исполнения и проверить полноту сведений о закрытии.'
    items=ai.validated_items(json.dumps([{'metric_id':'overdue_tasks','text':prose}],ensure_ascii=False),facts)
    assert items[0]['text'].endswith(prose)
    assert items[0]['trend']=='improved' and '10 шт.' in items[0]['text']


@pytest.mark.parametrize('prose',[
    'Количество снизилось благодаря проведённой работе ответственных специалистов.',
    'Показатель стал лучше на 20 процентов после последнего обновления.',
    'Подробности доступны по ссылке https://example.invalid/report.',
    'Сведения улучшились, поэтому ситуация стабильна и благополучна.',
])
def test_obvious_unsupported_model_claims_are_rejected(prose):
    facts=ai.facts_for('tasks',source(previous=None))
    with pytest.raises(ValueError):
        ai.validated_items(json.dumps([{'metric_id':'overdue_tasks','text':prose}],ensure_ascii=False),facts)


def test_commentary_cannot_replace_metric_ids_or_include_extraneous_fields():
    facts=ai.facts_for('tasks',source())
    for item in ({'metric_id':'unknown','text':'Проверьте полноту сведений для дальнейшей оценки.'},
                 {'metric_id':'overdue_tasks','text':'Проверьте полноту сведений для дальнейшей оценки.','trend':'worsened'}):
        with pytest.raises(ValueError): ai.validated_items(json.dumps([item]),facts)
