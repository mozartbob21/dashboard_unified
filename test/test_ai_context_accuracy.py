"""Synthetic reporting regressions: no model, portals, credentials or real DB."""
from datetime import date, datetime, timezone
import gzip
import json
from pathlib import Path

import httpx
import pytest

from services.aichat import report_context as reports, report_periods, collective_source
from services import water_ai_context as water, report_municipalities
from services.collective import store
from utils import db

REAL_TODAY = report_periods.today
MIGRATION = Path(db.__file__).resolve().parents[1] / 'migrations/20261019_collective_workspace.sql'


@pytest.fixture(autouse=True)
def local_data(monkeypatch, tmp_path):
    monkeypatch.setattr(reports, 'DATA_DIR', tmp_path)
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'dashboard.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(store, 'MIGRATION', MIGRATION)
    monkeypatch.setattr(store, 'now', lambda: datetime(2026,10,19,8,tzinfo=timezone.utc))
    monkeypatch.setattr(report_periods, 'today', lambda: date(2026,10,19))
    monkeypatch.setattr(water, 'load_snapshot', lambda: {'schema_version':3,'sources':{},'table':[]})
    monkeypatch.setattr(httpx.HTTPTransport, 'handle_request', lambda *a, **kw: pytest.fail('Live network forbidden'))
    registry = tmp_path / 'municipality_registry.json'
    registry.write_text(json.dumps({'municipalities':{'Власиха':{},'Химки':{}}}))
    monkeypatch.setattr(report_municipalities, 'REGISTRY', registry)
    return tmp_path


def write(root, path, value):
    path = root / path
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(value,ensure_ascii=False),encoding='utf-8')


def record(primary='SYNTH-1', municipality='Власиха'):
    row=['']*16
    row[0],row[1],row[2],row[5],row[8],row[10],row[11],row[12],row[14] = ('МСЭД','18.10.2026',municipality,primary,
        '<p>Перебои водоснабжения.</p>','12','PRIVATE_ANSWER','PRIVATE_EMAIL','В работе')
    return row


def collective_data():
    data=store.save_snapshot('2026-10-01','2026-10-19',[],[record(),record('SYNTH-2','Химки')])
    assignee=store.create_assignee('Синтетический ответственный','local:1')
    store.update_workflow(data['keys'][0],0,{'assignee_id':assignee['id'],'report':'Проведено обследование; срок 20 октября.','theme':'ii'},'local:1')
    return data


def test_collective_exact_module_not_generic_dashboard(local_data):
    collective_data()
    bundle=reports.prepare('Дай отчёт по коллективным жалобам МинЖКХ',['collective','mingkh','edds'])
    assert bundle['selection']['modules']==['collective']
    assert {item['module'] for item in bundle['sources']}=={'collective'}
    source=bundle['sources'][0]
    assert source['metrics'][0]['value']==2
    assert source['metrics'][1]['value']==1
    assert any(row['Ручной доклад'].startswith('Проведено обследование') for row in source['rows'])
    assert source['rows'][0]['Ответственный']=='Синтетический ответственный'
    assert source['rows'][0]['Аннотация']=='Перебои водоснабжения.'
    assert source['rows'][0]['Подписей']==12
    assert source['metrics'][-1]['value']==24
    assert source['period']=={'from':'2026-10-01','to':'2026-10-19'}
    serialized=reports.to_prompt(bundle)
    assert all(value not in serialized for value in ('PRIVATE_', 'local:1','report_by','assigned_by'))


def test_collective_denied_grant_never_reads_sql_or_other_module(monkeypatch):
    monkeypatch.setattr(collective_source,'_saved',lambda *a,**kw:pytest.fail('Denied collective data read'))
    bundle=reports.prepare('Дай отчёт по коллективным жалобам',['mingkh'])
    assert bundle['clarification'] and not bundle['sources']


def test_collective_municipal_filter_no_region_counts(local_data):
    collective_data()
    bundle=reports.prepare('Отчёт по коллективным жалобам по Власихе',['collective'])
    assert bundle['selection']['municipality']=='Власиха'
    source=bundle['sources'][0]
    assert source['metrics'][0]['value']==1
    assert all(row['Муниципалитет']=='Власиха' for row in source['rows'])
    assert 'Химки' not in reports.to_prompt(bundle)


def test_missing_collective_does_not_create_database(local_data):
    source=collective_source.reports_for_module(local_data)[0]
    assert source['status']=='missing' and source['metrics']==[]
    assert not db.DB_FILE.exists()


def test_synthetic_context_root_never_reads_different_database(local_data,monkeypatch):
    monkeypatch.setattr(db,'DB_FILE',local_data/'other'/'dashboard.db')
    monkeypatch.setattr(store,'latest_snapshot',lambda:pytest.fail('Different DB was read'))
    assert collective_source.reports_for_module(local_data)[0]['status']=='missing'


@pytest.mark.parametrize('question',['Дай отчёт по СМС','Покажи показатели SMS МинЖКХ','Сводка СМС за вчера'])
def test_sms_does_not_silently_select_all_modules(question):
    assert reports.wants_context(question,['mingkh','edds'])
    result=reports.prepare(question,['mingkh','edds'])
    assert result['clarification'] and not result['sources']
    assert 'СМС' in result['clarification']


@pytest.mark.parametrize('module,question,path',[
    ('edo','Дай отчёт по заполненности данных','edo/result.json'),
    ('watercontrol','Дай отчёт по контролю воды','watercontrol/result.json')])
def test_failed_empty_run_is_not_successful_zero(local_data,module,question,path):
    write(local_data,path,{'created_at':'2026-10-19','summary_message':'Проверка не завершена.',
        'extraction_note':'SECRET_ERROR_RAW','rows':[]})
    source=reports.prepare(question,[module])['sources'][0]
    assert source['status']=='missing' and not source['metrics'] and source['matching_rows'] is None
    assert 'не завершена' in source['warning']
    assert 'SECRET_ERROR_RAW' not in reports.to_prompt({'sources':[source]})


def test_valid_empty_run_keeps_confirmed_zero(local_data):
    write(local_data,'edo/result.json',{'created_at':'2026-10-19','summary_message':'Проверка завершена.','rows':[]})
    source=reports.prepare('Заполненность данных',['edo'])['sources'][0]
    assert source['status']=='saved' and source['metrics'][0]['value']==0


@pytest.mark.parametrize('question,expected',[
    ('Дай отчёт по ЕДДС за вчера',('2026-10-18','2026-10-18')),
    ('Дай отчёт по ЕДДС сегодня',('2026-10-19','2026-10-19')),
    ('Дай отчёт по ЕДДС с 01.10.2026 по 18.10.2026',('2026-10-01','2026-10-18')),
    ('Дай отчёт по ЕДДС за 2026-10-17 — 2026-10-18',('2026-10-17','2026-10-18'))])
def test_request_dates_are_resolved_and_stored(question,expected):
    selection=reports.resolve_scope(question,['edds'])
    period=selection['requested_period']
    assert (period['from'],period['to'])==expected and period['timezone']=='Europe/Moscow'


def test_moscow_date_rolls_forward_at_21utc(monkeypatch):
    class Clock(datetime):
        @classmethod
        def now(cls,tz=None):
            return datetime(2026,10,18,22,tzinfo=timezone.utc).astimezone(tz)
    monkeypatch.setattr(report_periods,'today',REAL_TODAY)
    monkeypatch.setattr(report_periods,'datetime',Clock)
    assert report_periods.today()==date(2026,10,19)


@pytest.mark.parametrize('question',['Отчёт по ЕДДС с 31.09.2026 по 18.10.2026','Отчёт по ЕДДС за 2026-11-01',
    'Отчёт по ЕДДС с 18.10.2026 по 01.10.2026','Отчёт по ЕДДС за октябрь','Сравни ЕДДС 01.10.2026 и 18.10.2026'])
def test_invalid_or_ambiguous_period_requires_clarification(question):
    bundle=reports.prepare(question,['edds'])
    assert bundle['clarification'] and bundle['sources']==[]


def test_daily_edds_filters_yesterday_without_whole_history(local_data):
    write(local_data,'edds/water_daily.json',{'updated':'2026-10-19','days':{
        '2026-10-17':{'Власиха':[1000,2000,3000]},'2026-10-18':{'Власиха':[1,2,3]},'2026-10-19':{'Власиха':[99,99,99]}}})
    bundle=reports.prepare('Дай отчёт по ЕДДС за вчера',['edds'])
    assert not bundle.get('clarification')
    source=next(s for s in bundle['sources'] if s['id']=='edds-complaints')
    assert [m['value'] for m in source['metrics']]==[1,2,3]
    assert source['period_selection']=='exact'
    assert source['period']['from']==source['period']['to']=='2026-10-18'
    assert all(not s['metrics'] for s in bundle['sources'] if s['id']!='edds-complaints')


def test_daily_gaps_are_not_counted_as_zero(local_data):
    write(local_data,'edds/water_daily.json',{'updated':'2026-10-19','days':{
        '2026-10-17':{'Власиха':[1,2,3]},'2026-10-19':{'Власиха':[1,2,3]}}})
    bundle=reports.prepare('Отчёт по ЕДДС с 17.10.2026 по 19.10.2026',['edds'])
    assert bundle['clarification'] and bundle['sources']==[]


def test_daily_missing_city_on_one_day_does_not_become_zero(local_data):
    write(local_data,'edds/water_daily.json',{'updated':'2026-10-19','days':{
        '2026-10-17':{'Власиха':[1,2,3]},'2026-10-18':{'Химки':[10,20,30]}}})
    bundle=reports.prepare('Отчёт ЕДДС по Власихе с 17.10.2026 по 18.10.2026',['edds'])
    assert bundle['clarification'] and bundle['sources']==[]


def test_collection_timestamp_cannot_substitute_for_data_period(local_data):
    write(local_data,'cameras/state/dashboard_state.json',{'updated_at':'2026-10-18','rows':[{'municipality':'Власиха','camera_status':'working'}]})
    bundle=reports.prepare('Отчёт по камерам за вчера',['cameras'])
    assert bundle['clarification'] and not bundle['sources']


def test_collective_selects_exact_persisted_period_and_excludes_other_period(local_data):
    collective_data()
    exact=reports.prepare('Отчёт по коллективным жалобам с 01.10.2026 по 19.10.2026',['collective'])
    assert exact['sources'][0]['metrics'][0]['value']==2
    mismatch=reports.prepare('Отчёт по коллективным жалобам за вчера',['collective'])
    assert mismatch['clarification'] and not mismatch['sources']


def test_mingkh_period_selects_only_matching_comparison_period(local_data):
    write(local_data,'mingkh/report-summary.json',{'schema_version':1,'collected_at':'2026-10-19','periods':{
        'curr':{'bounds':['2026-10-12','2026-10-18'],'total':200,'municipalities':{}},
        'prev':{'bounds':['2026-10-05','2026-10-11'],'total':50,'municipalities':{}}}})
    bundle=reports.prepare('Отчёт МинЖКХ с 05.10.2026 по 11.10.2026',['mingkh'])
    exact=next(s for s in bundle['sources'] if s['id']=='mingkh-appeals')
    assert [m['value'] for m in exact['metrics']]==[50]
    assert list(exact['periods'])==['prev']


def test_followup_period_is_retained_and_can_be_cleared(local_data):
    first=reports.resolve_scope('Отчёт по ЕДДС за вчера',['edds'])
    assert reports.wants_context('А за сегодня?',['edds'],previous=first)
    followup=reports.resolve_scope('А теперь по Власихе',['edds'],previous=first)
    assert followup['requested_period']==first['requested_period']
    clear=reports.resolve_scope('Последний сохранённый срез',['edds'],previous=followup)
    assert 'requested_period' not in clear and clear['modules']==['edds']


@pytest.mark.parametrize('question',['Привет!','Напиши стих про коллективные жалобы','Что такое коллективное обращение?',
    'Как отправить СМС?','Расскажи об истории Власихи','Переведи слово today'])
def test_ordinary_chat_still_does_not_load_operational_data(question):
    assert not reports.wants_context(question,list(reports.MODULES))


def test_grouped_collective_report_reuses_one_data_snapshot(local_data, monkeypatch):
    collective_data()
    calls=[]
    original=collective_source._saved
    def tracked(*args,**kwargs):
        calls.append(1)
        return original(*args,**kwargs)
    monkeypatch.setattr(collective_source,'_saved',tracked)
    bundle=reports.prepare('Отчёт по коллективным жалобам по каждому муниципалитету',['collective'])
    assert len(calls)<=3  # catalogue, optional municipal catalogue, one shared data snapshot
    table=bundle['sources'][0]['municipality_table']
    assert len(table['rows'])==2


@pytest.mark.parametrize('value,expected',[(None,None),('',None),('—',None),('-2',None),('0',0),(0,0),
    (' 1 234 подписей',1234),('+12',12),('9007199254740992',None),(True,None)])
def test_collective_signature_missing_is_not_zero(value,expected):
    assert collective_source._signatures(value)==expected


def test_collective_signatures_total_unknown_and_safe_annotation(local_data):
    first,second=record(),record('SYNTH-2')
    first[8]='<script>HIDDEN_SCRIPT</script><p>'+('Описание '*500)+'</p>'
    first[10],second[10]='0','нет данных'
    store.save_snapshot('2026-10-01','2026-10-19',[],[first,second])
    source=reports.prepare('Отчёт по коллективным жалобам',['collective'])['sources'][0]
    assert source['metrics'][-1]['value'] is None
    assert source['signatures_coverage']=={'known':1,'unknown':1,'total_is_complete':False}
    row=source['rows'][0]
    assert row['Подписей']==0 and row['Аннотация сокращена']
    assert len(row['Аннотация'])<=collective_source.MAX_ANNOTATION_CHARS
    assert 'HIDDEN_SCRIPT' not in row['Аннотация'] and '<p>' not in row['Аннотация']


def test_collective_finds_municipality_only_in_requested_older_snapshot(local_data):
    store.save_snapshot('2026-10-01','2026-10-10',[],[record(municipality='Реутов')])
    store.save_snapshot('2026-10-17','2026-10-19',[],[record('NEW-1')])
    question='Отчёт по коллективным жалобам по Реутову с 01.10.2026 по 10.10.2026'
    bundle=reports.prepare(question,['collective'])
    assert not bundle.get('clarification')
    assert bundle['selection']['municipality']=='Реутов'
    assert bundle['sources'][0]['metrics'][0]['value']==1
    first=reports.resolve_scope('Отчёт по коллективным жалобам с 01.10.2026 по 10.10.2026',['collective'])
    assert reports.wants_context('А теперь по Реутову',['collective'],previous=first)
    followup=reports.resolve_scope('А теперь по Реутову',['collective'],previous=first)
    assert followup['municipality']=='Реутов' and followup['requested_period']==first['requested_period']


@pytest.mark.parametrize('question',[
    'Проанализируй коллективную жалобу: «письмо от 10.10.2026, отчет за вчера»',
    'Коллективные жалобы: "номер 2026-10-10 и срок 2026-10-17"',
    'Отчёт по ЕДДС\n> В источнике проверка от 10.10.2026',
    'Дай отчёт по ЕДДС. Пример: `за 10.10.2026`',
    'Дай отчёт по ЕДДС\n```\nза 10.10.2026\n```'])
def test_dates_inside_source_quotations_are_not_request_periods(question):
    assert report_periods.extract(question)==(None,None)


def test_real_period_ignores_quoted_dates():
    period,error=report_periods.extract('Отчёт по коллективным жалобам за вчера, включая «письмо от 10.10.2026»')
    assert not error and period['from']==period['to']=='2026-10-18'


@pytest.mark.parametrize('today,question,expected',[
    ('2026-01-01','за 7 дней',('2025-12-26','2026-01-01')),
    ('2026-01-01','за неделю',('2025-12-26','2026-01-01')),
    ('2026-03-01','за последние 30 дней',('2026-01-31','2026-03-01')),
    ('2024-03-01','за 30 дней',('2024-02-01','2024-03-01')),
    ('2026-01-01','за текущий месяц',('2026-01-01','2026-01-01')),
    ('2026-01-01','за прошлый месяц',('2025-12-01','2025-12-31')),
    ('2024-03-12','за предыдущий месяц',('2024-02-01','2024-02-29')),
    ('2026-10-19','за текущую неделю',('2026-10-19','2026-10-19')),
    ('2026-10-19','за прошлую неделю',('2026-10-12','2026-10-18')),
    ('2026-01-01','за прошлую неделю',('2025-12-22','2025-12-28')),
    ('2026-01-01','за эту неделю',('2025-12-29','2026-01-01')),
    ('2026-01-01','за неделю включая сегодня',('2025-12-26','2026-01-01')),
])
def test_common_periods_calendar_boundaries(monkeypatch,today,question,expected):
    monkeypatch.setattr(report_periods,'today',lambda:date.fromisoformat(today))
    period,error=report_periods.extract('Отчёт по ЕДДС '+question)
    assert not error
    assert (period['from'],period['to'])==expected and period['label'] and period['timezone']=='Europe/Moscow'


@pytest.mark.parametrize('question',['Отчёт за 0 дней','Отчёт за 999999999 дней',
    'Отчёт за неделю и прошлый месяц','Отчёт за неделю и вчера','Отчёт за неделю и 01.10.2026'])
def test_incompatible_or_unbounded_periods_ask_clarification(question):
    assert report_periods.extract(question)[1]


def test_mixed_request_statistics_and_instructions_is_operational():
    assert reports.wants_context('Покажи статистику коллективных жалоб и объясни, как сохранить доклад',['collective'])
    assert reports.wants_context('Сформируй краткий отчёт по ЕДДС и объясни показатели',['edds'])
    assert not reports.wants_context('Дай инструкцию как пользоваться блоком коллективных жалоб',['collective'])


@pytest.mark.parametrize('phrase',['за завтра','на послезавтра','за «завтра»'])
def test_future_relative_period_does_not_return_saved_snapshot(local_data,phrase):
    write(local_data,'edo/result.json',{'created_at':'2026-10-19','rows':[{'municipality':'Власиха','value':999}]})
    bundle=reports.prepare('Отчёт по заполненности данных '+phrase,['edo'])
    assert bundle['clarification'] and bundle['sources']==[]
    assert 'Будущий период' in bundle['clarification']


@pytest.mark.parametrize('phrase,expected',[
    ('за «19.10.2026»',('2026-10-19','2026-10-19')),
    ('за "2026-10-18"',('2026-10-18','2026-10-18')),
    ('за “19.10.2026”',('2026-10-19','2026-10-19')),
    ('с «10.10.2026» по «19.10.2026»',('2026-10-10','2026-10-19')),
    ('за «10.10.2026 — 19.10.2026»',('2026-10-10','2026-10-19')),
])
def test_explicit_quoted_date_is_a_period(phrase,expected):
    period,error=report_periods.extract('Отчёт по ЕДДС '+phrase)
    assert not error and (period['from'],period['to'])==expected


@pytest.mark.parametrize('quoted',[
    '«письмо от 10.10.2026, ответить завтра»',
    '`за «19.10.2026» и завтра`',
    '\n> за «19.10.2026» и завтра',
    '\n```\nза «19.10.2026» и завтра\n```',
])
def test_quoted_future_and_dates_in_source_text_or_code_are_ignored(quoted):
    assert report_periods.extract('Отчёт по ЕДДС. Текст: '+quoted)==(None,None)


def test_explicit_file_report_does_not_get_replaced_by_platform_data():
    question = 'Составь отчёт по коллективным жалобам из приложенного файла'
    assert not reports.wants_context(question, ['collective'], has_files=True)
    assert reports.wants_context('Сравни файл с показателями коллективных жалоб', ['collective'], has_files=True)
