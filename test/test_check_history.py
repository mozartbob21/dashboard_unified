"""Real SQLite and process leases; all inputs and result artifacts are synthetic."""
from datetime import date, datetime, timedelta, timezone
import json
from pathlib import Path
import subprocess
import sys

from fastapi import HTTPException
import pytest

from services import check_history as history, check_metrics, run_history
from utils import db

USER = {'id':5,'modules':['edo','mingkh','cameras','cds','water-dashboard','tools','utnkr']}
OTHER = {**USER,'id':8}


@pytest.fixture(autouse=True)
def isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'history.db')
    monkeypatch.setattr(db,'_SCHEMA_READY',False)
    monkeypatch.setattr(check_metrics,'BASE',tmp_path)
    yield
    for ident in list(history._guards):
        history.release(ident)


def completed(value,module='edo',params=None,status='success',private=None,metric='issues'):
    ident = run_history.record_start(module,params=params,private_owner=private)
    run_history.record_finish(ident,status,'secret diagnostics',metrics={metric:value})
    return ident


def set_date(ident,stamp):
    with db.get_db_connection() as conn:
        conn.execute('UPDATE run_history SET started_at=?,finished_at=? WHERE run_id=?',(stamp,stamp,ident))


def test_fresh_artifact_zero_is_preserved_without_rows_or_secrets(tmp_path):
    target = tmp_path/check_metrics.PATHS['edo']
    target.parent.mkdir(parents=True)
    target.write_text(json.dumps({'rows':[{'token':'old secret'}],'missing_data_issues':[{}]}))
    ident = run_history.record_start('edo',user='sensitive username')
    target.write_text(json.dumps({'rows':[],'missing_data_issues':[], 'source_url':'internal-token',
                                 'personal_messages':[{'password':'secret_password'}]}))
    run_history.record_finish(ident)
    record, = history.read_checks(USER,'edo')
    assert record['metrics'] == {'rows':0,'issues':0} and record['result_code'] == 'measured'
    with db.get_db_connection() as conn:
        stored = '\n'.join(conn.iterdump())
    for secret in ('secret_password','internal-token','old secret','sensitive username'):
        assert secret not in stored


def test_no_artifact_update_unsupported_param_and_nonfinite_value_fail_closed(tmp_path):
    target = tmp_path/check_metrics.PATHS['edo']
    target.parent.mkdir(parents=True)
    target.write_text(json.dumps({'rows':[],'missing_data_issues':[]}))
    ident = run_history.record_start('edo',params={'api_key':'secret_api_key'})
    run_history.record_finish(ident)
    first, = history.read_checks(USER,'edo')
    assert first['metrics'] == {} and first['result_code'] == 'unconfirmed'
    assert first['params'] == {'unknown':True}
    ident = run_history.record_start('edo')
    run_history.record_finish(ident,metrics={'issues':float('nan'),'secret_metric':42,'rows':False})
    assert history.read_checks(USER,'edo')[0]['metrics'] == {}


def test_restart_reads_completed_and_recovers_dead_executor(tmp_path):
    saved = completed(2)
    # A separate interpreter acquires and then loses its real OS lock. Its
    # unfinished row remains; a new reader can prove the executor disappeared.
    script = """import sys,os
from pathlib import Path
from utils import db
db.DB_FILE=Path(sys.argv[1])
from services import run_history
ident=run_history.record_start('edo')
print(ident,flush=True)
os._exit(2)"""
    child = subprocess.run([sys.executable,'-c',script,str(db.DB_FILE)],capture_output=True,text=True,timeout=12)
    assert child.returncode == 2, child.stderr
    orphan = child.stdout.strip().splitlines()[-1]
    results = {r['run_id']:r for r in history.read_checks(USER,'edo')}
    assert results[saved]['metrics'] == {'issues':2}
    assert results[orphan]['status'] == 'interrupted' and results[orphan]['metrics'] == {}
    run_history.record_finish(orphan,'success',metrics={'issues':8})
    assert history.read_checks(USER,'edo')[0]['status'] == 'interrupted'


def test_live_run_is_preserved_and_terminal_finish_is_immutable():
    ident = run_history.record_start('edo')
    assert history.read_checks(USER,'edo')[0]['status'] == 'running'
    run_history.record_finish(ident,'error','private-url?token=key')
    run_history.record_finish(ident,'success',metrics={'issues':7})
    row, = history.read_checks(USER,'edo')
    assert row['status'] == 'error' and row['metrics'] == {}
    assert 'private-url' not in repr(run_history.get_recent())


def test_pre_migration_rows_have_no_fabricated_indicators(tmp_path):
    with db.get_db_connection() as conn:
        conn.execute("INSERT INTO run_history(run_id,module_id,user,started_at,status) VALUES ('legacy','edo','old','2025-01-01','success')")
    completed(0)
    results = history.read_checks(USER,'edo')
    assert results[-1]['version'] == 0 and results[-1]['metrics'] == {}
    assert history.compare(USER,'edo')['status'] == 'unavailable'
    assert history.read_checks(USER,'edo',start='2025-01-01',end='2025-01-02')[0]['run_id'] == 'legacy'


def test_permissions_private_runs_and_public_chat_never_leak():
    shared = completed(1,'tools',metric='files')
    own = completed(4,'tools',private='local:5',metric='files')
    theirs = completed(6,'tools',private='local:8',metric='files')
    assert {row['run_id'] for row in history.read_checks(USER,'tools')} == {own,shared}
    assert {row['run_id'] for row in history.read_checks(OTHER,'tools')} == {theirs,shared}
    assert [r['run_id'] for r in history.read_checks(USER,'tools',public_only=True)] == [shared]
    assert [r['run_id'] for r in run_history.get_all()] == [shared]
    assert history.compare(USER,'tools',public_only=True)['status'] == 'insufficient'
    assert history.compare(USER,'tools')['status'] == 'incompatible'
    for denied in (None, {'id':5,'modules':[]}):
        with pytest.raises(HTTPException) as exc:
            history.read_checks(denied,'tools')
        assert exc.value.status_code in (401,403)


def test_last_two_dates_delta_direction_and_error_are_factual():
    completed(6)
    completed(0)
    result = history.compare(USER,'edo')
    delta, = result['deltas']
    assert result['status'] == 'compared' and result['latest']['at'] and result['baseline']['at']
    assert (delta['previous'],delta['current'],delta['delta'],delta['direction'],delta['assessment']) == (6,0,-6,'down','improved')
    completed(99,status='error')
    assert history.compare(USER,'edo')['status'] == 'unavailable'
    assert history.read_checks(USER,'edo',status='error')[0]['metrics'] == {}


def test_rolling_month_uses_true_earliest_even_beyond_page_cap():
    as_of = datetime(2026,3,31,12,tzinfo=timezone.utc)
    outside = completed(100)
    first = completed(7)
    set_date(outside,'2026-02-27T15:00:00+00:00')
    set_date(first,'2026-02-28T15:00:00+00:00')
    for i in range(103):
        middle = completed(4)
        set_date(middle,(as_of-timedelta(hours=200-i)).isoformat())
    last = completed(2)
    set_date(last,'2026-03-31T11:00:00+00:00')
    result = history.compare(USER,'edo',window='month',as_of=as_of)
    assert result['status']=='compared' and result['deltas'][0]['delta']==-5
    assert result['baseline']['run_id']==first and result['latest']['run_id']==last
    assert len(history.read_checks(USER,'edo',limit=200)) == 100


def test_incompatible_periods_subchecks_and_unknown_params():
    completed(0,'cds',{'start':'2026-11-01','end':'2026-11-02'},metric='appeals')
    completed(2,'cds',{'start':'2026-11-02','end':'2026-11-04'},metric='appeals')
    assert history.compare(USER,'cds')['status']=='incompatible'
    assert len(history.read_checks(USER,'cds',params={'start':'2026-11-01'}))==1
    completed(5,'mingkh-water-map',metric='appeals')
    assert history.read_checks(USER,'mingkh') == []
    assert history.read_checks(USER,'mingkh-water-map')[0]['metrics'] == {'appeals':5}
    assert history.compare(USER,'cameras')['status']=='insufficient'


def test_legacy_unknown_source_and_old_timestamp_do_not_become_fresh(tmp_path):
    path=tmp_path/check_metrics.PATHS['utnkr']
    path.parent.mkdir(parents=True)
    ident=run_history.record_start('utnkr')
    path.write_text(json.dumps({'source_mode':'local_rows','rows_count':42,'violators_count':4}))
    run_history.record_finish(ident)
    assert history.read_checks(USER,'utnkr')[0]['result_code']=='unconfirmed'
    second=run_history.record_start('utnkr')
    path.write_text(json.dumps({'updated_at':'2024-01-01T00:00:00','violators_count':1,'rows_count':2}))
    run_history.record_finish(second)
    assert history.read_checks(USER,'utnkr')[0]['metrics']=={}


def test_another_live_process_keeps_its_check_until_exit():
    script = """import sys
from pathlib import Path
from utils import db
db.DB_FILE=Path(sys.argv[1])
from services import run_history
ident=run_history.record_start('edo')
print(ident,flush=True)
sys.stdin.readline()
run_history.record_finish(ident,metrics={'issues':3})"""
    child = subprocess.Popen([sys.executable,'-c',script,str(db.DB_FILE)],stdin=subprocess.PIPE,stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE,text=True)
    try:
        ident = child.stdout.readline().strip()
        assert ident
        with db.get_db_connection() as conn:
            conn.execute('UPDATE run_history SET started_at=? WHERE run_id=?',
                         ((datetime.now(timezone.utc)-timedelta(days=1)).isoformat(),ident))
        live=history.read_checks(USER,'edo')[0]
        assert live['status']=='running' and live['long_running']
        child.communicate(input='finish\n',timeout=12)
        assert child.returncode == 0
        assert history.read_checks(USER,'edo')[0]['metrics'] == {'issues':3}
    finally:
        if child.poll() is None:
            child.kill()
            child.wait(timeout=5)


def test_partial_and_incompatible_private_configs_do_not_generate_positive_trend():
    first = completed(2,params={'config_hash':'a'*64},private='local:5')
    completed(0,params={'config_hash':'b'*64},private='local:5')
    assert history.compare(USER,'edo')['status'] == 'incompatible'
    partial = completed(7,status='partial')
    record = history.read_checks(USER,'edo',status='partial')[0]
    assert record['run_id']==partial and record['result_code']=='partial' and record['metrics']=={'issues':7}
    assert history.compare(USER,'edo')['status']=='unavailable'
    assert history.read_checks(USER,'edo',params={'config_hash':'a'*64})[0]['run_id']==first


def test_newer_or_damaged_versions_are_excluded_from_comparisons():
    ident = completed(4)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE check_results SET version=2,metrics=? WHERE run_id=?',('sensitive incompatible payload',ident))
    record,=history.read_checks(USER,'edo')
    assert record['version']==2 and not record['metrics'] and record['result_code']=='legacy'
    with db.get_db_connection() as conn:
        conn.execute('UPDATE check_results SET version=1 WHERE run_id=?',(ident,))
    assert history.read_checks(USER,'edo')[0]['result_code']=='unconfirmed'


def test_partial_water_dashboard_never_compares_old_carried_forward_values(tmp_path):
    path = tmp_path/check_metrics.PATHS['water_dashboard']
    path.parent.mkdir(parents=True)
    ident = run_history.record_start('water_dashboard',params={'source':'all'})
    path.write_text(json.dumps({'last_checked_sources':['valves','tasks'], 'last_updated_sources':['valves'],
                                'sources':{'tasks':{'ok':True,'metric_schema':1,'metrics':[{'id':'total','value':777,'unit':'шт.'}]}}}))
    run_history.record_finish(ident)
    record, = history.read_checks(USER,'water-dashboard')
    assert record['status']=='partial' and record['result_code']=='partial'
    assert record['metrics']=={'sources':2,'failed_sources':1}
    assert history.compare(USER,'water-dashboard')['status']=='unavailable'


def test_manual_zip_registry_upload_records_partial_without_filenames(monkeypatch):
    import ast
    import io
    import zipfile
    from fastapi import FastAPI, Request
    from fastapi.testclient import TestClient
    from services.zip_curator import core as zc

    # Use the real multipart entry point and existing workbook decoder. Only
    # the registry state store is replaced; uploaded names must not enter history.
    app = FastAPI()
    path = Path(__file__).resolve().parents[1] / 'app.py'
    node = next(n for n in ast.parse(path.read_text(encoding='utf-8-sig')).body
                if isinstance(n,ast.AsyncFunctionDef) and n.name=='zc_upload')
    exec(compile(ast.Module(body=[node],type_ignores=[]),str(path),'exec'),
         {'app':app,'Request':Request,'zc':zc})
    monkeypatch.setattr(zc,'ingest',lambda records:len(records))
    monkeypatch.setattr(zc,'load_state',lambda:{})
    workbook = io.BytesIO()
    with zipfile.ZipFile(workbook,'w') as archive:
        archive.writestr('xl/worksheets/sheet1.xml', '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
          <row><c r="A1" t="inlineStr"><is><t>Наименование</t></is></c><c r="C1" t="inlineStr"><is><t>Количество</t></is></c></row>
          <row><c r="A2" t="inlineStr"><is><t>Труба</t></is></c><c r="C2"><v>5</v></c></row>
          </sheetData></worksheet>''')
    with TestClient(app) as client:
        response=client.post('/zip_curator/api/upload',files=[('file',('internal-name.xlsx',workbook.getvalue())),('file',('broken.xlsx',b'invalid'))])
    assert response.status_code==200 and response.json()['added']==1
    record,=history.read_checks({'id':5,'modules':['zips']},'zips')
    assert record['status']=='partial' and record['metrics']=={'files':2,'added_registries':1}
    with db.get_db_connection() as conn:
        assert 'internal-name' not in '\n'.join(conn.iterdump())


def test_manual_map_import_creates_one_measured_run_and_failed_validation_is_distinct(tmp_path,monkeypatch):
    from services.mingkh import water_map as maps
    monkeypatch.setattr(maps,'DATA_FILE',tmp_path/'map.json.gz')
    monkeypatch.setattr(maps,'SEED_FILE',tmp_path/'no-seed.gz')
    assert not maps.ensure_seed()
    assert history.read_checks(USER,'mingkh-water-map')==[]
    point={'id':5,'created':'2026-10-29 12:30:00','kind':'hvs','omsu':'synthetic','geo':'none'}
    meta=maps.import_upload(json.dumps({'points':{'5':point}}).encode())
    assert meta['count']==1
    record,=history.read_checks(USER,'mingkh-water-map')
    assert record['metrics']=={'appeals':1} and record['params']=={'operation':'upload','start':'2026-10-29','end':'2026-10-29'}
    with pytest.raises(ValueError):
        maps.import_upload(b'invalid archive')
    assert [r['status'] for r in history.read_checks(USER,'mingkh-water-map')]==['error','success']
