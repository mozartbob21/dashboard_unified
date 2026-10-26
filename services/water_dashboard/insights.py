"""Bounded, durable model commentary for verified aggregate dashboard metrics.

Only canonical metric labels, finite numbers and server-computed trends reach
Qwen. Numeric comparison and trend labels remain server-owned. Model prose is
bounded and screened for obvious unsupported claims, not semantically certified.
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import hashlib
import json
import math
import re
from pathlib import Path
import sqlite3
import time
import uuid

from services.water_dashboard.config import DATA_DIR
from services.water_dashboard.metrics import source_metrics

_POOL = ThreadPoolExecutor(max_workers=2, thread_name_prefix='water-insight')
LEASE_SECONDS = 240
MAX_CACHE_ROWS = 100
HIGHER = {('valves','completion_pct'), ('edo_rso','signers_ecp_pct'),
          ('edo_rso','electronic_documents_pct'), ('meetings','attendance_pct'),
          ('meetings','rso_attendance_pct'), ('energy_audit','optimal_tariff_pct'),
          ('gkh_kontur','completion_pct'), ('nvos','collection_pct')}
LOWER = {('tasks','overdue_tasks'), ('controller_tasks','overdue_tasks'),
         ('sys_vs','system_addresses'), ('sys_vs','resonant_addresses'),
         ('sys_kr','system_addresses'), ('sys_kr','resonant_addresses')}
# These checks reject obvious format/numeric/causal violations. They cannot
# establish the truth of arbitrary natural language; the UI labels model output.
_DISALLOWED_PROSE = re.compile(r'[\d<>]|https?://|www\.|\b(?:ftp|javascript|data):|[\x00-\x1f\x7f]', re.I)
_CAUSAL_CLAIMS = re.compile(r'благодаря|вызван|обусловлен|привел[оаи]? к|привёл|вследствие|из-за|причина (?:в|состоит)|это результат', re.I)
_UNKNOWN_ASSESSMENT = re.compile(r'улучш|ухудш|вырос|возрос|снизил|снижен|раст[её]т|положительн|отрицательн|стабил|сохранил|достигнут|критическ|лидир|отста[её]', re.I)



def _number(value):
    try:
        return value if type(value) in (int, float) and math.isfinite(value) and 0 <= value <= 1e18 else None
    except (OverflowError, ValueError):
        return None


def comparison_scopes(source):
    # Population hashes stay on the server; identities never enter model facts.
    from services.water_dashboard.builder import build_table
    from services.water_dashboard.details import meeting_rso_records
    def digest(values):
        return hashlib.sha256(json.dumps(sorted(values),ensure_ascii=False).encode()).hexdigest() if values else None
    om=build_table({'meetings':source})
    rs, excluded=meeting_rso_records(source.get('rso') or {})
    return {'attendance_pct':digest([r['name'].casefold() for r in om if r.get('att') is not None]),
            'rso_attendance_pct':digest([(r['name'].casefold(),r['municipality'].casefold()) for r in rs]) if not excluded else None}


def _same_calendar_period(source_id, metric_id, current, previous):
    weekly = source_id == 'edo_rso' or (source_id == 'nvos' and metric_id.endswith('_week'))
    annual = source_id == 'valves' or (source_id == 'nvos' and metric_id.endswith('_year'))
    if not (weekly or annual): return True
    try:
        a=datetime.fromisoformat(current['updated_at']); b=datetime.fromisoformat(previous['updated_at'])
        return a.isocalendar()[:2] == b.isocalendar()[:2] if weekly else a.year == b.year
    except (KeyError,TypeError,ValueError): return False


def observation(source):
    """A compact successful observation, with no rows, source text or URLs."""
    if not isinstance(source, dict) or source.get('metric_schema') != 1 or not source.get('updated_at'):
        return None
    metrics = source.get('metrics')
    if not isinstance(metrics, list) or not metrics or not isinstance(metrics[0],dict) or _number(metrics[0].get('value')) is None:
        return None
    return {'metric_schema': 1, 'updated_at': source['updated_at'], 'data_date': source.get('data_date'),
            'comparison_scopes': comparison_scopes(source),
            'metrics': [{'id': m.get('id'), 'value': _number(m.get('value')), 'unit': m.get('unit')}
                        for m in metrics if isinstance(m,dict)]}


def facts_for(source_id, source):
    schema = source_metrics(source_id, {}, [])
    if source_id == 'gkh_kontur':
        from services.water_dashboard.details import KONTUR_STAGE_LABELS
        schema += [{'id':f'stage_{i}_pending','label':label+': не завершили','unit':'РСО / подразделений'}
                   for i,label in enumerate(KONTUR_STAGE_LABELS,1)]
    scopes = comparison_scopes(source) if source_id == 'meetings' else {}
    def indexed(items):
        result = {}
        for item in items if isinstance(items,list) else []:
            if not isinstance(item,dict) or not isinstance(item.get('id'),str): continue
            key = item['id']
            result[key] = None if key in result else item  # conflicting/duplicate ids are unknown
        return result
    current = indexed(source.get('metrics'))
    previous = source.get('previous_observation') or {}
    if not isinstance(previous,dict): previous = {}
    before = indexed(previous.get('metrics'))
    same_period = (previous.get('metric_schema') == 1 and bool(previous.get('updated_at'))
                   and previous.get('data_date') == source.get('data_date'))
    facts = []
    for spec in schema:
        metric_id, unit = spec['id'], spec['unit']
        now = current.get(metric_id) or {}
        old = before.get(metric_id) or {}
        value = _number(now.get('value')) if now.get('unit') == unit else None
        comparable_scope = (source_id != 'meetings' or metric_id not in scopes or
            (bool(scopes[metric_id]) and scopes[metric_id] == (previous.get('comparison_scopes') or {}).get(metric_id)))
        prior = _number(old.get('value')) if old.get('unit') == unit and same_period and comparable_scope and _same_calendar_period(source_id,metric_id,source,previous) else None
        comparable = value is not None and prior is not None
        direction = 1 if (source_id,metric_id) in HIGHER else -1 if (source_id,metric_id) in LOWER else 0
        trend = 'unknown'
        if comparable:
            if value == prior: trend = 'stable'
            elif direction: trend = 'improved' if (value-prior)*direction > 0 else 'worsened'
        facts.append({'metric_id':metric_id, 'label':spec['label'], 'unit':unit,
                      'value':value, 'previous':prior, 'comparable':comparable, 'trend':trend})
    return facts


def _display(value):
    return format(value, '.12g').replace('.', ',')


def _summary(fact):
    if fact['value'] is None:
        return 'Подтверждённое числовое значение недоступно; оценка динамики невозможна.'
    current = _display(fact['value']) + (' '+fact['unit'] if fact['unit'] else '')
    if not fact['comparable']:
        return f'Текущее значение — {current}. Сопоставимого предыдущего наблюдения нет.'
    old = _display(fact['previous']) + (' '+fact['unit'] if fact['unit'] else '')
    descriptions = {'improved':'Изменение оценивается как улучшение', 'worsened':'Изменение оценивается как ухудшение',
                    'stable':'Числовое значение не изменилось', 'unknown':'Изменение само по себе не определяет улучшение или ухудшение'}
    return f'Текущее значение — {current}; предыдущее — {old}. {descriptions[fact["trend"]]}.'


def model_commentary(facts):
    from services.summarizer.engine import _qwen_chat
    import httpx
    prompt = {'metrics': facts}
    raw = _qwen_chat([
        {'role':'system','content':'Ты аналитик ЖКХ. Подготовь короткий профессиональный мини-вывод к каждому показателю '
         'с учётом его назначения, единицы, доступности сравнения и заданного сервером trend. '
         'Числа и оценку динамики сервер уже покажет отдельно: не повторяй значения, проценты, даты и числа словами. '
         'Дай одну-две конкретные фразы о том, что стоит проверить или учитывать при интерпретации именно этого показателя. '
         'Не устанавливай причины, нормативы, достигнутые цели или последствия: таких данных нет. '
         'Не называй положение критическим или благополучным. Не приписывай эффект мерам или ответственным. '
         'При trend=unknown используй нейтральный вывод о недостающих данных/контексте; '
         'не утверждай улучшение, ухудшение или стабильность. Нейтральные счётчики не оценивай как успех или проблему. '
         'Никаких имён, ссылок, HTML, Markdown. Не обращайся к внешним источникам. '
         'Верни только JSON-массив [{"metric_id":"...","text":"профессиональный мини-вывод"}], '
         'по одному объекту на каждый переданный metric_id, текст до шестисот символов.'},
        {'role':'user','content':json.dumps(prompt,ensure_ascii=False,allow_nan=False)}],
        module_id='water-dashboard', max_tokens=2500,
        timeout=httpx.Timeout(connect=5,read=60,write=10,pool=5))
    return raw


def validated_items(raw, facts):
    if not isinstance(raw,str) or len(raw)>24000:
        raise ValueError('invalid commentary')
    text = raw.strip()
    if text.startswith('```json\n') and text.endswith('```'): text=text[8:-3].strip()
    elif text.startswith('```\n') and text.endswith('```'): text=text[4:-3].strip()
    data=json.loads(text)
    if not isinstance(data,list) or len(data)!=len(facts): raise ValueError('invalid item count')
    expected={fact['metric_id']:fact for fact in facts}
    result={}
    for item in data:
        if not isinstance(item,dict) or set(item)!={'metric_id','text'}: raise ValueError('invalid item')
        mid=item['metric_id']
        if not isinstance(mid,str) or mid not in expected or mid in result: raise ValueError('invalid metric')
        fact=expected[mid]
        prose=item['text']
        if not isinstance(prose,str) or not 20 <= len(prose.strip()) <= 600:
            raise ValueError('invalid commentary length')
        prose=prose.strip()
        if _DISALLOWED_PROSE.search(prose) or _CAUSAL_CLAIMS.search(prose):
            raise ValueError('unsupported claim')
        if fact['trend']=='unknown' and _UNKNOWN_ASSESSMENT.search(prose):
            raise ValueError('unsupported assessment')
        result[mid]={'metric_id':mid,'text':_summary(fact)+' '+prose,'trend':fact['trend']}
    return [result[fact['metric_id']] for fact in facts]


class InsightService:
    def __init__(self, path=None, *, model=None, submit=None, clock=None):
        self.path=Path(path or DATA_DIR/'insights.sqlite')
        self.model=model or model_commentary
        self.submit=submit or _POOL.submit
        self.clock=clock or time.time

    def _connect(self):
        self.path.parent.mkdir(parents=True,exist_ok=True)
        conn=sqlite3.connect(self.path,timeout=10)
        conn.row_factory=sqlite3.Row
        conn.execute('CREATE TABLE IF NOT EXISTS insight_jobs (revision TEXT PRIMARY KEY, source_id TEXT NOT NULL, status TEXT NOT NULL, owner TEXT NOT NULL, lease_until REAL NOT NULL, touched REAL NOT NULL, result TEXT)')
        conn.commit()
        return conn

    def _unavailable(self, base, message='Вывод ИИ пока недоступен. Повторите попытку позже.'):
        return dict(base,status='unavailable',message=message,items=[])

    def request(self, source_id, source, *, start=False):
        source = source if isinstance(source,dict) else {}
        facts = facts_for(source_id,source)
        updated = source.get('updated_at','')
        updated = updated if isinstance(updated,str) and len(updated)<=80 else ''
        revision=hashlib.sha256(json.dumps({'v':1,'source':source_id,'updated':updated,'facts':facts,
                'data_date':str(source.get('data_date') or '')[:80]},sort_keys=True,ensure_ascii=False,allow_nan=False).encode()).hexdigest()
        base={'revision':revision,'source_updated_at':updated}
        if not source.get('ok') or source.get('metric_schema')!=1 or not updated or not any(f['value'] is not None for f in facts):
            return self._unavailable(base,'Для вывода ИИ сначала обновите подтверждённые данные источника.')
        now=self.clock(); owner=uuid.uuid4().hex; launch=False
        conn=self._connect()
        try:
            conn.execute('BEGIN IMMEDIATE')
            row=conn.execute('SELECT * FROM insight_jobs WHERE revision=?',(revision,)).fetchone()
            if row and row['status']=='ready':
                return json.loads(row['result'])
            if row and row['status']=='pending' and row['lease_until']>now:
                return dict(base,status='pending',retry_after=3)
            if not start:
                return self._unavailable(base,'Вывод ещё не подготовлен. Откройте карточку повторно для запуска анализа.')
            if row and row['status']=='unavailable' and row['touched']+60>now:
                return self._unavailable(base)
            active=conn.execute("SELECT COUNT(*) FROM insight_jobs WHERE status='pending' AND lease_until>?",(now,)).fetchone()[0]
            if active>=2:
                return self._unavailable(base,'ИИ обрабатывает другие карточки. Повторите попытку через минуту.')
            conn.execute("INSERT INTO insight_jobs VALUES(?,?,'pending',?,?,?,NULL) ON CONFLICT(revision) DO UPDATE SET status='pending',owner=excluded.owner,lease_until=excluded.lease_until,touched=excluded.touched,result=NULL",(revision,source_id,owner,now+LEASE_SECONDS,now))
            conn.execute("DELETE FROM insight_jobs WHERE revision NOT IN (SELECT revision FROM insight_jobs ORDER BY touched DESC, rowid DESC LIMIT ?) AND (status!='pending' OR lease_until<=?)",(MAX_CACHE_ROWS,now))
            conn.commit(); launch=True
        finally:
            conn.close()
        if launch:
            try: self.submit(self._run,revision,owner,facts,base)
            except Exception: self._finish(revision,owner,'unavailable',None)
        return dict(base,status='pending',retry_after=3)

    def _run(self,revision,owner,facts,base):
        try:
            items=validated_items(self.model(facts),facts)
            result=dict(base,status='ready',items=items,generated_at=datetime.fromtimestamp(self.clock(),timezone.utc).isoformat())
            self._finish(revision,owner,'ready',result)
        except Exception:
            self._finish(revision,owner,'unavailable',None)

    def _finish(self,revision,owner,status,result):
        conn=self._connect()
        try:
            now=self.clock()
            conn.execute('UPDATE insight_jobs SET status=?,result=?,touched=? WHERE revision=? AND owner=? AND lease_until>?',
                         (status,json.dumps(result,ensure_ascii=False) if result else None,now,revision,owner,now))
            conn.commit()
        finally: conn.close()
