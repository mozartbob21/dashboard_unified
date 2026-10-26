"""Atomic persisted snapshots and refresh lease shared by all web workers."""
import json
import sqlite3
import time
import uuid
from contextlib import contextmanager
from datetime import datetime,timezone
from pathlib import Path
from . import calculate, source

DATA_DIR=Path(__file__).resolve().parents[2]/'data/omsu_rating'
LEASE_SECONDS=900

def now_iso():return datetime.now(timezone.utc).isoformat()
@contextmanager
def connection():
    DATA_DIR.mkdir(parents=True,exist_ok=True)
    conn=sqlite3.connect(DATA_DIR/'rating.sqlite',timeout=15);conn.row_factory=sqlite3.Row
    try:
        conn.execute('PRAGMA busy_timeout=15000')
        conn.executescript('''CREATE TABLE IF NOT EXISTS state(id INTEGER PRIMARY KEY CHECK(id=1),sources TEXT NOT NULL DEFAULT '{}',updated_at TEXT);
        INSERT OR IGNORE INTO state(id) VALUES(1);
        CREATE TABLE IF NOT EXISTS refresh(id INTEGER PRIMARY KEY CHECK(id=1),token TEXT,until REAL,status TEXT,error TEXT,finished_at TEXT);
        INSERT OR IGNORE INTO refresh(id,status) VALUES(1,'idle');
        CREATE TABLE IF NOT EXISTS snapshots(id INTEGER PRIMARY KEY,sources TEXT NOT NULL,created_at TEXT NOT NULL);''')
        yield conn
        conn.commit()
    except Exception:conn.rollback();raise
    finally:conn.close()

def _status(conn):
    row=dict(conn.execute('SELECT status,error,finished_at,until FROM refresh WHERE id=1').fetchone())
    if row['status']=='running' and (row['until'] or 0)<=time.time():row.update(status='error',error='Обновление было прервано. Можно повторить; предыдущие данные сохранены.')
    row.pop('until');return row

def status():
    with connection() as conn:return _status(conn)

def read():
    with connection() as conn:
        row=conn.execute('SELECT * FROM state WHERE id=1').fetchone()
        return {**calculate.build(json.loads(row['sources'])),'updated_at':row['updated_at'],'refresh':_status(conn),'history':[dict(r) for r in conn.execute('SELECT id,created_at FROM snapshots ORDER BY id DESC LIMIT 12')]}

def claim():
    with connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        current=conn.execute('SELECT * FROM refresh WHERE id=1').fetchone()
        if current['status']=='running' and current['until']>time.time():return None
        token=uuid.uuid4().hex
        conn.execute("UPDATE refresh SET token=?,until=?,status='running',error=NULL WHERE id=1",(token,time.time()+LEASE_SECONDS))
        return token

def refresh(token):
    try:
        from services.auth.source_api_keys import get_source_api_key
        key=get_source_api_key('omsu_rating')
        with connection() as conn: previous=json.loads(conn.execute('SELECT sources FROM state WHERE id=1').fetchone()[0])
        sources=dict(previous);success=0;stamp=now_iso()
        for ind in (4,5,6,7):
            try:
                if ind<7 and not key:raise source.SourceError('API-ключ не настроен. Администратор может добавить его в разделе «Пароли».')
                values=source.fetch_redmine(calculate.CONFIG[ind]['query'],key) if ind<7 else source.fetch_sheet()
                sources[str(ind)]={'data':values,'updated_at':stamp,'error':None};success+=1
            except source.SourceError as exc:
                sources[str(ind)]={**sources.get(str(ind),{}),'error':str(exc)}
        # Calculations must succeed before committing the replacement snapshot.
        calculate.build(sources)
        with connection() as conn:
            conn.execute('BEGIN IMMEDIATE')
            active=conn.execute('SELECT token,until FROM refresh WHERE id=1').fetchone()
            if active['token']!=token or active['until']<=time.time():return
            encoded=json.dumps(sources,ensure_ascii=False,allow_nan=False)
            conn.execute('UPDATE state SET sources=?,updated_at=CASE WHEN ? THEN ? ELSE updated_at END WHERE id=1',(encoded,success,stamp))
            if success:
                conn.execute('INSERT INTO snapshots(sources,created_at) VALUES(?,?)',(encoded,stamp))
                conn.execute('DELETE FROM snapshots WHERE id NOT IN (SELECT id FROM snapshots ORDER BY id DESC LIMIT 12)')
            conn.execute('UPDATE refresh SET status=?,error=?,until=0,finished_at=? WHERE id=1',('success' if success==4 else 'partial' if success else 'error',None if success==4 else 'Не все источники обновлены. Причины указаны в показателях; ранее полученные данные сохранены.',stamp))
    except Exception:
        with connection() as conn:
            conn.execute("UPDATE refresh SET status='error',error=?,until=0,finished_at=? WHERE id=1 AND token=?",('Не удалось завершить обновление. Сохранённые данные доступны; повторите попытку.',now_iso(),token))
