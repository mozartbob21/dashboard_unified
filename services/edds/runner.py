"""Bounded, read-only collection of Dobrodel reports for the EDDS dashboard."""
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path
from services.auth.integrations import credentials
from services.run_history import record_start, record_finish
from utils.db import get_db_connection

BASE=Path(__file__).resolve().parents[2]
DATA=BASE/'data'/'edds'


def status():
    with get_db_connection() as conn:
        row=dict(conn.execute('SELECT * FROM edds_job WHERE id=1').fetchone())
    row['running']=bool(row['running'] and time.time()-row['started_at']<1900)
    return row


def water_daily():
    path=DATA/'water_daily.json'
    if not path.exists(): path=Path(__file__).with_name('water_daily_seed.json')
    data=json.loads(path.read_text(encoding='utf-8-sig'))
    data['building']=status()['running']
    return data


def run(user='Авто-запуск'):
    if not credentials(): return False
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        row=conn.execute('SELECT * FROM edds_job WHERE id=1').fetchone()
        if row['running'] and time.time()-row['started_at']<1900: return False
        conn.execute("UPDATE edds_job SET running=1,started_at=?,message='Сбор жалоб выполняется' WHERE id=1",(time.time(),))
    run_id=record_start('edds',user=user)
    ok=False
    message='Не удалось обновить жалобы. Проверьте доступ к Доброделу, логин и пароль, установку Chromium. Возможна капча или двухфакторная проверка.'
    try:
        DATA.mkdir(parents=True,exist_ok=True)
        if not (DATA/'water_daily.json').exists():
            shutil.copyfile(Path(__file__).with_name('water_daily_seed.json'),DATA/'water_daily.json')
        result=subprocess.run([sys.executable,'-m','services.edds.collector'],cwd=BASE,timeout=1800,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        ok=result.returncode==0
        if ok: message='Свод жалоб обновлён'
    except subprocess.TimeoutExpired:
        message='Сбор остановлен по лимиту времени (30 минут). Повторите позднее.'
    except Exception:
        message='Не удалось запустить сборщик. Проверьте настройки сервера.'
    finally:
        with get_db_connection() as conn:
            conn.execute('UPDATE edds_job SET running=0,message=? WHERE id=1',(message,))
        record_finish(run_id,status='success' if ok else 'error',error_message='' if ok else message)
    return ok


if __name__=='__main__':
    raise SystemExit(0 if run() else 1)
