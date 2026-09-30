"""Bounded, read-only collection of Dobrodel reports for the EDDS dashboard."""
import json
import os
import re
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
    if path.exists():
        data=json.loads(path.read_text(encoding='utf-8-sig'))
    else:
        data={'days': {}, 'source': 'Добродел', 'error': 'Свод жалоб ещё не загружен. Настройте доступ к Доброделу и нажмите «Обновить жалобы».', 'available': False}
    data['building']=status()['running']
    return data


def failure_message(output):
    """Only fixed diagnostics; Playwright traces can include form values."""
    if 'Администратор должен настроить логин и пароль' in output:
        return 'Логин и пароль Добродела не настроены в разделе «Пользователи → Логины и пароли».'
    if ('Executable doesn\'t exist' in output or ('executable_path' in output and 'does not exist' in output)
            or 'EDDS_CHROME_EXECUTABLE не указывает' in output):
        return 'Не найден браузер для сбора жалоб. Проверьте EDDS_CHROME_EXECUTABLE на компьютере-сервере.'
    if 'Не удалось войти.' in output:
        return 'Добродел не подтвердил вход. Проверьте сохранённые логин и пароль, капчу или код подтверждения.'
    if 'пустой отчёт за все запрошенные периоды' in output:
        return 'Добродел вернул пустой отчёт за все периоды. Прежние данные сохранены; проверьте права и фильтр МинЖКХ.'
    match=re.search(r'Сервер вернул ошибку (\d{3}) на странице', output)
    if match:
        return f'Добродел вернул HTTP {match.group(1)} при загрузке жалоб. Проверьте права доступа и доступность портала.'
    if 'Chromium не смог получить отчёт Добродела по сети.' in output:
        return 'Chromium на компьютере-сервере не смог получить отчёт Добродела. Проверьте доступ к порталу в этом браузере.'
    if 'Добродел не ответил за 60 секунд' in output:
        return 'Добродел не ответил на запрос отчёта за 60 секунд. Повторите позже или проверьте доступ на компьютере-сервере.'
    if 'Timeout' in output or 'ERR_TIMED_OUT' in output:
        return 'Добродел не ответил вовремя с компьютера-сервера. Проверьте его доступность без VPN.'
    if 'ERR_CERT' in output:
        return 'Браузер на компьютере-сервере не смог проверить сертификат Добродела.'
    return 'Сбор жалоб завершился с ошибкой на компьютере-сервере. Проверьте доступ к Доброделу.'


def run(user='Авто-запуск'):
    if not credentials(): return False
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        row=conn.execute('SELECT * FROM edds_job WHERE id=1').fetchone()
        if row['running'] and time.time()-row['started_at']<1900: return False
        conn.execute("UPDATE edds_job SET running=1,started_at=?,message='Сбор жалоб выполняется' WHERE id=1",(time.time(),))
    run_id=record_start('edds',user=user)
    ok=False
    message='Сбор жалоб не завершён.'
    try:
        DATA.mkdir(parents=True,exist_ok=True)
        result=subprocess.run([sys.executable,'-m','services.edds.collector'],cwd=BASE,timeout=1800,
                              stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                              text=True,encoding='utf-8',errors='replace',
                              env={**os.environ, 'PYTHONIOENCODING': 'utf-8'})
        ok=result.returncode==0
        if ok: message='Свод жалоб обновлён'
        else: message=failure_message(result.stdout or '')
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
