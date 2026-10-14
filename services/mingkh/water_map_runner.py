"""Scheduler adapter for the same bounded incremental map collector used by admins."""
from services.auth.integrations import credentials
from services.mingkh import water_map


MODULE_ID = 'mingkh-water-map'


def run(user='Автозапуск'):
    """Return a safe scheduler outcome; collection never runs in a subprocess."""
    try:
        account = credentials('edds')
    except Exception:
        account = None
    if not isinstance(account, dict) or not all(isinstance(account.get(key), str) and account[key]
                                               for key in ('username', 'password')):
        error = 'Для обновления карты администратору нужно сохранить доступ к Доброделу.'
        from services.run_history import record_start, record_finish
        record_finish(record_start(MODULE_ID, user=user), status='error', error_message=error)
        return {'status': 'error', 'error': error}
    try:
        ok = water_map.run_refresh({'username': account['username'], 'password': account['password']}, user=user)
    except water_map.StoreBusy:
        return {'status': 'skipped', 'error': 'Карта уже обновляется.'}
    except ValueError as exc:
        error = str(exc)
        from services.run_history import record_start, record_finish
        record_finish(record_start(MODULE_ID, user=user), status='error', error_message=error)
        return {'status': 'error', 'error': error}
    return {'status': 'success' if ok else 'error',
            'error': '' if ok else water_map.status().get('error') or 'Не удалось обновить карту МИНЖКХ.'}
