# -*- coding: utf-8 -*-
"""Сбор жалоб Добродела для свода ЕДДС.

Заходит на портал под отдельной учётной записью из настроек «Нейроны»,
запрашивает отчёт МинЖКХ небольшими периодами и сохраняет water_daily.json.
Пароль в файл свода не записывается.
"""

import os
import re
import sys
import html
import json
import datetime as dt
import time
from pathlib import Path

from services.auth.integrations import credentials
from services.edds.dobrodel import DobrodelClient, DobrodelError

# ─────────────────────── НАСТРОЙКИ ───────────────────────
HERE = Path(__file__).resolve().parents[2] / "data" / "edds"
HTML = HERE / "Контроль_жалоб_ЕЦУР (2).html"
DATA_JS = HERE / "data.js"
AUTH = HERE / "auth.json"

BASE = "https://admin.vmeste.mosreg.ru"
REPORT_URL = BASE + "/OperativeReportGenerating"
LOGIN_URL = BASE + "/login"

KEYRING_SERVICE = "vmeste-kliker"

# Фильтр свода
CURATOR = "Министерство жилищно-коммунального хозяйства Московской области"
# ВСЕ статусы ДоброДела. Коды сняты со страницы отчёта портала (<option name="status">).
# Почему все, а не только активные: график «жалобы по дням» строится по дате подачи.
# Если брать лишь незакрытые, прошлые дни проседают не потому, что жалоб было меньше,
# а потому что их успели отработать, — и параллель с инцидентами читается неверно.
STATUSES = ",".join([
    "9",                            # премодерация видео
    "30", "31", "32", "34", "35",   # на модерации, опубликовано, в работе, получен ответ, решено
    "37", "38",                     # направлено в работу, предоставлен ответ
    "50", "51",                     # в работе (просрочено), повторная отправка
    "53", "56", "60",               # закрыто, закрыто пользователем
    "54", "57",                     # указан срок, указан срок (просрочено)
    "61",                           # сбор подписантов
    "111", "310", "320", "330",     # премодерация, на рассмотрении, на уточнении, в работе (согл.)
    "511", "512", "513", "515", "531",  # доработка, подготовлен ответ, ЕДЦ, модерация ЕДС, закрыто в ЕДС
])
# Намеренно НЕ берём: 4 и 62 — незавершённая регистрация и черновик коллективной
# жалобы, 42 и 55 — отклонено модератором. Это обращения, которые жалобами так
# и не стали, и в потоке «сколько поступило за день» им не место.

# Статусы для data.js — то, что ждёт дашборд «Контроль жалоб ЕЦУР»: открытые
# заявки. Ровно тот фильтр, что стоял в кликере изначально.
ACTIVE_STATUSES = "37,32,50,54,57,51,511,512"

# Свод по воде копится. Каждый прогон перетягивает только хвост, история
# остаётся в файле, поэтому глубина растёт сама и ничем не ограничена сверху.
OVERLAP = int(os.environ.get("KLIKER_OVERLAP", "7"))     # дней перезапроса внахлёст
# Limit work per run; the portal's operative report times out on long windows.
# Older history is backfilled over later runs.
MAX_CATCHUP = max(1, int(os.environ.get("KLIKER_CATCHUP", "30")))
MAX_WINDOW_DAYS = 7
COLLECTOR_VERSION = '2026-10-01.2'
KEEP_DAYS = int(os.environ.get("KLIKER_KEEP", "1825"))   # сколько истории держим, 0 = вечно

# Глубина истории по «Дате подачи» (filters.createdAfter/createdBefore на портале).
# Желаемая глубина истории. Свод копится, поэтому глубокий сбор случается один
# раз: дальше кликер тянет только хвост. Поднять глубину можно в любой момент —
# следующий прогон доберёт недостающее назад (см. pull_windows).
# Меняется без правки кода: set KLIKER_DAYS=365 перед запуском.
DAYS_BACK = int(os.environ.get("KLIKER_DAYS", "730"))

# Свод для графика ЕДДС: жалобы по воде, свёрнутые в «дата → ОМСУ → сколько».
WATER_JSON = HERE / "water_daily.json"

PAGE_SIZE = 500

# Колонки, которые ждёт дашборд (порядок важен — совпадает с рядами ниже)
HEADER = ["Номер", "Адрес", "Район", "Категория ЕЦУР", "Подкатегория ЕЦУР",
          "ЕЦУР факт", "Исполнитель", "Текст", "Дата создания", "Срок", "Статус"]


def log(msg):
    print(msg, flush=True)


def progress(message):
    """Only locally composed stages and counters, never portal responses."""
    log(message)
    path = WATER_JSON.with_name('progress.json')
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix('.tmp')
        tmp.write_text(json.dumps({'updated_at': time.time(), 'message': message,
            'collector_version': COLLECTOR_VERSION}, ensure_ascii=False), encoding='utf-8')
        tmp.replace(path)
    except OSError:
        pass  # An unavailable progress file must not discard a valid report.


def report_progress(event):
    period = f"{event['from']} — {event['to']}"
    if event['stage'] == 'split':
        progress(f'Добродел: делю медленный отчёт {period} на меньшие периоды.')
    elif event['stage'] == 'retry':
        progress(f'Добродел: повторяю отчёт {period} меньшими страницами.')
    else:
        progress(f"Добродел: {period}, страница {event['page']}, получено {event['count']} обращений.")


def get_credentials():
    try:
        value = credentials('edds')
    except Exception:
        raise DobrodelError('config') from None
    if not value:
        raise DobrodelError('credentials')
    return value


def clean(v):
    if v is None:
        return ""
    return html.unescape(str(v)).strip()


def build_rows(records):
    rows = [HEADER]
    for r in records:
        rows.append([
            r.get("cardId", ""),
            clean(r.get("address")),
            clean(r.get("district")),
            clean(r.get("ecurCategory")),
            clean(r.get("subcategory")),
            clean(r.get("ecurFact")) or "—",
            clean(r.get("org")),
            clean(r.get("body")),
            clean(r.get("created")),
            clean(r.get("deadline")),
            clean(r.get("status")),
        ])
    return rows


# ─────────────────── свод «жалобы по воде» для дашборда ЕДДС ───────────────────
# Дашборд ЕДДС считает инциденты по объектам ВС/ВО: ВЗУ, сети водоснабжения, КНС,
# КОС, сети водоотведения. Это холодная вода и водоотведение. ГВС там отдельный
# ресурс и в выборку не входит, поэтому в основную линию графика его не кладём —
# но считаем отдельно, чтобы цифра не потерялась.
RE_GVS = re.compile(r"горяч\w*\s*вод|\bгвс\b", re.I)
RE_VO = re.compile(r"водоотвед|канализ|\bстоки\b|очистн|\bкос\b|\bкнс\b|септик|выгреб", re.I)
RE_HVS = re.compile(r"холодн\w*\s*вод|\bхвс\b|водоснабж|водопровод|водозабор|"
                    r"водоразбор|колонк|скважин|подвоз\w*\s*вод|качеств\w*\s*вод|ржав", re.I)


def water_kind(subcat, fact, cat):
    """'gvs' | 'vo' | 'hvs' | None. Порядок важен: ГВС проверяем первым,
    иначе «горячее водоснабжение» поймает шаблон водоснабжения."""
    t = " ".join([subcat or "", fact or "", cat or ""])
    if RE_GVS.search(t):
        return "gvs"
    if RE_VO.search(t):
        return "vo"
    if RE_HVS.search(t):
        return "hvs"
    return None


def day_of(created):
    """«Дата создания» портала → 'ГГГГ-ММ-ДД'. Формат бывает ISO и ДД.ММ.ГГГГ."""
    s = str(created or "").strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return "%s-%s-%s" % m.groups()
    m = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})", s)
    if m:
        return "%s-%s-%s" % (m.group(3), m.group(2), m.group(1))
    return None


def load_water():
    try:
        return json.loads(WATER_JSON.read_text(encoding="utf-8"))
    except Exception:
        return None


def _day(v):
    try:
        return dt.date.fromisoformat(v or "")
    except (ValueError, TypeError):
        return None


def pull_windows(existing):
    """Окна по дате подачи для этого прогона — список (с, по, зачем).

    Хвост берём всегда: свежие дни с нахлёстом, чтобы подхватить поздние
    регистрации. Добор назад — когда желаемая глубина DAYS_BACK больше того,
    что уже накоплено: иначе свод рос бы только вперёд и прошлое навсегда
    осталось бы недостижимым. Поднял KLIKER_DAYS, запустил — история удлинилась."""
    until = dt.date.today()
    want_from = until - dt.timedelta(days=DAYS_BACK)
    first = _day((existing or {}).get("from"))
    last = _day((existing or {}).get("to"))

    def small_windows(start, finish, reason):
        result = []
        while start <= finish:
            beginning = max(start, finish - dt.timedelta(days=MAX_WINDOW_DAYS - 1))
            result.append((beginning, finish, reason))
            finish = beginning - dt.timedelta(days=1)
        return result

    if last is None:
        start = max(want_from, until - dt.timedelta(days=MAX_CATCHUP - 1))
        return small_windows(start, until, 'первый сбор')

    wins = []
    since = last - dt.timedelta(days=OVERLAP)
    floor = until - dt.timedelta(days=MAX_CATCHUP - 1)
    if since < floor:
        log(f"⚠ Свод не обновлялся с {last}. Беру последние {MAX_CATCHUP} дн., "
            f"в истории останется разрыв — при необходимости увеличьте KLIKER_CATCHUP.")
        since = floor
    # Recover gaps left by an interrupted earlier run. An absent day is unknown;
    # a saved empty bucket is a successfully checked zero.
    known = (existing or {}).get('days') or {}
    for offset in range(MAX_CATCHUP):
        day = until - dt.timedelta(days=offset)
        if (first is None or day >= first) and day.isoformat() not in known:
            since = min(since, day)
    since = max(want_from, since)
    wins.extend(small_windows(since, until, f"хвост, нахлёст {OVERLAP} дн."))

    if first is not None and want_from < first:
        back_to = min(first, since) - dt.timedelta(days=1)
        back_from = max(want_from, back_to - dt.timedelta(days=MAX_CATCHUP - 1))
        wins.extend(small_windows(back_from, back_to, f"добор назад до {DAYS_BACK} дн."))
    return wins


def write_water_daily(rows, since, until, existing):
    """rows — сетка с HEADER в первой строке за окно [since, until].
    Свод накопительный: свежее окно заменяем целиком, всё что вне его —
    оставляем как накопилось. День → ОМСУ → [ХВС, водоотведение, ГВС].
    Пустой словарь дня означает проверенный ноль; отсутствие дня — неизвестно.
    Старые разреженные своды дополняем только реально запрошенным окном."""
    start, finish = _day(since), _day(until)
    if start is None or finish is None or start > finish:
        raise ValueError('Некорректное окно свода жалоб')
    # The dashboard also derives coverage from day keys. Explicit empty buckets
    # preserve queried boundary days and replace stale counts in browser caches.
    fresh = {(start + dt.timedelta(days=i)).isoformat(): {}
             for i in range((finish - start).days + 1)}
    hit = 0
    for r in rows[1:]:
        kind = water_kind(r[4], r[5], r[3])
        if not kind:
            continue
        d = day_of(r[8])
        if d not in fresh:
            continue
        omsu = (r[2] or "").strip() or "—"
        slot = fresh.setdefault(d, {}).setdefault(omsu, [0, 0, 0])
        slot[{"hvs": 0, "vo": 1, "gvs": 2}[kind]] += 1
        hit += 1

    days = dict((existing or {}).get("days") or {})
    was = len(days)
    # Дни окна вычищаем перед вставкой, а не просто перезаписываем: если жалобу
    # удалили и день опустел, старое число не должно там остаться.
    for d in [x for x in days if since <= x <= until]:
        del days[d]
    days.update(fresh)

    # потолок хранения, чтобы свод не рос бесконечно (0 — держать всё)
    if KEEP_DAYS > 0:
        floor = (dt.date.today() - dt.timedelta(days=KEEP_DAYS)).isoformat()
        for d in [x for x in days if x < floor]:
            del days[d]

    alld = sorted(days)
    payload = {
        "updated": dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
        "from": alld[0] if alld else since,
        "to": alld[-1] if alld else until,
        "window": {"from": since, "to": until},   # что перетягивали в этот раз
        "source": "ДоброДел · ЕЦУР · МинЖКХ · все статусы · накопительный",
        "kinds": ["hvs", "vo", "gvs"],
        "days": days,
    }
    tmp = WATER_JSON.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    tmp.replace(WATER_JSON)
    return hit, len(days), len(fresh), was


def write_data_js(rows):
    meta = {
        "file": "ДоброДел · МинЖКХ (активные, срок с сегодня)",
        "generated": dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
    }
    payload = (
        "/* Автоматически сгенерировано kliker.py — не редактировать вручную. */\n"
        "window.PRELOADED_META = " + json.dumps(meta, ensure_ascii=False) + ";\n"
        "window.PRELOADED_ROWS = " + json.dumps(rows, ensure_ascii=False) + ";\n"
    )
    DATA_JS.write_text(payload, encoding="utf-8")


def main():
    existing = load_water()
    wins = pull_windows(existing)
    have = len((existing or {}).get("days") or {})
    log(f"▶ Кликер запущен. В своде {have} дн. истории. Окон к запросу: {len(wins)}.")

    WATER_JSON.parent.mkdir(parents=True, exist_ok=True)
    pending_empty = []
    confirmed = False
    completed = 0

    def save(a, b, recs):
        hit, total, ndays, _ = write_water_daily(build_rows(recs), a, b, load_water())
        progress(f'Сохранён свод {a} — {b}: {hit} жалоб по воде за {ndays} дн.')

    with DobrodelClient(**get_credentials()) as client:
        client.deadline = time.monotonic() + 20 * 60
        client.on_progress = report_progress
        progress('Вход в Добродел…')
        client.login()
        progress('Вход подтверждён. Получаю данные операционного отчёта Добродела…')
        for a, b, why in wins:
            progress(f'Добродел: окно {completed + 1} из {len(wins)}, {a} — {b} ({why}).')
            recs = client.fetch_all({
                'filters.curators': CURATOR,
                'filters.statuses': STATUSES,
                'filters.createdAfter': a.isoformat(),
                # Portal dates are midnight boundaries; include the whole last day.
                'filters.createdBefore': (b + dt.timedelta(days=1)).isoformat(),
            })
            recs = [r for r in recs if a.isoformat() <= (day_of(r.get('created')) or '') <= b.isoformat()]
            # Save each complete interval immediately. A later historical export
            # failure cannot discard the already downloaded fresh days.
            if recs:
                confirmed = True
            if confirmed:
                for empty_a, empty_b in pending_empty:
                    save(empty_a, empty_b, [])
                pending_empty.clear()
                save(a.isoformat(), b.isoformat(), recs)
            else:
                pending_empty.append((a.isoformat(), b.isoformat()))
            completed += 1

    if not completed:
        raise RuntimeError('Не задано окно для обновления свода жалоб.')
    if not confirmed:
        raise RuntimeError('Добродел вернул пустой отчёт за все запрошенные периоды; прежний свод сохранён.')
    was0 = have
    final = load_water() or {}
    nd = len(final.get("days") or {})
    log(f"✅ Свод по воде: история {was0} → {nd} дн. ({nd - was0:+d}), "
        f"{final.get('from')} — {final.get('to')} → {WATER_JSON.name}")

    if True:  # Embedded module never opens an OS browser.
        log("(тестовый режим: дашборд не открываю)")
    elif HTML.exists():
        log("🌐 Открываю дашборд…")
        pass
    else:
        log(f"⚠ Не найден HTML: {HTML.name}. data.js обновлён — открой дашборд вручную.")


if __name__ == "__main__":
    try:
        main()
    except DobrodelError as error:
        # Fixed diagnostic code only. Never print credentials or a response body.
        log(f'DOBRODEL_ERROR:{error.code}' + (f':{error.status}' if error.status else ''))
        raise SystemExit(1) from None
