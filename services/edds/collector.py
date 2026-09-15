# -*- coding: utf-8 -*-
"""
КЛИКЕР для дашборда «Контроль жалоб ЕЦУР».

Что делает сам, без рук:
  1. Заходит на admin.vmeste.mosreg.ru под сохранённой сессией
     (если сессия протухла — логинится паролем из хранилища Windows).
  2. Тянет операционный отчёт МинЖКХ (активные статусы, срок решения с сегодня)
     напрямую через API — постранично, без Excel.
  3. Складывает данные в data.js рядом с HTML.
  4. Открывает дашборд в браузере.

Пароль нигде в файлах не хранится — только в Диспетчере учётных данных Windows.
Обновить логин/пароль: запусти настройки в админ-панели.
"""

import os
import re
import sys
import html
import json
import datetime as dt
from pathlib import Path

from services.auth.integrations import credentials
from playwright.sync_api import sync_playwright

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
# Потолок ОДНОГО добора назад. Держим порцию небольшой намеренно: 550 дней
# одним запросом — это ~35 тыс. записей, прогон не укладывается в таймаут и
# пропадает целиком. Порциями свод дорастает до нужной глубины за несколько
# прогонов, и каждый из них завершается и сохраняется.
MAX_CATCHUP = int(os.environ.get("KLIKER_CATCHUP", "180"))
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


def get_credentials():
    value = credentials()
    if not value:
        raise RuntimeError("Администратор должен настроить логин и пароль Добродела")
    return value['username'], value['password']


def logged_in(page):
    """На странице отчёта и не выкинуло на логин."""
    if "/login" in page.url.lower():
        return False
    try:
        return page.locator("#curatorSelect").count() > 0
    except Exception:
        return False


def do_login(page):
    email, pwd = get_credentials()
    log("🔑 Сессия недоступна — вхожу под сохранённым паролем…")
    page.goto(LOGIN_URL, wait_until="networkidle", timeout=60000)
    # форма Spring Security: j_username / j_password
    page.fill("input[name=j_username]", email)
    page.fill("input[name=j_password]", pwd)
    try:  # «запомнить меня» — сессия живёт дольше
        page.check("#_spring_security_remember_me")
    except Exception:
        pass
    # надёжный сабмит: Enter в поле пароля, при неудаче — нативная отправка формы
    page.press("input[name=j_password]", "Enter")
    try:
        page.wait_for_url(lambda u: "/login" not in u, timeout=30000)
    except Exception:
        try:
            page.eval_on_selector("form", "f => f.submit()")
            page.wait_for_url(lambda u: "/login" not in u, timeout=30000)
        except Exception:
            pass


def ensure_session(browser):
    ctx = browser.new_context(storage_state=None,
                              accept_downloads=False)
    page = ctx.new_page()
    page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=60000)
    if not logged_in(page):
        do_login(page)
        page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=60000)
        if not logged_in(page):
            log("❌ Не удалось войти. Проверь логин/пароль (настройки в админ-панели) "
                "или капчу/2FA на портале.")
            sys.exit(1)
    # Session stays in memory, never in a public file.
    return ctx, page


FETCH_JS = """async (args) => {
    const [cur, filters, pageIdx, size] = args;
    const base = "filters.curators=" + encodeURIComponent(cur) + "&" + filters;
    const url = "/report/operative?orderBy=ID&page=" + pageIdx + "&size=" + size + "&" + base;
    const r = await fetch(url, {headers:{Accept:'application/json'}});
    if (!r.ok) return {error: r.status};
    return {rows: await r.json()};
}"""


def fetch_all(page, filters):
    """filters — готовая строка вида «filters.statuses=…&filters.createdAfter=…»."""
    all_rows = []
    idx = 0
    while True:
        res = page.evaluate(FETCH_JS, [CURATOR, filters, idx, PAGE_SIZE])
        if isinstance(res, dict) and res.get("error"):
            log(f"❌ Сервер вернул ошибку {res['error']} на странице {idx}.")
            sys.exit(1)
        batch = res.get("rows") or []
        if not batch:
            break
        all_rows += batch
        log(f"   …страница {idx + 1}: +{len(batch)} (всего {len(all_rows)})")
        if len(batch) < PAGE_SIZE:
            break
        idx += 1
    return all_rows


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

    if last is None:
        return [(want_from, until, f"первый сбор, {DAYS_BACK} дн.")]

    wins = []
    since = last - dt.timedelta(days=OVERLAP)
    floor = until - dt.timedelta(days=MAX_CATCHUP)
    if since < floor:
        log(f"⚠ Свод не обновлялся с {last}. Беру последние {MAX_CATCHUP} дн., "
            f"в истории останется разрыв — при необходимости увеличьте KLIKER_CATCHUP.")
        since = floor
    wins.append((since, until, f"хвост {(until - since).days} дн., нахлёст {OVERLAP}"))

    if first is not None and want_from < first:
        back_to = first - dt.timedelta(days=1)
        back_from = max(want_from, back_to - dt.timedelta(days=MAX_CATCHUP))
        wins.append((back_from, back_to,
                     f"добор назад {(back_to - back_from).days + 1} дн. "
                     f"до глубины {DAYS_BACK} дн."))
    return wins


def write_water_daily(rows, since, until, existing):
    """rows — сетка с HEADER в первой строке за окно [since, until].
    Свод накопительный: свежее окно заменяем целиком, всё что вне его —
    оставляем как накопилось. День → ОМСУ → [ХВС, водоотведение, ГВС]."""
    fresh = {}
    hit = 0
    for r in rows[1:]:
        kind = water_kind(r[4], r[5], r[3])
        if not kind:
            continue
        d = day_of(r[8])
        if not d:
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

    pulled = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            ctx, page = ensure_session(browser)
            log("✅ Сессия активна.")

            # 1) для дашборда жалоб — открытые заявки со сроком от сегодня.
            #    Без окна по дате подачи: открытая жалоба двухлетней давности
            #    тоже должна быть видна, она никуда не делась.
            log("   ── открытые жалобы для data.js …")
            act_recs = fetch_all(page, "filters.statuses=" + ACTIVE_STATUSES +
                                 "&filters.deadlineAfter=" + dt.date.today().isoformat())

            # 2) для графика ЕДДС — все статусы за окна подачи
            for a, b, why in wins:
                log(f"   ── свод по воде: {a} — {b} ({why}) …")
                recs = fetch_all(page, "filters.statuses=" + STATUSES +
                                 "&filters.createdAfter=" + a.isoformat() +
                                 "&filters.createdBefore=" + b.isoformat())
                pulled.append((a.isoformat(), b.isoformat(), recs))
        finally:
            browser.close()

    if not act_recs and not any(r for _, _, r in pulled):
        log("⚠ Сервер вернул 0 записей по всем запросам. Файлы не тронуты.")
        sys.exit(1)

    if act_recs:
        arows = build_rows(act_recs)
        write_data_js(arows)
        log(f"✅ Открытых жалоб: {len(arows) - 1}, "
            f"{len({r[2] for r in arows[1:]})} ОМСУ → {DATA_JS.name}")
    else:
        log("⚠ По открытым жалобам пришло 0 записей — data.js не тронут.")

    was0 = have
    for a, b, recs in pulled:
        if not recs:
            log(f"⚠ Окно {a} — {b}: 0 записей, свод не тронут.")
            continue
        hit, total, ndays, _ = write_water_daily(build_rows(recs), a, b, load_water())
        log(f"   свод {a} — {b}: +{hit} жалоб за {ndays} дн.")
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
    main()
