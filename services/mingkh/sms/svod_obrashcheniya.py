# -*- coding: utf-8 -*-
"""Этап 1 СМС: свод по обращениям МинЖКХ.

С начала года по отчётную дату против АППГ: итог, помесячно и по тематикам.
Данные — портал ЕЦУР через защищённое подключение платформы.
Доступ задаётся администратором в настройках МинЖКХ.

"""
import datetime
import hashlib
import os
import re
from pathlib import Path
from .filters import SOURCES, thu_wed_window
from .client import Pentaho, PortalError
from .context import DATA, atomic_json, json_read, date_range, timestamp, today as local_today, process_lock

# Направление «Водоснабжение/Водоотведение» портала = ровно эти 4 подтемы
# (сверено по снимку 22.09: 169 + 141 + 75 + 13 = 398 = итог направления).
WATER = ["161", "419", "308", "233"]

# Подтемы, которые в свод не входят вовсе — ни в итог, ни в месяцы, ни в
# тематики (решение пользователя 16.10.2026; как на дашборде тепловой карты).
# Сверяем по названию: id «Вне компетенции» в справочнике портала 48,
# а q_top_subtopic отдаёт 49.
EXCLUDED_TITLES = {
    "Вне компетенции (МИНЖКХ)",
    "Платежные документы (ЕИРЦ)",
    "Деятельность УК\\Кооперативов",
    "Прямая линия с населением (МИНЖКХ)",
}

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]

# Фильтры портала, которые ставит свод. Ключ — имя параметра CDA, -1 — без
# ограничения. Справочники: источники — filters.SOURCES, ОМСУ — q_omsu,
# исполнители — q_departments.
#
# «Тип сообщений» всегда «Проблемы, Предложения» = -1: портал берёт только одно
# значение (0 — проблемы, 1 — предложения, «0,1» отвечает 501), а других типов
# в данных нет — 16.10.2026 проверено, что 0 + 1 = -1 и за 2026
# (34 118 + 1116 = 35 234), и за 2025 (47 879 + 1255 = 49 134).
FILTER_KEYS = ("source", "omsu", "department")

# Исполнитель ФКР — для пресета «Жалобы по ФКР».
DEPT_FKR = "1583"
FKR_TITLE = "Фонд капитального ремонта общего имущества многоквартирных домов"

def last_wednesday(today=None):
    today = today or local_today()
    _, end = thu_wed_window(today)
    return today - datetime.timedelta(days=end)


def year_start(d):
    return datetime.date(d.year, 1, 1)


def month_windows(date_from, date_to):
    """Окна по месяцам от date_from до date_to: [(год, месяц, с, по)];
    крайние месяцы обрезаются по датам периода."""
    out = []
    y, m = date_from.year, date_from.month
    while (y, m) <= (date_to.year, date_to.month):
        nxt = datetime.date(y + (m == 12), m % 12 + 1, 1)
        a = max(datetime.date(y, m, 1), date_from)
        b = min(nxt - datetime.timedelta(days=1), date_to)
        out.append((y, m, a, b))
        y, m = nxt.year, nxt.month
    return out


def _year_back(d):
    try:
        return d.replace(year=d.year - 1)
    except ValueError:          # 29 февраля
        return d.replace(year=d.year - 1, day=28)


def _flt(flt):
    """Полный набор фильтров: незаданные — -1."""
    if flt is not None and (not isinstance(flt, dict) or set(flt) - set(FILTER_KEYS)):
        raise ValueError('Неизвестный фильтр свода.')
    result = {k: str((flt or {}).get(k) or '-1') for k in FILTER_KEYS}
    if any(not re.fullmatch(r'-1|[0-9]{1,10}', value) for value in result.values()):
        raise ValueError('Некорректный фильтр свода.')
    return result


# ---------- выгрузка с портала ----------

def fetch_period(p, a, b, today, flt):
    """Обращения МинЖКХ по подтемам за [a, b]: {id: {title, n}}; id -1 — ИТОГО.

    Берётся только curr_summ. АППГ портала (appg_summ) не используем: строку
    подтемы портал отдаёт, лишь если в текущем периоде по ней есть обращения,
    и пустой месяц теряет её прошлый год (16.10 так пропало 378 «Мусор в МКД»).
    Поэтому прошлый год запрашивается как самостоятельный период.

    Портал принимает период смещением в днях назад от сегодня, оба конца
    включительно.
    """
    s, e = (today - a).days, (today - b).days
    params = p.periods(s, e, s + (s - e) + 1, s + 1, **flt)
    cols, rows = p.query("q_top_subtopic", params)
    ix = {c: i for i, c in enumerate(cols)}
    if not {'id', 'title', 'curr_summ'} <= set(ix):
        raise PortalError('Портал вернул неожиданные колонки свода.')
    result = {}
    for row in rows:
        try:
            sid, title, count = str(row[ix['id']]), row[ix['title']], row[ix['curr_summ']]
            count = 0 if count is None else count
            if type(count) not in (int, float) or not 0 <= count <= 10**10 or int(count) != count:
                raise ValueError
            if not isinstance(title, str) or len(title) > 1000:
                raise ValueError
            result[sid] = {'title': title, 'n': int(count)}
        except (IndexError, TypeError, ValueError, OverflowError):
            raise PortalError('Портал вернул некорректные показатели свода.') from None
    return result


def fetch_pair(p, a, b, today, flt):
    return {"curr": fetch_period(p, a, b, today, flt),
            "prev": fetch_period(p, _year_back(a), _year_back(b), today, flt)}


def collect(date_from, date_to, flt=None, log=None):
    """Период [date_from, date_to] против тех же дат год назад: по месяцам и целиком."""
    log = log or (lambda message: None)
    date_from, date_to = date_range(date_from, date_to)
    flt = _flt(flt)
    today = local_today()
    p = Pentaho()
    months = []
    for y, m, a, b in month_windows(date_from, date_to):
        item = {"year": y, "month": m, "from": a.isoformat(), "to": b.isoformat()}
        item.update(fetch_pair(p, a, b, today, flt))
        months.append(item)
        log("  %s: %s %d ok" % (flt, MONTHS[m - 1], y))
    ytd = fetch_pair(p, date_from, date_to, today, flt)       # «ytd» — весь период
    upd, _ = p.last_updated()
    return {"date_from": date_from.isoformat(), "date_to": date_to.isoformat(),
            "filters": flt, "portal_updated": upd,
            "fetched_at": timestamp(),
            "months": months, "ytd": ytd}


def load(date_from, date_to, flt=None, log=None):
    """Выгрузка за период под один набор фильтров — всегда заново с портала
    (решение 16.10.2026: без кэша). Копия кладётся в data/ для истории."""
    flt = _flt(flt)
    date_from, date_to = date_range(date_from, date_to)
    signature = hashlib.sha256('|'.join(flt[k] for k in FILTER_KEYS).encode()).hexdigest()[:20]
    path = os.path.join(DATA, "obrashcheniya_%s_%s_%s.json" % (
        date_from.isoformat(), date_to.isoformat(), signature))
    d = collect(date_from, date_to, flt, log)
    atomic_json(path, d)
    return d


def merge(datasets):
    """Сложить выгрузки по разным источникам: источники не пересекаются,
    у каждого обращения он один (сверено 16.10: сумма 8 источников = все)."""
    base = datasets[0]
    if len(datasets) == 1:
        return base

    def add(dst, src):
        for sid, v in src.items():
            cell = dst.setdefault(sid, {"title": v["title"], "n": 0})
            cell["n"] += v["n"]

    out = {"date_from": base["date_from"], "date_to": base["date_to"], "portal_updated": base["portal_updated"],
           "filters": [d["filters"] for d in datasets],
           "months": [], "ytd": {"curr": {}, "prev": {}}}
    for i, m in enumerate(base["months"]):
        mm = {"year": m["year"], "month": m["month"], "from": m["from"], "to": m["to"], "curr": {}, "prev": {}}
        for d in datasets:
            add(mm["curr"], d["months"][i]["curr"])
            add(mm["prev"], d["months"][i]["prev"])
        out["months"].append(mm)
    for d in datasets:
        add(out["ytd"]["curr"], d["ytd"]["curr"])
        add(out["ytd"]["prev"], d["ytd"]["prev"])
    return out


def build(date_from, date_to, sources=None, omsu="-1", dept="-1", log=None):
    """Данные под выбор за [date_from, date_to]. sources: список id источников,
    пусто/None — все; omsu — id ОМСУ (q_omsu), dept — исполнитель; -1 — без ограничения."""
    date_from, date_to = date_range(date_from, date_to)
    if sources is not None and (not isinstance(sources, (list, tuple)) or len(sources) > len(SOURCES)):
        raise ValueError('Некорректный список источников.')
    if any(str(s) not in dict(SOURCES) for s in (sources or [])):
        raise ValueError('Неизвестный источник обращений.')
    sources = list(dict.fromkeys(str(s) for s in (sources or []) if str(s) != '-1')) or ['-1']
    return merge([load(date_from, date_to, {"source": s, "omsu": omsu, "department": dept}, log)
                  for s in sources])


def omsu_list(refresh=False):
    """Справочник ОМСУ портала [(id, название)] без «Все»; кэшируется в data/."""
    with process_lock(Path(DATA) / '.omsu.lock', timeout=20):
        return _omsu_list_locked(refresh)


def _omsu_list_locked(refresh=False):
    path = os.path.join(DATA, "omsu.json")
    if not refresh and os.path.exists(path):
        try:
            cached = json_read(path, max_bytes=2 * 1024 * 1024)
            if not isinstance(cached, list) or len(cached) > 1000:
                raise ValueError('Некорректный справочник ОМСУ.')
            return _omsu_rows(cached)
        except (ValueError, TypeError):
            pass  # Re-fetch a corrupted dictionary, never use partial names.
    p = Pentaho()
    _, rows = p.query("q_omsu", p.periods(6, 0, 13, 7), timeout=15)
    out = _omsu_rows(rows)
    atomic_json(path, out)
    return out


def _omsu_rows(rows):
    out = []
    for row in rows:
        if not isinstance(row, (list, tuple)) or len(row) < 2:
            raise PortalError('Портал вернул некорректный справочник ОМСУ.')
        ident, title = str(row[0]), row[1]
        if not re.fullmatch(r'-1|[0-9]{1,10}', ident) or not isinstance(title, str) or len(title) > 300:
            raise PortalError('Портал вернул некорректный справочник ОМСУ.')
        if ident != '-1':
            out.append((ident, title))
    return sorted(out, key=lambda item: item[1])


def subtopics(d):
    """Все подтемы, которые есть в выгрузке, кроме исключённых: [(id, название)],
    по убыванию прошлого года."""
    seen = {}
    for key in ("prev", "curr"):
        for sid, v in d["ytd"][key].items():
            if sid != "-1" and v["title"] not in EXCLUDED_TITLES:
                seen.setdefault(sid, [v["title"], 0])[1] += v["n"] if key == "prev" else 0
    return [(sid, t) for sid, (t, _) in sorted(seen.items(), key=lambda x: -x[1][1])]


# ---------- текст СМС ----------

def k(n):
    """27012 -> '27', 2480 -> '2,5', 2000 -> '2' (тысячи, один знак)."""
    v = round(n / 1000.0, 1)
    s = ("%.1f" % v).replace(".", ",")
    return s[:-2] if s.endswith(",0") else s


def pct(curr, appg):
    if not appg:
        return "н/д"
    d = round((curr - appg) * 100.0 / appg)
    return ("+%d%%" % d) if d > 0 else ("−%d%%" % -d if d < 0 else "0%")


def line(curr, appg, unit="тыс"):
    return "%s %s → %s %s (%s)" % (k(appg), unit, k(curr), unit, pct(curr, appg))


def total(rows, ids):
    return sum(rows.get(sid, {}).get("n", 0) for sid in ids)


HEADER = "Информация по жалобам МинЖКХ / —"


def render(d, ids=None):
    """Текст блока. ids — выбранные подтемы (None — все, кроме исключённых).
    В заголовке после «МинЖКХ /» — прочерк, его заполняет сотрудник (решение 16.10.2026)."""
    ids = list(ids) if ids else [sid for sid, _ in subtopics(d)]
    date_to = datetime.date.fromisoformat(d["date_to"])
    date_from = datetime.date.fromisoformat(d.get("date_from") or year_start(date_to).isoformat())
    ytd = date_from == year_start(date_to)
    tot_c = sum(total(m["curr"], ids) for m in d["months"])
    tot_a = sum(total(m["prev"], ids) for m in d["months"])
    if ytd:
        trend = "✅ Снижение по году сохраняется" if tot_c < tot_a else "⚠️ Рост к прошлому году"
        head = "Динамика с начала года по %s" % _day_month(date_to)
    else:
        trend = "✅ Снижение к прошлому году" if tot_c < tot_a else "⚠️ Рост к прошлому году"
        both = date_from.year != date_to.year
        head = "Динамика с %s по %s" % (_day_month(date_from, both), _day_month(date_to, both))

    out = [HEADER, "", head, "", trend,
           "Всего: " + line(tot_c, tot_a), "",
           "По месяцам:"]
    years = len({m.get("year") for m in d["months"]}) > 1     # период через Новый год — с годом
    for m in d["months"]:
        name = MONTHS[m["month"] - 1] + (" %d" % m["year"] if years else "")
        out.append(" %s: %s" % (name, line(total(m["curr"], ids), total(m["prev"], ids))))

    # тематики — за год целиком; порядок — по прошлому году
    names = dict(subtopics(d))
    acc = {sid: (total(d["ytd"]["prev"], [sid]), total(d["ytd"]["curr"], [sid]))
           for sid in ids if sid in names}
    out += ["", "Тематики:"]
    # тематики, которые в обоих годах округляются до 0 тыс, — единичные
    # обращения, строкой «0 тыс → 0 тыс» их не показываем
    shown = [sid for sid in sorted(acc, key=lambda s: acc[s], reverse=True)
             if k(acc[sid][0]) != "0" or k(acc[sid][1]) != "0"]
    for i, sid in enumerate(shown, 1):
        a, c = acc[sid]
        out.append("%d. %s %s" % (i, names[sid], line(c, a, "тыс.")))
    return "\n".join(out)


def _gen_month(m):
    return ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа",
            "сентября", "октября", "ноября", "декабря"][m - 1]


def _day_month(d, with_year=False):
    """«14 октября», с годом — «14 октября 2026»."""
    return "%d %s" % (d.day, _gen_month(d.month)) + (" %d" % d.year if with_year else "")


# ---------- ежедневный доклад (вариант СМС блока 1, только на его вкладке) ----------
# «Динамика к АПП» — к тому же дню неделей раньше (сверено с докладом за 15.10.2026:
# за 08.10 портал даёт разницы −10 / −11 / −11 / +16, как в докладе).
# Темы (решение 16.10.2026): все подтемы МинЖКХ, кроме EXCLUDED_TITLES; в тексте —
# топ-5 подтем дня по числу обращений (как «Топ-10 подтем» на портале), остальные —
# одной строкой «Прочие вопросы». «Поступило» — все подтемы.

DAILY_TOP = 5
# подпись и значок в докладе; подтемы не из списка — с названием портала и значком 🔹
DAILY_LABELS = {
    "419": ("Ржавая вода", "🚰"),
    "308": ("Отсутствие воды", "🚰"),
    "233": ("Напор воды", "🚰"),
    "161": ("Канализация", "📍"),
    "548": ("ФКР", "🏠 "),
}
OTHER = ("Прочие вопросы", "📌")


def daily(day_from, day_to=None):
    """Обращения за период [day_from, day_to] (по умолчанию один день) и за такой же
    период целым числом недель раньше — чтобы дни недели совпадали: день — неделей
    раньше, 2–7 дней — неделей раньше, 8–14 дней — двумя неделями. Топ-5 подтем + прочие."""
    day_to = day_to or day_from
    day_from, day_to = date_range(day_from, day_to, max_days=366)
    p = Pentaho()
    today = local_today()
    shift = datetime.timedelta(days=-(-((day_to - day_from).days + 1) // 7) * 7)
    pf, pt = day_from - shift, day_to - shift
    cur = fetch_period(p, day_from, day_to, today, _flt(None))
    old = fetch_period(p, pf, pt, today, _flt(None))
    subs = {}
    for rows, key in ((cur, "curr"), (old, "prev")):
        for sid, v in rows.items():
            if sid == "-1" or v["title"] in EXCLUDED_TITLES:
                continue
            s = subs.setdefault(sid, {"id": sid, "portal": v["title"], "curr": 0, "prev": 0})
            s[key] = v["n"]
    ranked = sorted(subs.values(), key=lambda s: (-s["curr"], -s["prev"], s["portal"]))
    top, rest = ranked[:DAILY_TOP], ranked[DAILY_TOP:]
    rows = []
    for s in top:
        title, icon = DAILY_LABELS.get(s["id"], (s["portal"], "🔹"))
        rows.append(dict(s, title=title, icon=icon))
    if rest:
        rows.append({"id": ",".join(s["id"] for s in rest), "portal": ", ".join(s["portal"] for s in rest),
                     "title": OTHER[0], "icon": OTHER[1],
                     "curr": sum(s["curr"] for s in rest), "prev": sum(s["prev"] for s in rest)})
    return {"date_from": day_from.isoformat(), "date": day_to.isoformat(),
            "prev_from": pf.isoformat(), "prev": pt.isoformat(), "rows": rows}


def _obr(n):
    a, b = n % 10, n % 100
    return "обращение" if a == 1 and b != 11 else "обращения" if 2 <= a <= 4 and not 12 <= b <= 14 else "обращений"


def _dyn(curr, prev):
    """«снижение на 14% (- 15 обращений)» / «рост на 37% (+16 обращений)»."""
    d = curr - prev
    if d < 0:
        return "снижение на %d%% (- %d %s)" % (round(-d * 100.0 / prev), -d, _obr(-d))
    if prev:
        return "рост на %d%% (+%d %s)" % (round(d * 100.0 / prev), d, _obr(d))
    return "рост (+%d %s)" % (d, _obr(d))


def render_daily(d):
    a = datetime.date.fromisoformat(d.get("date_from") or d["date"])
    b = datetime.date.fromisoformat(d["date"])
    day = (b.strftime("%d.%m.%Y") if a == b
           else "период с %s по %s" % (a.strftime("%d.%m" if a.year == b.year else "%d.%m.%Y"), b.strftime("%d.%m.%Y")))
    cur = sum(r["curr"] for r in d["rows"])
    old = sum(r["prev"] for r in d["rows"])
    out = ["Уважаемый Кирилл Борисович, докладываю информацию по ключевым вопросам в части "
           "водоснабжения, водоотведения и ФКР за %s" % day, "",
           "🔷Поступило: %d %s;" % (cur, _obr(cur)),
           "🔷Динамика к АПП: %s;" % ("динамика отсутствует" if cur == old else _dyn(cur, old)), ""]
    for r in d["rows"]:
        if r["curr"] == r["prev"]:
            out.append("%s%s: динамика отсутствует." % (r["icon"], r["title"]))
        else:
            out.append("%s%s: %d - %s" % (r["icon"], r["title"], r["curr"], _dyn(r["curr"], r["prev"])))
    return "\n".join(out)
