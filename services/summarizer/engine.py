"""Сумматор: GigaChat + Qwen (с выбором движка) + алгоритмический фолбэк.

Переменные .env:
  SUMMARIZER_BACKEND: qwen | gigachat | algo
  SUMMARIZER_GIGACHAT_CREDENTIALS (fallback GIGACHAT_CREDENTIALS)
  SUMMARIZER_QWEN_API_KEY (fallback QWEN_API_KEY)
  QWEN_API_BASE, QWEN_MODEL
  GIGACHAT_MODEL
"""
import os
import re
import datetime
from collections import Counter, OrderedDict


# ═══════════════ КЛИЕНТЫ ═══════════════
_giga_cache = {}
GIGA_FALLBACKS = ["GigaChat-2-Pro", "GigaChat-2-Max", "GigaChat-2", "GigaChat-3-Ultra"]


def _giga_client(credentials):
    if not credentials:
        raise RuntimeError("GigaChat credentials не заданы в .env")
    if credentials not in _giga_cache:
        from gigachat import GigaChat
        _giga_cache[credentials] = GigaChat(credentials=credentials, verify_ssl_certs=False)
    return _giga_cache[credentials]


def _get_giga():
    return _giga_client(os.getenv("GIGACHAT_CREDENTIALS", "").strip())


def _gigachat_chat(messages, creds=None, model=None, max_tokens=2000):
    giga = _giga_client((creds or os.getenv("GIGACHAT_CREDENTIALS", "")).strip())
    preferred = (model or os.getenv("GIGACHAT_MODEL", "GigaChat-2-Pro")).strip()
    models = [preferred] + [m for m in GIGA_FALLBACKS if m != preferred]
    last = None
    for m in models:
        try:
            resp = giga.chat({"model": m, "messages": messages,
                              "temperature": 0.4, "max_tokens": max_tokens})
            return resp.choices[0].message.content
        except Exception as e:
            last = e
            if "No such model" in str(e) or "404" in str(e):
                continue
            raise
    raise last or RuntimeError("GigaChat: модель недоступна")


def _qwen_chat(messages, max_tokens=2000, base=None, key=None, model=None):
    import httpx
    base = (base or os.getenv("QWEN_API_BASE", "")).strip().rstrip("/")
    key = (key or os.getenv("QWEN_API_KEY", "")).strip()
    model = (model or os.getenv("QWEN_MODEL", "qwen3.8-27b-fp8")).strip()
    if not base:
        raise RuntimeError("QWEN_API_BASE не задан в .env")
    headers = {"Content-Type": "application/json"}
    if key:
        headers["Authorization"] = f"Bearer {key}"
    paths = ["/v1/chat/completions", "/chat/completions", "/api/v1/chat/completions"]
    if base.endswith("/chat/completions"):
        paths = [""]
    last = None
    for p in paths:
        try:
            r = httpx.post(base + p,
                           json={"model": model, "messages": messages,
                                 "temperature": 0.4, "max_tokens": max_tokens},
                           headers=headers, timeout=120, verify=False)
            if r.status_code == 404:
                last = RuntimeError(f"404 на {base + p}")
                continue
            r.raise_for_status()
            return r.json()["choices"][0]["message"]["content"]
        except Exception as e:
            if "404" in str(e):
                last = e
                continue
            raise
    raise last or RuntimeError("Qwen: эндпоинт не найден")


SYSTEM_PROMPT = """Ты — старший оперативный дежурный ЖКХ-Центра Московской области.
Из полученной переписки собери краткий оперативный отчёт СТРОГО в такой структуре:

По информации оперативного дежурного ЖКХ-Центра в ЧЧ:ММ произошло [тип события] [ресурсы] в [место]:

Причина: ...

Затронуло: N МКД, N СЗО (перечисление).

Общий охват жителей: N чел.

На месте бригада [организация] в составе N чел. и N ед. техники.

План. срок завершения работ: ЧЧ:ММ

Информирование жителей проведено.

[Должности с ФИО] проинформированы.

Старший оперативный дежурный ЖКХ-Центра
Источник:
[список авторов сообщений]

Правила: не выдумывай факты; если данных нет — пиши "уточняется";
время в формате 24ч; без лишних слов и воды; сохраняй официально-деловой стиль."""


# ═══════════════ СЛОВАРИ И РЕГУЛЯРКИ ═══════════════
RESOURCES = ["ХВС", "ГВС", "отоплени", "электроснабжен", "электричеств",
             "газоснабжен", "водоснабжен", "теплоснабжен"]
EVENTS = [
    ("аварийное отключение", ["аварийн"]),
    ("отключение", ["отключен", "прекращен"]),
    ("утечка", ["утечк"]),
    ("прорыв", ["прорыв"]),
    ("авария", ["авари"]),
    ("плановое отключение", ["планов"]),
    ("снижение давления", ["давлен", "напор"]),
]
PLACES_RE = re.compile(
    r"(?:г\.о\.|г\.|мкр\.|мкр|п\.|пос\.|д\.|с\.)\s*[А-ЯЁ][А-Яа-яё\-\.]*"
    r"(?:\s*,?\s*(?:ул\.|улица|проспект|пр-кт|пер\.|ш\.|пл\.|мкр\.|д\.|д|дом)\s*[А-Яа-яё0-9\-\.]+(?:\s*,?\s*д\.?\s*\d+[а-яА-Я]?)?)?",
)
TIME_RE = re.compile(r"\b([01]?\d|2[0-3]):([0-5]\d)\b")
DIGITS_RE = re.compile(r"\b(\d[\d\s]*)\b")
TEAM_ORG_RE = re.compile(
    r"(МУП|АО|ООО|ОАО|ГУП|ГБУЗ|ГБОУ)\s*«[^»]+»|[А-ЯЁ][А-Яа-яё\-]+\s+(?:МУП|АО|ООО)",
    re.IGNORECASE,
)
POSITIONS = [
    "министр", "первый замминистра", "замминистра", "заместитель министра",
    "глава", "губернатор", "начальник", "директор", "главный инженер",
    "заместитель директора", "заместитель главного", "зам. министра",
]
AUTHOR_RE = re.compile(r"^([А-ЯЁA-Z][а-яёa-z\-]+(?:\s[А-ЯЁA-Z][а-яёa-z\-]+)?(?:\s[А-ЯЁ][а-яё]+)?)\s*:\s*", re.MULTILINE)
COVERAGE_RE = re.compile(r"(\d[\d\s\.]*)\s*(?:МКД|многоквартирн)", re.IGNORECASE)
SZO_RE = re.compile(r"(\d+)\s*СЗО", re.IGNORECASE)
PEOPLE_RE = re.compile(r"(\d[\d\s\.]*)\s*(?:чел\.?|человек|жител)", re.IGNORECASE)
BRIGADE_RE = re.compile(r"(\d+)\s*чел\.?", re.IGNORECASE)
TECH_RE = re.compile(r"(\d+)\s*ед\.\s*техник", re.IGNORECASE)
DEADLINE_RE = re.compile(r"(?:до|к|план[^\n]{0,30})\s*([01]?\d|2[0-3]):([0-5]\d)", re.IGNORECASE)


# ═══════════════ ЭКСТРАКТОРЫ ═══════════════
def _clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def _clean_text(text):
    text = re.sub(r"[🛁🚰👥🏢🏪🕑📅☎📍⚡⏳✅❌🔄📋📄📨🗜🛎📆📊📈📉📌📎📝🔍🔎💡⚠️🚨🔥❄️💧🌡️🏠🏘️🏗️]", "", text)
    text = re.sub(r"ХРОНОЛОГИЯ.*?(?=Заявка:|Источник:|$)", "", text, flags=re.IGNORECASE | re.DOTALL)
    return text


def _extract_time(text):
    m = re.search(r"(?:в|около|примерно)\s+([01]?\d|2[0-3]):([0-5]\d)", text)
    if m:
        return f"{m.group(1).zfill(2)}:{m.group(2)}"
    m = re.search(r"Начал[оа][^\n]*?([01]?\d|2[0-3]):([0-5]\d)", text, re.IGNORECASE)
    if m:
        return f"{m.group(1).zfill(2)}:{m.group(2)}"
    matches = TIME_RE.findall(text)
    if matches:
        return f"{matches[0][0].zfill(2)}:{matches[0][1]}"
    return datetime.datetime.now().strftime("%H:%M")


def _extract_event(text):
    low = text.lower()
    for label, kws in EVENTS:
        if any(k in low for k in kws):
            return label
    return "инцидент"


STOP_PLACES = {"ХВС", "ГВС", "МКД", "СЗО", "МО", "РФ", "АО", "ООО", "МУП", "ВС"}


def _extract_resources(text):
    low = text.lower()
    found = []
    mapping = {
        "хвс": "ХВС", "гвс": "ГВС",
        "отоплени": "отопления", "электроснабжен": "электроснабжения",
        "электричеств": "электроснабжения", "газоснабжен": "газоснабжения",
        "водоснабжен": "водоснабжения", "теплоснабжен": "теплоснабжения",
    }
    for key, val in mapping.items():
        if key in low:
            found.append(val)
    return list(OrderedDict.fromkeys(found)) or ["ресурсов"]


def _extract_place(text):
    for m in PLACES_RE.finditer(text):
        cand = _clean(m.group(0))
        last = cand.split()[-1] if cand.split() else ""
        if last.upper() in STOP_PLACES:
            continue
        return cand
    return "Московской области"


def _extract_cause(text, event):
    sents = re.split(r"[.!?]\s+", text)
    for s in sents:
        if re.match(r"^(какая|какой|какие|почему|что|кто|когда|где)\b", s.strip(), re.IGNORECASE):
            continue
        if re.search(r"причин[аы]|связан|из-за|вследствие", s, re.IGNORECASE):
            cleaned = _clean(s)
            if "|" in cleaned or re.search(r"\d{4,}", cleaned):
                continue
            return cleaned
    for s in sents:
        if any(k in s.lower() for k in ["утечк", "прорыв", "отключен", "авари"]):
            cleaned = _clean(s)
            if "|" in cleaned or re.search(r"\d{4,}", cleaned):
                continue
            return cleaned
    return f"{event} (причина уточняется)"


def _extract_affected(text):
    parts = []
    m_mkd = COVERAGE_RE.search(text)
    if m_mkd:
        parts.append(f"{m_mkd.group(1).strip()} МКД")
    m_szo = SZO_RE.search(text)
    if m_szo:
        n = int(m_szo.group(1))
        szo = []
        for kw, label in [("школ", "школа"), ("детск", "детский сад"),
                          ("дом культуры", "дом культуры"), ("больниц", "больница"),
                          ("поликлин", "поликлиника"), ("сад", "детский сад")]:
            if kw in text.lower():
                szo.append(label)
        szo = list(OrderedDict.fromkeys(szo))[:n]
        if szo:
            parts.append(f"{n} СЗО ({', '.join(szo)})")
        else:
            parts.append(f"{n} СЗО")
    return ", ".join(parts) if parts else "объекты уточняются"


def _extract_coverage(text):
    m = PEOPLE_RE.search(text)
    if m:
        num = m.group(1).replace(" ", "").replace(".", "")
        try:
            return f"{int(num):,}".replace(",", " ") + " чел."
        except Exception:
            pass
    for phrase in ["общий охват", "затронут", "жителей", "населени"]:
        idx = text.lower().find(phrase)
        if idx != -1:
            tail = text[idx:idx+80]
            m2 = DIGITS_RE.search(tail)
            if m2:
                try:
                    n = int(m2.group(1).replace(" ", ""))
                    if n > 50:
                        return f"{n:,}".replace(",", " ") + " чел."
                except Exception:
                    pass
    return "—"


def _extract_team(text):
    lines = re.split(r"[.\n]", text)
    team_lines = [l for l in lines if TEAM_ORG_RE.search(l)
                  and re.search(r"бригад|выехал|направлен|на мест|работает", l, re.IGNORECASE)]
    if not team_lines:
        return None
    line = team_lines[0]
    org = TEAM_ORG_RE.search(line).group(0)
    people = BRIGADE_RE.search(line)
    tech = TECH_RE.search(line)
    parts = [f"бригада {org}"]
    if people:
        parts.append(f"в составе {people.group(1)} чел.")
    if tech:
        parts.append(f"и {tech.group(1)} ед. техники")
    return ". ".join(parts).rstrip(". ")


def _extract_deadline(text):
    m = DEADLINE_RE.search(text)
    if m:
        return f"{m.group(1).zfill(2)}:{m.group(2)}"
    for phrase in ["завершени", "восстановлени", "ликвидац"]:
        idx = text.lower().find(phrase)
        if idx != -1:
            tail = text[idx:idx+80]
            m2 = TIME_RE.search(tail)
            if m2:
                return f"{m2.group(1).zfill(2)}:{m2.group(2)}"
    return None


def _extract_informed(text):
    found = []
    for p in POSITIONS:
        pattern = re.compile(re.escape(p) + r"[а-яё]*\s*[а-яё]*", re.IGNORECASE)
        for m in pattern.finditer(text):
            chunk = text[max(0, m.start()-5):min(len(text), m.end()+40)]
            chunk = re.sub(r"\s+", " ", chunk).strip()
            if "проинформирован" in text[max(0,m.start()-50):m.end()+80].lower() \
               or "уведомлен" in text[max(0,m.start()-50):m.end()+80].lower() \
               or any(k in text[m.start():m.end()+80].lower() for k in ["министр", "замминистр", "зам."]):
                cleaned = re.sub(r"[,.;:]+$", "", chunk).strip()
                if cleaned and len(cleaned) < 80:
                    found.append(cleaned)
    return list(OrderedDict.fromkeys(found))[:6]


STOP_AUTHORS = {"причина", "заявка", "исполнитель", "адрес", "объект", "описание",
                "комментарий", "статус", "источник", "примечание", "хронология"}


def _extract_authors(text):
    authors = []
    for m in AUTHOR_RE.finditer(text):
        name = m.group(1).strip()
        if name.lower() not in STOP_AUTHORS and len(name) > 1:
            authors.append(name)
    return list(OrderedDict.fromkeys(authors))


def _extract_info_provided(text):
    low = text.lower()
    return any(k in low for k in ["информир", "оповещ", "жителям сообщ", "рассылк"])


# ═══════════════ АЛГОРИТМИЧЕСКИЙ СБОРЩИК ═══════════════
def _summarize_algo(text: str) -> dict:
    time = _extract_time(text)
    event = _extract_event(text)
    place = _extract_place(text)
    cause = _extract_cause(text, event)
    resources = _extract_resources(text)
    affected = _extract_affected(text)
    coverage = _extract_coverage(text)
    team = _extract_team(text)
    deadline = _extract_deadline(text)
    informed = _extract_informed(text)
    authors = _extract_authors(text)
    info_provided = _extract_info_provided(text)

    res_line = " и ".join(resources) if resources else "ресурсов"
    lines = [
        f"По информации оперативного дежурного ЖКХ-Центра в {time} произошло "
        f"{event} {res_line} в {place}:",
        "",
        f"Причина: {cause}",
        "",
        f"Затронуло: {affected}.",
        "",
    ]
    if coverage and coverage != "—":
        lines.append(f"Общий охват жителей: {coverage}")
        lines.append("")
    if team:
        lines.append(f"На месте {team}.")
        lines.append("")
    if deadline:
        lines.append(f"План. срок завершения работ: {deadline}")
        lines.append("")
    if info_provided:
        lines.append("Информирование жителей проведено.")
        lines.append("")
    if informed:
        lines.append(", ".join(informed) + " проинформированы.")
        lines.append("")
    lines.append("Старший оперативный дежурный ЖКХ-Центра")
    if authors:
        lines.append("Источник:")
        for a in authors[:5]:
            lines.append(a)

    report_text = "\n".join(lines)
    key_points = []
    if not (team or deadline or informed or affected != "объекты уточняются"):
        for ln in text.splitlines():
            ln = ln.strip()
            if ln and len(ln) > 20:
                key_points.append(ln[:200])
                if len(key_points) >= 5:
                    break

    return {
        "ok": True, "backend": "algo", "report_text": report_text,
        "stats": {
            "lines": len(text.splitlines()), "authors": len(authors),
            "words": len(re.findall(r"\S+", text)),
            "compression": round(100 - 100 * len(report_text) / max(1, len(text)), 1),
            "time": time, "event": event, "place": place,
        },
        "key_points": key_points, "facts": [],
        "top_authors": Counter(authors).most_common(5),
        "extracted": {
            "resources": resources, "affected": affected, "coverage": coverage,
            "team": team, "deadline": deadline, "informed": informed, "sources": authors,
        },
    }


# ═══════════════ ТОЧКА ВХОДА ═══════════════
def summarize(text, max_points=7, backend=None, creds=None, qwen_key=None):
    if len(text.strip()) < 30:
        return {"ok": False, "error": "Слишком короткий текст"}
    text = _clean_text(text)
    backend = (backend or os.getenv("SUMMARIZER_BACKEND", "algo")).strip()
    creds = creds or os.getenv("SUMMARIZER_GIGACHAT_CREDENTIALS", "").strip() or None
    qwen_key = qwen_key or os.getenv("SUMMARIZER_QWEN_API_KEY", "").strip() or None

    if backend in ("gigachat", "qwen"):
        try:
            authors = _extract_authors(text)
            msgs = [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": text}]
            if backend == "gigachat":
                ai_text = _gigachat_chat(msgs, creds=creds)
            else:
                ai_text = _qwen_chat(msgs, key=qwen_key)
            return {
                "ok": True, "backend": backend, "report_text": ai_text,
                "stats": {
                    "lines": len(text.splitlines()), "authors": len(authors),
                    "words": len(re.findall(r"\S+", text)),
                    "compression": round(100 - 100 * len(ai_text) / max(1, len(text)), 1),
                    "time": _extract_time(text), "event": _extract_event(text),
                    "place": _extract_place(text),
                },
                "key_points": [], "facts": [],
                "top_authors": Counter(authors).most_common(5), "extracted": {},
            }
        except Exception as e:
            print(f"[summarizer] {backend} error: {e} → fallback на алгоритм", flush=True)
    return _summarize_algo(text)
