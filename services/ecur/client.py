# -*- coding: utf-8 -*-
import os
import html
import datetime as dt
import threading
from pathlib import Path
from urllib.parse import urljoin
from html.parser import HTMLParser

import requests
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent.parent
env_file = BASE_DIR / ".env"
if env_file.exists():
    load_dotenv(dotenv_path=env_file)
else:
    load_dotenv()

BASE = "https://admin.vmeste.mosreg.ru"
LOGIN_PAGE = BASE + "/login"
REPORT_PAGE = BASE + "/OperativeReportGenerating"
REPORT_API = BASE + "/report/operative"

CURATOR = os.getenv("ECUR_CURATOR", "Министерство жилищно-коммунального хозяйства Московской области").strip()
STATUSES = os.getenv("ECUR_STATUSES", "37,32,50,54,57,51,511,512").strip()
DISTRICTS = os.getenv("ECUR_DISTRICTS", "").strip()
PAGE_SIZE = int(os.getenv("ECUR_PAGE_SIZE") or "500")

# РОВНО 11 КОЛОНОК (КАК В KLIKER.PY)
HEADER = [
    "Номер", "Адрес", "Район", "Категория ЕЦУР", "Подкатегория ЕЦУР",
    "ЕЦУР факт", "Исполнитель", "Текст", "Дата создания", "Срок", "Статус"
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

STATE = {
    "session": None,
    "email": None,
    "password": None,
    "rows": None,
    "meta": None,
}
LOCK = threading.Lock()


class _FormParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_form = False
        self.action = None
        self.hidden = {}

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "form":
            names = " ".join(str(v or "").lower() for v in a.values())
            if self.action is None or "login" in names or "security" in names:
                self.action = a.get("action") or self.action
                self.in_form = True
        elif tag == "input" and self.in_form:
            if (a.get("type") or "").lower() == "hidden" and a.get("name"):
                self.hidden[a["name"]] = a.get("value", "")

    def handle_endtag(self, tag):
        if tag == "form":
            self.in_form = False


def _looks_like_login(text: str, url: str) -> bool:
    return ("j_password" in text) or ("/login" in url.lower())


def dobrodel_login(email: str, password: str):
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    try:
        r = s.get(LOGIN_PAGE, timeout=30)
    except requests.RequestException as e:
        return None, f"Портал недоступен: {e.__class__.__name__}. Проверьте сеть/VPN."

    p = _FormParser()
    try:
        p.feed(r.text)
    except Exception:
        pass
    action = urljoin(r.url, p.action) if p.action else urljoin(BASE, "/j_spring_security_check")

    data = dict(p.hidden)
    data.update({
        "j_username": email,
        "j_password": password,
        "_spring_security_remember_me": "on",
    })
    try:
        r2 = s.post(action, data=data, timeout=30, allow_redirects=True, headers={"Referer": r.url})
    except requests.RequestException as e:
        return None, f"Сбой при входе: {e.__class__.__name__}."

    if _looks_like_login(r2.text, r2.url) and action != urljoin(BASE, "/login"):
        try:
            r2 = s.post(urljoin(BASE, "/login"), data=data, timeout=30, allow_redirects=True, headers={"Referer": r.url})
        except requests.RequestException:
            pass

    if _looks_like_login(r2.text, r2.url):
        return None, "Неверный логин или пароль от ДоброДела."

    try:
        r3 = s.get(REPORT_PAGE, timeout=30)
        if _looks_like_login(r3.text, r3.url):
            return None, "Вход не подтвердился (портал вернул страницу логина)."
    except requests.RequestException as e:
        return None, f"Вход прошёл, но портал не отвечает: {e.__class__.__name__}."

    return s, None


def clean(v):
    if v is None:
        return ""
    return html.unescape(str(v)).strip()


def fetch_report(s: requests.Session):
    all_recs, idx = [], 0
    today = dt.date.today().isoformat()

    while True:
        params = {
            "orderBy": "ID",
            "page": idx,
            "size": PAGE_SIZE,
            "filters.curators": CURATOR,
            "filters.statuses": STATUSES,
            "filters.deadlineAfter": today,
        }
        if DISTRICTS:
            params["filters.districts"] = DISTRICTS

        try:
            r = s.get(REPORT_API, params=params, timeout=60, headers={
                "Accept": "application/json",
                "X-Requested-With": "XMLHttpRequest",
                "Referer": REPORT_PAGE,
            })
        except requests.RequestException as e:
            return None, f"Портал не ответил на выгрузку: {e.__class__.__name__}."

        if _looks_like_login(r.text[:2000], r.url):
            return None, "SESSION_EXPIRED"
        if r.status_code != 200:
            return None, f"Сервер вернул {r.status_code} на стр. {idx}."

        try:
            batch = r.json()
        except ValueError:
            return None, "Портал вернул не-JSON."

        if isinstance(batch, dict):
            batch = (batch.get("rows") or batch.get("content")
                     or batch.get("data") or batch.get("items") or [])

        if not batch:
            break

        all_recs += batch
        print(f"[ECUR] …страница {idx + 1}: +{len(batch)} (всего {len(all_recs)})", flush=True)
        if len(batch) < PAGE_SIZE:
            break
        idx += 1

    if not all_recs:
        return None, "Портал вернул 0 записей под выбранный фильтр."

    # РОВНО 11 ЭЛЕМЕНТОВ В СТРОКЕ
    rows = [HEADER]
    for rec in all_recs:
        rows.append([
            rec.get("cardId", ""),
            clean(rec.get("address")),
            clean(rec.get("district")),
            clean(rec.get("ecurCategory")),
            clean(rec.get("subcategory")),
            clean(rec.get("ecurFact")) or "—",
            clean(rec.get("org")),
            clean(rec.get("body")),
            clean(rec.get("created")),
            clean(rec.get("deadline")),
            clean(rec.get("status")),
        ])
    return rows, None


def authenticate_user(email: str, password: str):
    email = (email or "").strip()
    password = (password or "").strip()

    if not email or not password:
        return False, "Не указан email или пароль."

    print(f"[ECUR] Вход на ДоброДел для {email}…", flush=True)
    s, err = dobrodel_login(email, password)
    if err:
        return False, err

    rows, err = fetch_report(s)
    if err:
        return False, f"Вход выполнен, но ошибка при получении свода: {err}"

    meta = {
        "file": "ДоброДел · МинЖКХ (активные, срок с сегодня)",
        "generated": dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
    }

    with LOCK:
        STATE["session"] = s
        STATE["email"] = email
        STATE["password"] = password
        STATE["rows"] = rows
        STATE["meta"] = meta

    return True, len(rows) - 1


def refresh_data():
    with LOCK:
        s = STATE["session"]
        email = STATE["email"]
        password = STATE["password"]

    if not s or not email or not password:
        return False, "Нет активной сессии. Пожалуйста, войдите снова."

    rows, err = fetch_report(s)

    if err == "SESSION_EXPIRED":
        print("[ECUR] Сессия протухла — повторный вход…", flush=True)
        s_new, lerr = dobrodel_login(email, password)
        if lerr:
            return False, "Сессия истекла, повторный вход не удался: " + lerr

        with LOCK:
            STATE["session"] = s_new

        rows, err = fetch_report(s_new)

    if err:
        return False, err

    meta = {
        "file": "ДоброДел · МинЖКХ (активные, срок с сегодня)",
        "generated": dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
    }

    with LOCK:
        STATE["rows"] = rows
        STATE["meta"] = meta

    return True, len(rows) - 1


def get_current_data():
    with LOCK:
        return {
            "rows": STATE["rows"],
            "meta": STATE["meta"],
            "email": STATE["email"],
            "is_authed": STATE["session"] is not None
        }


def clear_session():
    with LOCK:
        STATE["session"] = None
        STATE["email"] = None
        STATE["password"] = None
        STATE["rows"] = None
        STATE["meta"] = None
        