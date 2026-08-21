"""Дамп таблиц ZULUGIS и РСО для настройки авто-сбора топ-5 лучших."""
from services.water_dashboard.scraper import scrape_all

ext = scrape_all()
for sid in ["valves", "edo_rso"]:
    d = ext.get(sid) or {}
    print("=" * 60)
    print(sid, "| таблиц:", len(d.get("tables") or []))
    for ti, t in enumerate((d.get("tables") or [])[:3]):
        print(f"--- таблица {ti} ---")
        rows = t if isinstance(t, list) else (t.get("rows") or [])
        for r in rows[:12]:
            print(r)