"""
Автономный запуск выгрузки ЦДС для планировщика.

По умолчанию выгружает данные за вчерашний день.
Период можно переопределить:
  - аргументами:  python -m services.cds.runner 01.01.2025 07.01.2025
  - переменными:  CDS_DATE_FROM / CDS_DATE_TO
"""
import asyncio
import os
import sys
from datetime import date, timedelta

from services.cds.scraper import scrape_cds_appeals


def _default_period():
    yesterday = date.today() - timedelta(days=1)
    d = yesterday.strftime("%d.%m.%Y")
    return d, d


def main():
    if len(sys.argv) > 2:
        date_from, date_to = sys.argv[1], sys.argv[2]
    else:
        date_from = os.environ.get("CDS_DATE_FROM", "")
        date_to = os.environ.get("CDS_DATE_TO", "")
        if not date_from or not date_to:
            date_from, date_to = _default_period()

    print(f"[cds.runner] Выгрузка за период {date_from} — {date_to}", flush=True)

    result = asyncio.run(
        scrape_cds_appeals(date_from, date_to, headless=True)
    )

    if result.get("success"):
        print(f"[cds.runner] ✅ Выгружено {result.get('count', 0)} обращений", flush=True)
        sys.exit(0)
    else:
        print(f"[cds.runner] ❌ Ошибка: {result.get('error')}", flush=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
