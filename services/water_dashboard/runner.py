from services.water_dashboard.scraper import scrape_all
from services.water_dashboard.builder import build_snapshot


def run_water_dashboard_pipeline():
    print("STAGE: Сбор данных из 8 дашбордов DataLens")
    extractions = scrape_all()

    print("STAGE: Сборка снимка")
    snap = build_snapshot(extractions)

    print(f"Готово: ОМСУ в таблице={len(snap['table'])}, "
          f"задач={snap['kpis']['tasks_total']}, резонансных ВС={snap['kpis']['res_vs']}")
    return snap


if __name__ == "__main__":
    run_water_dashboard_pipeline()
    