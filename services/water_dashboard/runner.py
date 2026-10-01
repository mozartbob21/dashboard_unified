import sys
import traceback

from services.water_dashboard.diagnostics import ERROR_PREFIX, error_code


class SourcesUnavailableError(RuntimeError):
    code = "sources_failed"


def run_water_dashboard_pipeline():
    # Load configuration before importing the scraper, including standalone runs.
    from dotenv import load_dotenv
    from pathlib import Path
    load_dotenv(Path(__file__).resolve().parents[2] / ".env")
    from services.water_dashboard.scraper import scrape_all
    from services.water_dashboard.builder import build_snapshot

    print("STAGE: Сбор данных из 8 дашбордов DataLens", flush=True)
    extractions = scrape_all()

    print("STAGE: Сборка снимка", flush=True)
    snap = build_snapshot(extractions)

    if not any(snap["sources_updated"].values()):
        raise SourcesUnavailableError("Ни один источник DataLens не обновлён")
    print(f"Готово: ОМСУ в таблице={len(snap['table'])}, "
          f"задач={snap['kpis']['tasks_total']}, резонансных ВС={snap['kpis']['res_vs']}", flush=True)
    return snap


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    try:
        run_water_dashboard_pipeline()
    except Exception as exc:
        traceback.print_exc()  # Full diagnostic is kept in the server console.
        print(ERROR_PREFIX + error_code(exc), flush=True)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
