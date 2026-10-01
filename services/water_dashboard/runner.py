import sys
import argparse
import traceback

from services.water_dashboard.diagnostics import ERROR_PREFIX, error_code


class SourcesUnavailableError(RuntimeError):
    code = "sources_failed"


def run_water_dashboard_pipeline(source=None):
    # Load configuration before importing the scraper, including standalone runs.
    from dotenv import load_dotenv
    from pathlib import Path
    load_dotenv(Path(__file__).resolve().parents[2] / ".env")
    from services.water_dashboard.scraper import scrape_all
    from services.water_dashboard.builder import build_snapshot

    from services.water_dashboard.config import SOURCES
    known_ids = {item['id'] for item in SOURCES}
    if source is not None and source not in known_ids:
        raise ValueError('Unknown water dashboard source')
    selected = [source] if source is not None else [item['id'] for item in SOURCES]
    print(f"STAGE: Сбор данных из {len(selected)} дашбордов DataLens", flush=True)
    extractions = scrape_all(source_ids=selected)

    print("STAGE: Сборка снимка", flush=True)
    snap = build_snapshot(extractions, source_ids=selected)

    if not snap["last_updated_sources"]:
        raise SourcesUnavailableError("Ни один источник DataLens не обновлён")
    print(f"Готово: ОМСУ в таблице={len(snap['table'])}, "
          f"задач={snap['kpis']['tasks_total']}, резонансных ВС={snap['kpis']['res_vs']}", flush=True)
    return snap


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    from services.water_dashboard.config import SOURCES
    parser = argparse.ArgumentParser(description="Обновление сводного дашборда")
    parser.add_argument('--source', choices=[item['id'] for item in SOURCES])
    args = parser.parse_args([] if argv is None else argv)
    try:
        run_water_dashboard_pipeline(source=args.source)
    except Exception as exc:
        traceback.print_exc()  # Full diagnostic is kept in the server console.
        print(ERROR_PREFIX + error_code(exc), flush=True)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
