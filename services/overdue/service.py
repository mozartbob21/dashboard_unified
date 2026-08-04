import html
import json
import re
import shutil
import time
from datetime import datetime

from playwright.sync_api import sync_playwright

from services.overdue.config import (
    DASHBOARD_URL,
    DATA_FILE,
    DEBUG_DIR,
    DEBUG_ENV_DIR,
    HEADLESS,
    IDLE_AFTER_DATA_SECONDS,
    MAX_WAIT_SECONDS,
    PLAYWRIGHT_PROFILE_DIR,
    RESPONSES_DIR,
)
from services.overdue.utils import load_json, save_json, safe_int


BLOCKED_HOST_PARTS = [
    "mc.yandex.ru",
    "metrika",
    "smartcaptcha",
    "showcaptcha",
    "captcha",
]


def prepare_dirs():
    PLAYWRIGHT_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_ENV_DIR.mkdir(parents=True, exist_ok=True)

    if RESPONSES_DIR.exists():
        shutil.rmtree(RESPONSES_DIR)

    RESPONSES_DIR.mkdir(parents=True, exist_ok=True)


def is_blocked_url(url: str) -> bool:
    lower_url = url.lower()
    return any(part in lower_url for part in BLOCKED_HOST_PARTS)


def looks_like_useful_json(data) -> bool:
    if isinstance(data, list):
        return len(data) > 0
    if isinstance(data, dict):
        return len(data) > 0
    return False


def try_launch_context(playwright):
    errors = []

    launch_variants = [
        {
            "name": "channel=chrome",
            "kwargs": {
                "user_data_dir": str(PLAYWRIGHT_PROFILE_DIR),
                "channel": "chrome",
                "headless": HEADLESS,
                "viewport": {"width": 1440, "height": 1100},
                "args": [
                    "--disable-blink-features=AutomationControlled",
                    "--start-maximized",
                ],
            },
        },
        {
            "name": "playwright chromium fallback",
            "kwargs": {
                "user_data_dir": str(PLAYWRIGHT_PROFILE_DIR),
                "headless": HEADLESS,
                "viewport": {"width": 1440, "height": 1100},
                "args": [
                    "--disable-blink-features=AutomationControlled",
                    "--start-maximized",
                ],
            },
        },
    ]

    for variant in launch_variants:
        try:
            print(f"[browser] trying: {variant['name']}")
            context = playwright.chromium.launch_persistent_context(**variant["kwargs"])
            print(f"[browser] success: {variant['name']}")
            return context, variant["name"]
        except Exception as e:
            errors.append(f"{variant['name']}: {e}")

    raise RuntimeError(
        "Не удалось запустить браузер.\n\n"
        + "\n\n".join(errors)
        + "\n\nУстановите браузер командой:\npython -m playwright install\n"
    )


def fetch_dashboard_data():
    prepare_dirs()

    screenshot_file = DEBUG_ENV_DIR / "dashboard_page.png"
    html_file = DEBUG_ENV_DIR / "dashboard_page.html"

    screenshot_web_path = "data/overdue/debug/dashboard_page.png"
    html_web_path = "data/overdue/debug/dashboard_page.html"

    saved_files = []
    saved_count = 0

    state = {
        "public_entry_seen": False,
        "dash_state_seen": False,
        "chart_run_count": 0,
        "last_useful_response_ts": None,
    }

    with sync_playwright() as p:
        context, browser_name = try_launch_context(p)
        page = context.pages[0] if context.pages else context.new_page()

        def handle_response(response):
            nonlocal saved_count

            try:
                url = response.url
                lower_url = url.lower()

                if is_blocked_url(lower_url):
                    return

                content_type = response.headers.get("content-type", "").lower()
                if "json" not in content_type:
                    return

                if response.status >= 400:
                    return

                data = response.json()
                if not looks_like_useful_json(data):
                    return

                saved_count += 1
                file_path = RESPONSES_DIR / f"{saved_count:03d}.json"

                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "url": url,
                            "status": response.status,
                            "content_type": content_type,
                            "data": data,
                        },
                        f,
                        ensure_ascii=False,
                        indent=2,
                    )

                saved_files.append(str(file_path.resolve()))
                state["last_useful_response_ts"] = time.time()

                if "getpublicentry" in lower_url:
                    state["public_entry_seen"] = True

                if "getpublicdashstate" in lower_url:
                    state["dash_state_seen"] = True

                if "/charts/api/run" in lower_url:
                    state["chart_run_count"] += 1

                print(f"[saved] {saved_count:03d} {url}")

            except Exception:
                pass

        page.on("response", handle_response)

        print(f"[open] {DASHBOARD_URL}")
        print(f"[browser] using: {browser_name}")
        page.goto(DASHBOARD_URL, wait_until="domcontentloaded", timeout=120000)

        print("")
        print("Если нужно — пройдите капчу / авторизацию вручную в окне браузера.")
        print("Ожидание завершится автоматически, когда дашборд догрузится.")
        print(f"Максимальный таймаут: {MAX_WAIT_SECONDS} сек.")
        print("")

        start_ts = time.time()

        while True:
            now = time.time()
            elapsed = now - start_ts

            enough_data_loaded = (
                state["public_entry_seen"]
                and state["dash_state_seen"]
                and state["chart_run_count"] >= 1
            )

            idle_enough = (
                state["last_useful_response_ts"] is not None
                and (now - state["last_useful_response_ts"]) >= IDLE_AFTER_DATA_SECONDS
            )

            if enough_data_loaded and idle_enough:
                print("[wait] Данные загружены, завершаем ожидание.")
                break

            if elapsed >= MAX_WAIT_SECONDS:
                print("[wait] Достигнут максимальный таймаут.")
                break

            page.wait_for_timeout(1000)

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass

        final_url = page.url
        title = page.title()

        try:
            html_file.write_text(page.content(), encoding="utf-8")
        except Exception:
            pass

        try:
            page.screenshot(path=str(screenshot_file), full_page=True)
            print(f"[saved] screenshot: {screenshot_file}")
        except Exception:
            pass

        # ── Извлечение блока ОМСУ напрямую из DOM ──
        try:
            dom_items = page.evaluate(
                """
                () => {
                    const norm = (s) => (s || '').trim();

                    // ── Способ 1: ЛЮБАЯ таблица с заголовками «ОМСУ» и «Кол-во задач» ──
                    const tables = Array.from(document.querySelectorAll('table'));
                    for (const table of tables) {
                        let headerCells = Array.from(table.querySelectorAll('thead th'));
                        if (!headerCells.length) {
                            const firstRow = table.querySelector('tr');
                            if (firstRow) headerCells = Array.from(firstRow.querySelectorAll('th, td'));
                        }
                        const headers = headerCells.map((n) => norm(n.textContent).toLowerCase());
                        const idxMun = headers.findIndex((h) => h.includes('омсу') || h.includes('муниципал'));
                        const idxCnt = headers.findIndex((h) =>
                            h.includes('кол-во') || h.includes('количество') ||
                            (h.includes('задач') && !h.includes('срок')));
                        if (idxMun === -1 || idxCnt === -1) continue;

                        const items = [];
                        for (const tr of Array.from(table.querySelectorAll('tbody tr'))) {
                            const cells = Array.from(tr.querySelectorAll('td'))
                                .map((td) => norm(td.textContent));
                            if (cells.length <= Math.max(idxMun, idxCnt)) continue;

                            const name = cells[idxMun];
                            const countRaw = (cells[idxCnt] || '').replace(/[^0-9]/g, '');
                            const count = countRaw === '' ? null : parseInt(countRaw, 10);

                            if (!name || count === null) continue;
                            if (/^итого/i.test(name)) continue;

                            items.push({
                                municipality: name,
                                organization: name,
                                overdue_count: count,
                                responsible_name: '',
                                responsible_phone: '',
                            });
                        }
                        if (items.length) {
                            return {
                                source: 'playwright_dom_omsu_table',
                                items,
                                debug: { tableRows: items.length },
                            };
                        }
                    }

                    // ── Способ 2 (фолбэк): виджет с заголовком «ОМСУ» как бар-чарт ──
                    const widgets = Array.from(document.querySelectorAll('.dl-widget'));
                    const widget = widgets.find((node) => {
                        const title = node.querySelector('.dl-widget__chart-title-text');
                        return title && title.textContent && title.textContent.trim() === 'ОМСУ';
                    });

                    if (!widget) {
                        return { source: 'dom_widget_not_found', items: [], debug: { widgetCount: widgets.length } };
                    }

                    const yLabels = Array.from(widget.querySelectorAll('.gcharts-y-axis__label'))
                        .map((node) => {
                            const text = (node.textContent || '').trim();
                            const rect = node.getBoundingClientRect();
                            return { text, top: rect.top };
                        })
                        .filter((item) => item.text && !/^\\d+$/.test(item.text));

                    const valueLabels = Array.from(widget.querySelectorAll('.gcharts-bar-y__label'))
                        .map((node) => {
                            const text = (node.textContent || '').trim();
                            const value = parseInt(text.replace(/\\s+/g, ''), 10);
                            const rect = node.getBoundingClientRect();
                            return { text, value: Number.isFinite(value) ? value : null, top: rect.top };
                        })
                        .filter((item) => item.value !== null);

                    yLabels.sort((a, b) => a.top - b.top);
                    valueLabels.sort((a, b) => a.top - b.top);

                    const count = Math.min(yLabels.length, valueLabels.length);
                    const items = [];
                    for (let i = 0; i < count; i += 1) {
                        items.push({
                            municipality: yLabels[i].text,
                            organization: yLabels[i].text,
                            overdue_count: valueLabels[i].value,
                            responsible_name: '',
                            responsible_phone: '',
                        });
                    }

                    return {
                        source: 'playwright_dom_omsu',
                        items,
                        debug: { yLabelsCount: yLabels.length, valueLabelsCount: valueLabels.length },
                    };
                }
                """
            )
            dom_items_file = DEBUG_ENV_DIR / "omsu_dom_items.json"
            with open(dom_items_file, "w", encoding="utf-8") as f:
                json.dump(dom_items, f, ensure_ascii=False, indent=2)

            print(
                "[saved] DOM ОМСУ items:",
                len(dom_items.get("items", [])),
                dom_items_file,
            )
        except Exception as e:
            print(f"[warn] DOM ОМСУ extract failed: {e}")

        context.close()

    print(f"[done] saved JSON responses: {saved_count}")

    return {
        "screenshot_path": screenshot_web_path,
        "screenshot_paths": [screenshot_web_path],
        "html_path": html_web_path,
        "responses_dir": str(RESPONSES_DIR.resolve()),
        "response_files": saved_files,
        "saved_json_count": saved_count,
        "final_url": final_url,
        "title": title,
        "public_entry_seen": state["public_entry_seen"],
        "dash_state_seen": state["dash_state_seen"],
        "chart_run_count": state["chart_run_count"],
    }


def load_wrappers():
    wrappers = []
    if not RESPONSES_DIR.exists():
        return wrappers

    for file_path in sorted(RESPONSES_DIR.glob("*.json")):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                wrapper = json.load(f)

            wrappers.append(
                {
                    "file": str(file_path),
                    "url": wrapper.get("url", ""),
                    "status": wrapper.get("status", 0),
                    "data": wrapper.get("data"),
                }
            )
        except Exception:
            continue

    return wrappers


def walk_json_objects(value):
    if isinstance(value, dict):
        yield value

        for nested_value in value.values():
            yield from walk_json_objects(nested_value)

    elif isinstance(value, list):
        for item in value:
            yield from walk_json_objects(item)


def extract_table_items_from_payload(payload):
    """Извлекает таблицу ОМСУ из ответа DataLens (columns + data)."""
    best = []

    for obj in walk_json_objects(payload):
        columns = obj.get("columns")
        rows = obj.get("data")

        if not isinstance(columns, list) or not isinstance(rows, list):
            continue
        if not columns or not rows:
            continue

        names = []
        for col in columns:
            if isinstance(col, dict):
                names.append(str(col.get("name") or col.get("title") or "").lower())
            else:
                names.append(str(col).lower())

        idx_mun = None
        idx_cnt = None
        for i, n in enumerate(names):
            if idx_mun is None and ("омсу" in n or "муниципал" in n):
                idx_mun = i
            if idx_cnt is None and ("кол-во" in n or "количество" in n or "задач" in n):
                idx_cnt = i

        if idx_mun is None or idx_cnt is None:
            continue

        items = []
        for row in rows:
            if not isinstance(row, (list, tuple)) or len(row) <= max(idx_mun, idx_cnt):
                continue
            municipality = str(row[idx_mun]).strip()
            count = safe_int(row[idx_cnt], 0)
            if not municipality or municipality.lower().startswith("итого"):
                continue
            items.append({
                "municipality": municipality,
                "organization": municipality,
                "overdue_count": count,
                "responsible_name": "",
                "responsible_phone": "",
            })

        if len(items) > len(best):
            best = items

    return best


def extract_chart_items_from_payload(payload):
    """Извлекает ОМСУ из ответа /charts/api/run: сначала таблица, затем бар-чарт."""
    table_items = extract_table_items_from_payload(payload)
    if table_items:
        return table_items

    candidates = []

    for obj in walk_json_objects(payload):
        categories = obj.get("categories")
        graphs = obj.get("graphs")

        if not isinstance(categories, list) or not isinstance(graphs, list):
            continue

        if not categories or not graphs:
            continue

        first_graph = graphs[0] or {}
        points = first_graph.get("data") or []

        if not isinstance(points, list) or not points:
            continue

        max_len = min(len(categories), len(points))
        items = []

        for i in range(max_len):
            municipality = str(categories[i]).strip()
            point = points[i] or {}

            if isinstance(point, dict):
                overdue_count = safe_int(
                    point.get("y", point.get("value", point.get("count", 0))),
                    0,
                )
            else:
                overdue_count = safe_int(point, 0)

            if not municipality:
                continue

            items.append(
                {
                    "municipality": municipality,
                    "organization": municipality,
                    "overdue_count": overdue_count,
                    "responsible_name": "",
                    "responsible_phone": "",
                }
            )

        if items:
            candidates.append(items)

    if not candidates:
        return []

    candidates.sort(
        key=lambda items: (
            len(items),
            sum(safe_int(item.get("overdue_count", 0), 0) for item in items),
        ),
        reverse=True,
    )

    return candidates[0]


def parse_translate_y(transform_value):
    if not transform_value:
        return None

    match = re.search(r"translate\(\s*[-\d.]+\s*,\s*([-\d.]+)\s*\)", transform_value)

    if not match:
        return None

    try:
        return float(match.group(1))
    except Exception:
        return None


def strip_tags(value):
    value = re.sub(r"<[^>]+>", "", value or "")
    value = html.unescape(value)
    return value.strip()


def extract_dashboard_data():
    wrappers = load_wrappers()
    omsu_chart_id = find_omsu_chart_id(wrappers)

    best_items = []
    matched_sources = []
    source = "empty"

    # ── Проход 1: DOM — отрисованная таблица (самый надёжный источник) ──
    dom_items_file = DEBUG_ENV_DIR / "omsu_dom_items.json"
    if dom_items_file.exists():
        try:
            dom_payload = load_json(dom_items_file, default={}) or {}
            dom_items = dom_payload.get("items", []) or []
            if dom_items:
                best_items = dom_items
                source = dom_payload.get("source", "playwright_dom_omsu")
                matched_sources = [
                    {
                        "file": str(dom_items_file),
                        "url": "dashboard_dom",
                        "items_found": len(best_items),
                        "chart_id": "dom_omsu",
                    }
                ]
        except Exception:
            pass

    # ── Проход 2: таблица ОМСУ из сетевых ответов ──
    if not best_items:
        for wrapper in wrappers:
            items = extract_table_items_from_payload(wrapper.get("data") or {})
            if not items:
                continue
            total = sum(safe_int(i.get("overdue_count", 0), 0) for i in items)
            matched_sources.append({
                "file": wrapper.get("file", ""),
                "url": wrapper.get("url", ""),
                "items_found": len(items),
                "total_overdue": total,
                "chart_id": "table_omsu",
            })
            if len(items) > len(best_items):
                best_items = items
                source = "saved_network_table"

    # ── Проход 3: бар-чарты (старый вид) ──
    if not best_items:
        for wrapper in wrappers:
            items = extract_from_chart_run_response(wrapper, omsu_chart_id)
            if not items:
                continue
            total = sum(safe_int(i.get("overdue_count", 0), 0) for i in items)
            matched_sources.append({
                "file": wrapper.get("file", ""),
                "url": wrapper.get("url", ""),
                "items_found": len(items),
                "total_overdue": total,
                "chart_id": (wrapper.get("data") or {}).get("id")
                or (wrapper.get("data") or {}).get("_confStorageConfig", {}).get("entryId", ""),
            })
            current_total = sum(safe_int(i.get("overdue_count", 0), 0) for i in best_items)
            if len(items) > len(best_items) or total > current_total:
                best_items = items
                source = "saved_network"

    # ── Проход 4: парсинг сохранённого HTML/SVG ──
    if not best_items:
        try:
            html_extracted = extract_dashboard_data_from_html()
            best_items = html_extracted.get("items", []) or []
            matched_sources = html_extracted.get("matched_sources", []) or []
            source = html_extracted.get("source", "saved_html_svg")
        except Exception:
            pass

    best_items.sort(key=lambda x: (-safe_int(x.get("overdue_count", 0), 0), x.get("municipality", "")))

    return {
        "source": source if best_items else "empty",
        "items": best_items,
        "matched_sources": matched_sources,
        "responses_scanned": len(wrappers),
        "items_count": len(best_items),
        "summary": {
            "total_records": len(best_items),
            "critical": sum(1 for x in best_items if safe_int(x.get("overdue_count", 0), 0) >= 20),
            "risk": sum(1 for x in best_items if 0 < safe_int(x.get("overdue_count", 0), 0) < 20),
            "ok": sum(1 for x in best_items if safe_int(x.get("overdue_count", 0), 0) == 0),
        },
        "debug": {
            "omsu_chart_id": omsu_chart_id,
        },
    }
    widget_html = page_html[svg_start:svg_end + len("</svg>")]

    y_labels = []

    for match in re.finditer(
        r'<text(?=[^>]*transform="([^"]+)")[^>]*>\s*<tspan(?=[^>]*class="gcharts-y-axis__label")[^>]*>(.*?)</tspan>',
        widget_html,
        flags=re.IGNORECASE | re.DOTALL,
    ):
        y = parse_translate_y(match.group(1))
        label = strip_tags(match.group(2))

        if y is None or not label:
            continue

        if re.fullmatch(r"\d+", label):
            continue

        y_labels.append(
            {
                "y": y,
                "municipality": label,
            }
        )

    value_labels = []

    for match in re.finditer(
        r'<text(?=[^>]*class="gcharts-bar-y__label")(?=[^>]*\by="([^"]+)")[^>]*>(.*?)</text>',
        widget_html,
        flags=re.IGNORECASE | re.DOTALL,
    ):
        try:
            y = float(match.group(1))
        except Exception:
            continue

        value = strip_tags(match.group(2))
        count = safe_int(value, None)

        if count is None:
            continue

        value_labels.append(
            {
                "y": y,
                "count": count,
            }
        )

    if not y_labels or not value_labels:
        return {
            "source": "html_empty_omsu_labels",
            "items": [],
            "matched_sources": [],
            "items_count": 0,
        }

    y_labels.sort(key=lambda x: x["y"])
    value_labels.sort(key=lambda x: x["y"])

    max_len = min(len(y_labels), len(value_labels))
    items = []

    for i in range(max_len):
        municipality = y_labels[i]["municipality"]
        overdue_count = value_labels[i]["count"]

        items.append(
            {
                "municipality": municipality,
                "organization": municipality,
                "overdue_count": overdue_count,
                "responsible_name": "",
                "responsible_phone": "",
            }
        )

    items.sort(key=lambda x: (-safe_int(x["overdue_count"], 0), x["municipality"]))

    return {
        "source": "saved_html_svg",
        "items": items,
        "matched_sources": [
            {
                "file": str(html_file),
                "url": "dashboard_page.html",
                "items_found": len(items),
                "chart_id": "html_svg_omsu",
            }
        ],
        "items_count": len(items),
    }


def find_omsu_chart_id(wrappers):
    for wrapper in wrappers:
        url = (wrapper.get("url") or "").lower()
        if "getpublicentry" not in url:
            continue

        payload = wrapper.get("data") or {}
        dash_data = payload.get("data") or {}
        tabs = dash_data.get("tabs") or []

        for tab in tabs:
            for item in tab.get("items", []):
                if item.get("type") != "widget":
                    continue

                item_data = item.get("data") or {}
                widget_tabs = item_data.get("tabs") or []

                for widget_tab in widget_tabs:
                    title = (widget_tab.get("title") or "").strip().lower()
                    chart_id = (widget_tab.get("chartId") or "").strip()

                    if title == "омсу" and chart_id:
                        return chart_id

    return None


def extract_from_chart_run_response(wrapper, omsu_chart_id):
    url = wrapper.get("url") or ""
    payload = wrapper.get("data") or {}

    if "/charts/api/run" not in url:
        return []

    wrapper_chart_id = (
        payload.get("id")
        or payload.get("_confStorageConfig", {}).get("entryId")
        or payload.get("entryId")
        or payload.get("chartId")
    )

    if omsu_chart_id and wrapper_chart_id and wrapper_chart_id != omsu_chart_id:
        return []

    items = extract_chart_items_from_payload(payload)

    if not items:
        return []

    return items


def extract_dashboard_data():
    wrappers = load_wrappers()
    omsu_chart_id = find_omsu_chart_id(wrappers)

    best_items = []
    matched_sources = []
    source = "empty"

    # Проход 1: таблица ОМСУ из сетевых ответов (новый вид дашборда)
    for wrapper in wrappers:
        items = extract_table_items_from_payload(wrapper.get("data") or {})
        if not items:
            continue

        total = sum(safe_int(i.get("overdue_count", 0), 0) for i in items)
        matched_sources.append({
            "file": wrapper.get("file", ""),
            "url": wrapper.get("url", ""),
            "items_found": len(items),
            "total_overdue": total,
            "chart_id": "table_omsu",
        })
        if len(items) > len(best_items):
            best_items = items
            source = "saved_network_table"

    # Проход 2: бар-чарты (старый вид), только если таблица не найдена
    if not best_items:
        for wrapper in wrappers:
            items = extract_from_chart_run_response(wrapper, omsu_chart_id)
            if not items:
                continue

            total = sum(safe_int(i.get("overdue_count", 0), 0) for i in items)
            matched_sources.append({
                "file": wrapper.get("file", ""),
                "url": wrapper.get("url", ""),
                "items_found": len(items),
                "total_overdue": total,
                "chart_id": (wrapper.get("data") or {}).get("id")
                or (wrapper.get("data") or {}).get("_confStorageConfig", {}).get("entryId", ""),
            })

            current_total = sum(safe_int(i.get("overdue_count", 0), 0) for i in best_items)
            if len(items) > len(best_items) or total > current_total:
                best_items = items
                source = "saved_network"

    # Проход 3: данные, снятые напрямую из отрисованного DOM
    if not best_items:
        dom_items_file = DEBUG_ENV_DIR / "omsu_dom_items.json"

        if dom_items_file.exists():
            try:
                dom_payload = load_json(dom_items_file, default={}) or {}
                dom_items = dom_payload.get("items", []) or []

                if dom_items:
                    best_items = dom_items
                    source = dom_payload.get("source", "playwright_dom_omsu")
                    matched_sources = [
                        {
                            "file": str(dom_items_file),
                            "url": "dashboard_dom",
                            "items_found": len(best_items),
                            "chart_id": "dom_omsu",
                        }
                    ]
            except Exception:
                pass

    # Проход 4: парсинг сохранённого HTML/SVG
    if not best_items:
        try:
            html_extracted = extract_dashboard_data_from_html()
            best_items = html_extracted.get("items", []) or []
            matched_sources = html_extracted.get("matched_sources", []) or []
            source = html_extracted.get("source", "saved_html_svg")
        except Exception:
            pass

    best_items.sort(key=lambda x: (-safe_int(x.get("overdue_count", 0), 0), x.get("municipality", "")))

    return {
        "source": source if best_items else "empty",
        "items": best_items,
        "matched_sources": matched_sources,
        "responses_scanned": len(wrappers),
        "items_count": len(best_items),
        "summary": {
            "total_records": len(best_items),
            "critical": sum(1 for x in best_items if safe_int(x.get("overdue_count", 0), 0) >= 20),
            "risk": sum(1 for x in best_items if 0 < safe_int(x.get("overdue_count", 0), 0) < 20),
            "ok": sum(1 for x in best_items if safe_int(x.get("overdue_count", 0), 0) == 0),
        },
        "debug": {
            "omsu_chart_id": omsu_chart_id,
        },
    }


def format_municipality(name):
    """ОДИНЦОВСКИЙ → Одинцовский, СЕРГИЕВО-ПОСАДСКИЙ → Сергиево-Посадский."""
    name = (name or "").strip()
    if not name or not name.isupper():
        return name

    words = name.split()
    formatted = ["-".join(p.capitalize() for p in w.split("-")) for w in words]

    if len(formatted) > 1:
        return formatted[0] + " " + " ".join(w.lower() for w in formatted[1:])
    return formatted[0]


def normalize_items(raw_items):
    normalized = []

    for item in raw_items or []:
        raw_mun = (item.get("municipality") or item.get("name") or "Не указано").strip()
        raw_org = (item.get("organization") or raw_mun).strip()

        municipality = format_municipality(raw_mun)
        organization = format_municipality(raw_org) or municipality
        overdue_count = safe_int(item.get("overdue_count", item.get("count", 0)), 0)

        normalized.append(
            {
                "municipality": municipality,
                "organization": organization,
                "overdue_count": overdue_count,
                "category": "Просроченные задачи",
                "responsible_name": item.get("responsible_name", ""),
                "responsible_phone": item.get("responsible_phone", ""),
            }
        )

    normalized.sort(key=lambda x: (-x["overdue_count"], x["municipality"]))
    return normalized


def build_summary(items):
    total_overdue = sum(safe_int(item.get("overdue_count", 0), 0) for item in items)

    critical_count = 0
    risk_count = 0
    ok_count = 0
    by_municipality = []

    for item in items:
        municipality = item.get("municipality", "Не указано")
        overdue_count = safe_int(item.get("overdue_count", 0), 0)

        if overdue_count >= 20:
            critical_count += 1
        elif overdue_count > 0:
            risk_count += 1
        else:
            ok_count += 1

        by_municipality.append(
            {
                "municipality": municipality,
                "count": overdue_count,
            }
        )

    by_municipality.sort(key=lambda x: (-x["count"], x["municipality"]))

    return {
        "total_overdue": total_overdue,
        "by_status": [
            {"status": "Критично", "count": critical_count},
            {"status": "Риск", "count": risk_count},
            {"status": "Норма", "count": ok_count},
        ],
        "by_municipality": by_municipality,
        "by_category": [],
    }


def build_public_message(summary, items):
    bad_items = [x for x in items if safe_int(x.get("overdue_count", 0), 0) > 0]
    bad_items.sort(key=lambda x: (-safe_int(x.get("overdue_count", 0), 0), x.get("municipality", "")))

    lines = [
        "Добрый день.",
        "По итогам проверки информационной панели выполнения поручений, зафиксированных в системе управления Министерства ЖКХ МО, выявлены невыполненные задачи.",

        f"Общее количество просроченных задач: {summary.get('total_overdue', 0)}.",
    ]

    if bad_items:
        lines.append("")
        lines.append("ОМСУ с наибольшим количеством просроченных задач:")
        for item in bad_items[:15]:
            lines.append(f"- {item.get('municipality', 'Не указано')}: {item.get('overdue_count', 0)}")

        lines.append("")
        lines.append("Просьба оперативно отработать просроченные позиции и актуализировать сведения.")
    else:
        lines.append("")
        lines.append("Просроченные задачи не выявлены. Спасибо за своевременное обновление данных.")

    return "\n".join(lines)


def build_missing_data_issues(items):
    issues = []

    for item in items:
        municipality = item.get("municipality", "Не указано")
        organization = item.get("organization", municipality)
        responsible_name = item.get("responsible_name", "")
        responsible_phone = item.get("responsible_phone", "")
        overdue_count = safe_int(item.get("overdue_count", 0), 0)

        missing_fields = []
        if not responsible_name:
            missing_fields.append("не указан ответственный")
        if not responsible_phone:
            missing_fields.append("не указан телефон")

        if missing_fields:
            issues.append(
                {
                    "municipality": municipality,
                    "organization": organization,
                    "responsible_name": responsible_name or "Не указан",
                    "responsible_phone": responsible_phone or "",
                    "message": (
                        f"По записи '{municipality} / {organization}' обнаружены незаполненные данные: "
                        f"{', '.join(missing_fields)}."
                        f"{' Дополнительно зафиксировано просроченных задач: ' + str(overdue_count) + '.' if overdue_count > 0 else ''}"
                    ),
                }
            )

    return issues


def build_personal_messages(items):
    messages = []

    for item in items:
        overdue_count = safe_int(item.get("overdue_count", 0), 0)
        if overdue_count <= 0:
            continue

        status = "critical" if overdue_count >= 20 else "risk"
        municipality = item.get("municipality", "Не указано")
        organization = item.get("organization", municipality)
        responsible_name = item.get("responsible_name", "Коллега")
        responsible_phone = item.get("responsible_phone", "")

        message = (
            f"Добрый день!\n\n"
            f"По итогам проверки информационной панели выполнения поручений, зафиксированных в системе управления Министерства ЖКХ МО по ОМСУ '{municipality}' "
            f"выявлены невыполненные задачи: {overdue_count}.\n"
            f"Просьба проверить блок '{organization}', отработать просроченные позиции и актуализировать отчёт.\n\n"
            f"Необходимо до конца следующего рабочего дня внести комментарии о текущем статусе исполнения поручения и перевести задачу на контролёра."
        )

        messages.append(
            {
                "municipality": municipality,
                "organization": organization,
                "responsible_name": responsible_name,
                "responsible_phone": responsible_phone,
                "status": status,
                "message": message,
                "is_edited": False,
            }
        )

    messages.sort(key=lambda x: (0 if x["status"] == "critical" else 1, -len(x["message"])))
    return messages


def build_report_text(summary, items):
    lines = [
        "Текстовый отчёт",
        "",
        f"Дата формирования: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Всего просроченных задач: {summary.get('total_overdue', 0)}",
        "",
        "Распределение по статусам:",
    ]

    for item in summary.get("by_status", []):
        lines.append(f"- {item['status']}: {item['count']}")

    lines.append("")
    lines.append("Топ ОМСУ:")

    for item in summary.get("by_municipality", [])[:15]:
        lines.append(f"- {item['municipality']}: {item['count']}")

    if items:
        lines.append("")
        lines.append("Детализация:")
        for item in items[:30]:
            lines.append(f"- {item.get('municipality', 'Не указано')}: {item.get('overdue_count', 0)}")

    return "\n".join(lines)


def run_overdue_pipeline():
    screenshot_paths = []
    extraction_note = ""

    try:
        fetch_result = fetch_dashboard_data()
        if isinstance(fetch_result, dict):
            screenshot_paths = fetch_result.get("screenshot_paths", []) or []
            if fetch_result.get("screenshot_path"):
                screenshot_paths.append(fetch_result["screenshot_path"])
    except Exception as e:
        extraction_note = f"Не удалось полностью выполнить Playwright-сценарий: {e}"

    try:
        extracted = extract_dashboard_data()
    except Exception as e:
        extracted = {"items": [], "source": "fallback"}
        if extraction_note:
            extraction_note += f"\nТакже не удалось извлечь данные: {e}"
        else:
            extraction_note = f"Не удалось извлечь данные: {e}"

    items = normalize_items(extracted.get("items", []))
    summary = build_summary(items)
    public_message = build_public_message(summary, items)
    personal_messages = build_personal_messages(items)
    missing_data_issues = build_missing_data_issues(items)
    report_text = build_report_text(summary, items)

    screenshot_paths = list(dict.fromkeys([x for x in screenshot_paths if x]))
    screenshot_path = screenshot_paths[0] if screenshot_paths else ""

    result = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "public_message": public_message,
        "report_text": report_text,
        "summary": summary,
        "items": items,
        "screenshot_path": screenshot_path,
        "screenshot_paths": screenshot_paths,
        "missing_data_issues": missing_data_issues,
        "personal_messages": personal_messages,
        "extraction_note": extraction_note,
        "redmine_url": "",
    }

    save_json(DATA_FILE, result)
    return result


def load_overdue_result():
    return load_json(DATA_FILE, default=None)