import os
import re
import csv
import json
import math
import shutil
import hashlib
import asyncio
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv
from playwright.async_api import TimeoutError as PlaywrightTimeoutError
from playwright.async_api import async_playwright

try:
    from PIL import Image, ImageOps, ImageStat, ImageChops

    PIL_OK = True
except Exception:
    Image = None
    ImageOps = None
    ImageStat = None
    ImageChops = None
    PIL_OK = False


load_dotenv()

TARGET_LOGIN_URL = os.getenv("TARGET_LOGIN_URL", "").strip()
TARGET_CAMERAS_URL = os.getenv("TARGET_CAMERAS_URL", "").strip()
TARGET_USERNAME = os.getenv("TARGET_USERNAME", "").strip()
TARGET_PASSWORD = os.getenv("TARGET_PASSWORD", "").strip()

PLAYWRIGHT_HEADLESS = os.getenv("PLAYWRIGHT_HEADLESS", "true").lower() == "true"
PLAYWRIGHT_CHANNEL = os.getenv("PLAYWRIGHT_CHANNEL", "").strip()

PROJECT_ROOT = Path(__file__).resolve().parents[2]

BASE_DIR = PROJECT_ROOT
CAMERAS_DATA_DIR = PROJECT_ROOT / "data" / "cameras"

DEBUG_DIR = CAMERAS_DATA_DIR / "debug"
STREAM_DEBUG_DIR = DEBUG_DIR / "streams"
FFMPEG_DEBUG_DIR = DEBUG_DIR / "ffmpeg"
ADDRESS_TABLE_PATH = CAMERAS_DATA_DIR / "addresses.tsv"

DEBUG_DIR.mkdir(parents=True, exist_ok=True)
STREAM_DEBUG_DIR.mkdir(parents=True, exist_ok=True)
FFMPEG_DEBUG_DIR.mkdir(parents=True, exist_ok=True)

STREAM_CHECK_CONCURRENCY = int(os.getenv("STREAM_CHECK_CONCURRENCY", "3"))
STREAM_RETRY_ATTEMPTS = int(os.getenv("STREAM_RETRY_ATTEMPTS", "2"))

STREAM_GOTO_TIMEOUT_MS = int(os.getenv("STREAM_GOTO_TIMEOUT_MS", "12000"))
STREAM_POST_GOTO_WAIT_MS = int(os.getenv("STREAM_POST_GOTO_WAIT_MS", "2000"))
STREAM_NETWORKIDLE_TIMEOUT_MS = int(os.getenv("STREAM_NETWORKIDLE_TIMEOUT_MS", "2500"))

PLAY_ENABLED = os.getenv("PLAY_ENABLED", "true").lower() == "true"
PLAY_JS_TIMEOUT_MS = int(os.getenv("PLAY_JS_TIMEOUT_MS", "4000"))
PLAY_TOTAL_TIMEOUT_MS = int(os.getenv("PLAY_TOTAL_TIMEOUT_MS", "12000"))
PLAY_SETTLE_WAIT_MS = int(os.getenv("PLAY_SETTLE_WAIT_MS", "2500"))

MEDIA_COLLECT_WAIT_MS = int(os.getenv("MEDIA_COLLECT_WAIT_MS", "6000"))
MEDIA_COLLECT_WAIT_DEEP_MS = int(os.getenv("MEDIA_COLLECT_WAIT_DEEP_MS", "10000"))
MAX_MEDIA_CANDIDATES = int(os.getenv("MAX_MEDIA_CANDIDATES", "10"))

FFMPEG_ENABLED = os.getenv("FFMPEG_ENABLED", "true").lower() == "true"
FFMPEG_BIN = os.getenv("FFMPEG_BIN", "ffmpeg").strip()
FFPROBE_BIN = os.getenv("FFPROBE_BIN", "ffprobe").strip()
FFMPEG_CONCURRENCY = int(os.getenv("FFMPEG_CONCURRENCY", "2"))
FFMPEG_TIMEOUT_SEC = float(os.getenv("FFMPEG_TIMEOUT_SEC", "25"))
FFPROBE_TIMEOUT_SEC = float(os.getenv("FFPROBE_TIMEOUT_SEC", "12"))
FFMPEG_FRAME_COUNT = int(os.getenv("FFMPEG_FRAME_COUNT", "8"))
FFMPEG_RW_TIMEOUT_US = int(os.getenv("FFMPEG_RW_TIMEOUT_US", "12000000"))
FFMPEG_RTSP_TRANSPORT = os.getenv("FFMPEG_RTSP_TRANSPORT", "tcp").strip().lower()

BROWSER_SCREENSHOT_FALLBACK = os.getenv("BROWSER_SCREENSHOT_FALLBACK", "true").lower() == "true"

BLACK_SCREEN_DETECT_ENABLED = os.getenv("BLACK_SCREEN_DETECT_ENABLED", "true").lower() == "true"
FRAME_BLACK_RATIO = float(os.getenv("FRAME_BLACK_RATIO", "0.86"))
FRAME_CENTER_BLACK_RATIO = float(os.getenv("FRAME_CENTER_BLACK_RATIO", "0.90"))
FRAME_MAX_MEAN_LUMA = float(os.getenv("FRAME_MAX_MEAN_LUMA", "26"))
FRAME_MAX_BRIGHT_RATIO = float(os.getenv("FRAME_MAX_BRIGHT_RATIO", "0.045"))
FRAME_MAX_ENTROPY = float(os.getenv("FRAME_MAX_ENTROPY", "2.05"))
FRAME_STATIC_MEAN_DIFF_MAX = float(os.getenv("FRAME_STATIC_MEAN_DIFF_MAX", "1.6"))
FRAME_STATIC_CHANGED_RATIO_MAX = float(os.getenv("FRAME_STATIC_CHANGED_RATIO_MAX", "0.012"))

FFMPEG_SEMAPHORE = asyncio.Semaphore(FFMPEG_CONCURRENCY)


def log(message: str):
    print(f"[camera_checker_ffmpeg] {message}", flush=True)


async def run_camera_check(progress_callback):
    if not TARGET_LOGIN_URL:
        raise ValueError("Не указан TARGET_LOGIN_URL в .env")
    if not TARGET_CAMERAS_URL:
        raise ValueError("Не указан TARGET_CAMERAS_URL в .env")
    if not TARGET_USERNAME:
        raise ValueError("Не указан TARGET_USERNAME в .env")
    if not TARGET_PASSWORD:
        raise ValueError("Не указан TARGET_PASSWORD в .env")

    if FFMPEG_ENABLED:
        if not shutil.which(FFMPEG_BIN):
            log(f"ВНИМАНИЕ: ffmpeg не найден: {FFMPEG_BIN}")
        if not shutil.which(FFPROBE_BIN):
            log(f"ВНИМАНИЕ: ffprobe не найден: {FFPROBE_BIN}")

    async with async_playwright() as p:
        await progress_callback("Запуск браузера", "Открывается Chromium")

        launch_kwargs = {
            "headless": PLAYWRIGHT_HEADLESS,
            "args": [
                "--autoplay-policy=no-user-gesture-required",
                "--ignore-certificate-errors",
                "--allow-running-insecure-content",
                "--mute-audio",
            ],
        }

        if PLAYWRIGHT_CHANNEL:
            launch_kwargs["channel"] = PLAYWRIGHT_CHANNEL

        browser = await p.chromium.launch(**launch_kwargs)

        context = await browser.new_context(
            viewport={"width": 1600, "height": 1200},
            ignore_https_errors=True,
            user_agent=default_user_agent(),
        )

        page = await context.new_page()

        try:
            await progress_callback("Открытие сайта", TARGET_LOGIN_URL)
            await page.goto(TARGET_LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_timeout(1500)
            await save_debug(page, "01_login_page")

            await progress_callback("Авторизация", "Выполняется вход")
            await login_to_site(page)
            await save_debug(page, "02_after_login")

            await progress_callback("Переход к камерам", TARGET_CAMERAS_URL)
            await open_cameras_page(page)
            await save_debug(page, "03_cameras_opened")

            await progress_callback("Подготовка списка", "Открывается вкладка Список")
            await switch_to_list_tab(page)
            await ensure_list_grid_loaded(page)
            await save_debug(page, "04_list_grid_loaded")

            await progress_callback("Сканирование таблицы", "Сбор строк")
            site_rows = await collect_all_rows_from_grid(page)

            await progress_callback("Проверка камер", "FFmpeg-проверка потоков")
            site_rows = await verify_streams(browser, site_rows, progress_callback)

            table_addresses = load_addresses_from_tsv(ADDRESS_TABLE_PATH)
            merged_rows = merge_site_rows_with_table(site_rows, table_addresses)

            checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            results = []
            counts = {"working": 0, "not_working": 0, "not_connected": 0, "unknown": 0}

            for index, row in enumerate(merged_rows, start=1):
                status = row.get("camera_status", "unknown")
                counts[status] = counts.get(status, 0) + 1

                results.append(
                    {
                        "id": str(index),
                        "address": row.get("address", ""),
                        "owner": row.get("owner", ""),
                        "municipality": row.get("municipality", ""),
                        "city": row.get("municipality", "").strip(),
                        "camera_status": status,
                        "checked_at": checked_at,
                        "document_type": "prescription" if status == "not_working" else None,
                        "document_path": None,
                        "contract": row.get("contract", ""),
                        "work_type": row.get("work_type", ""),
                        "stream_url": mask_sensitive_url(row.get("stream_url", "")),
                        "last_check": row.get("last_check", ""),
                        "responsible": row.get("responsible", ""),
                        "stream_check_details": row.get("stream_check_details", ""),
                    }
                )

            await progress_callback(
                "Завершено",
                (
                    f"Всего: {len(results)}, "
                    f"работает: {counts.get('working', 0)}, "
                    f"не работает: {counts.get('not_working', 0)}, "
                    f"не подключены: {counts.get('not_connected', 0)}, "
                    f"unknown: {counts.get('unknown', 0)}"
                ),
                results,
            )

            return results

        except Exception as e:
            log(f"ОШИБКА: {type(e).__name__}: {e}")
            try:
                await save_debug(page, "99_error_state")
            except Exception:
                pass
            raise

        finally:
            await context.close()
            await browser.close()
            log("Браузер закрыт")


async def login_to_site(page):
    login_input = await find_first_visible(
        page,
        [
            'input[name="username"]',
            'input[name="login"]',
            'input[type="email"]',
            'input[type="text"]',
        ],
        timeout=20000,
    )

    password_input = await find_first_visible(
        page,
        [
            'input[name="password"]',
            'input[type="password"]',
        ],
        timeout=20000,
    )

    submit_button = await find_first_visible(
        page,
        [
            'button[type="submit"]',
            'input[type="submit"]',
            'button:has-text("Войти")',
            'button:has-text("Вход")',
            'button:has-text("Авторизоваться")',
        ],
        timeout=20000,
    )

    if login_input is None:
        raise RuntimeError("Не найдено поле логина")
    if password_input is None:
        raise RuntimeError("Не найдено поле пароля")
    if submit_button is None:
        raise RuntimeError("Не найдена кнопка входа")

    await login_input.fill(TARGET_USERNAME)
    await password_input.fill(TARGET_PASSWORD)
    await submit_button.click()

    try:
        await page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass

    await page.wait_for_timeout(2500)


async def open_cameras_page(page):
    await page.goto(TARGET_CAMERAS_URL, wait_until="domcontentloaded", timeout=60000)
    await page.wait_for_timeout(2500)

    selectors = [
        "#cctv",
        "#cams-tabs",
        "#tabs_tabs_tab_list",
        "#tabs_tabs",
        ".w2ui-tabs",
        ".w2ui-grid",
        'text="Видеонаблюдение"',
    ]

    for selector in selectors:
        try:
            locator = page.locator(selector).first
            await locator.wait_for(state="visible", timeout=8000)
            return
        except Exception:
            continue

    raise RuntimeError("Не удалось открыть раздел камер")


async def switch_to_list_tab(page):
    if await is_list_tab_active(page):
        return

    for _ in range(6):
        try:
            await page.evaluate(
                """
                () => {
                    try {
                        if (window.w2ui && w2ui['tabs']) {
                            w2ui['tabs'].click('list');
                        }
                    } catch (e) {}
                }
                """
            )
        except Exception:
            pass

        await page.wait_for_timeout(1000)

        if await is_list_tab_active(page):
            return

    try:
        tab = page.locator('text="Список"').first
        if await tab.count() > 0:
            await tab.click(force=True)
            await page.wait_for_timeout(1500)
    except Exception:
        pass

    if not await is_list_tab_active(page):
        raise RuntimeError('Не удалось переключиться на вкладку "Список"')


async def is_list_tab_active(page):
    try:
        return await page.evaluate(
            """
            () => {
                const listTab = document.querySelector('#tabs_tabs_tab_list .w2ui-tab');
                if (listTab && (listTab.className || '').includes('active')) return true;

                const active = document.querySelector('#tabs_tabs .w2ui-tab.active');
                const text = active ? ((active.innerText || active.textContent || '').toLowerCase()) : '';
                return text.includes('список');
            }
            """
        )
    except Exception:
        return False


async def ensure_list_grid_loaded(page):
    selectors = [
        ".w2ui-grid",
        ".w2ui-grid-body",
        ".w2ui-grid-records",
        "#grid_cctv_grid_records",
        "#grid_camPlayerGrid_records",
        "table",
    ]

    for _ in range(20):
        for selector in selectors:
            try:
                locator = page.locator(selector).first
                if await locator.count() > 0 and await locator.is_visible():
                    return
            except Exception:
                pass

        await page.wait_for_timeout(700)

    raise RuntimeError("Таблица камер не загрузилась")


async def collect_all_rows_from_grid(page):
    grid_body = await find_grid_scroll_container(page)

    if grid_body is None:
        raise RuntimeError("Не найден scroll-контейнер таблицы")

    all_rows = {}
    previous_count = 0
    stable_rounds = 0

    for iteration in range(1, 100):
        visible_rows = await extract_visible_rows(page)
        added = 0

        for row in visible_rows:
            key = make_row_key(row)

            if key and key not in all_rows:
                all_rows[key] = row
                added += 1

        current_count = len(all_rows)

        log(
            f"Сканирование таблицы: iteration={iteration}, "
            f"visible={len(visible_rows)}, total={current_count}, added={added}"
        )

        stable_rounds = stable_rounds + 1 if current_count == previous_count else 0
        previous_count = current_count

        scroll_result = await scroll_grid_down(grid_body)
        await page.wait_for_timeout(450)

        if scroll_result.get("reachedBottom") and stable_rounds >= 2:
            break

        if stable_rounds >= 8:
            break

    await save_debug(page, "06_table_scanned")
    return list(all_rows.values())


async def find_grid_scroll_container(page):
    selectors = [
        "#grid_camPlayerGrid_records",
        "#grid_camPlayerGrid_body",
        "#grid_cctv_grid_records",
        "#grid_cctv_grid_body",
        ".w2ui-grid-records",
        ".w2ui-grid-body",
        "div[id$='_records']",
        "div[id$='_body']",
    ]

    best_locator = None
    best_scroll_height = 0

    for selector in selectors:
        try:
            locator = page.locator(selector).first

            if await locator.count() == 0:
                continue

            if not await locator.is_visible():
                continue

            info = await locator.evaluate(
                """
                el => ({
                    clientHeight: el.clientHeight || 0,
                    scrollHeight: el.scrollHeight || 0
                })
                """
            )

            if info["scrollHeight"] > info["clientHeight"] and info["scrollHeight"] > best_scroll_height:
                best_locator = locator
                best_scroll_height = info["scrollHeight"]

        except Exception:
            continue

    return best_locator


async def scroll_grid_down(grid_body):
    try:
        return await grid_body.evaluate(
            """
            el => {
                const before = el.scrollTop || 0;
                const clientHeight = el.clientHeight || 0;
                const scrollHeight = el.scrollHeight || 0;
                const maxScroll = Math.max(0, scrollHeight - clientHeight);
                const step = Math.max(240, Math.floor(clientHeight * 0.8));
                const target = Math.min(before + step, maxScroll);
                el.scrollTop = target;
                const after = el.scrollTop || 0;

                return {
                    before,
                    after,
                    changed: after !== before,
                    maxScroll,
                    reachedBottom: after >= maxScroll || maxScroll <= 0
                };
            }
            """
        )
    except Exception:
        return {"changed": False, "reachedBottom": True}


async def extract_visible_rows(page):
    rows = page.locator(".w2ui-grid-records tr")
    count = await rows.count()
    result = []

    for i in range(count):
        row_info = await parse_row(rows.nth(i))

        if row_info:
            result.append(row_info)

    return result


async def parse_row(row):
    try:
        row_text = normalize_text(await row.inner_text())
    except Exception:
        return None

    if not row_text:
        return None

    cells = row.locator("td")
    td_count = await cells.count()

    if td_count < 5:
        return None

    values = {}

    for i in range(td_count):
        cell = cells.nth(i)

        try:
            col_attr = await cell.get_attribute("col")
        except Exception:
            col_attr = None

        try:
            text = (await cell.inner_text()).strip()
        except Exception:
            text = ""

        key = col_attr if col_attr is not None else str(i)
        values[key] = text

    address = values.get("2", "").strip()
    municipality = values.get("3", "").strip()
    contract = values.get("4", "").strip()
    owner = values.get("6", "").strip() or values.get("5", "").strip()
    work_type = values.get("8", "").strip()
    stream_url = values.get("15", "").strip()
    last_check = values.get("16", "").strip()
    responsible = values.get("17", "").strip()

    if not address and not stream_url:
        return None

    return {
        "address": address,
        "municipality": municipality,
        "contract": contract,
        "owner": owner,
        "work_type": work_type,
        "stream_url": stream_url,
        "last_check": last_check,
        "responsible": responsible,
        "camera_status": "not_connected" if not stream_url else "unknown",
        "stream_check_details": "",
    }


def make_row_key(row):
    parts = [
        row.get("address", ""),
        row.get("contract", ""),
        row.get("stream_url", ""),
    ]

    return " | ".join(part.strip().lower() for part in parts if part and part.strip())


async def verify_streams(browser, rows, progress_callback):
    verified_rows = [None] * len(rows)
    queue = asyncio.Queue()

    for index, row in enumerate(rows):
        stream_url = (row.get("stream_url") or "").strip()

        if not stream_url:
            row["camera_status"] = "not_connected"
            row["stream_check_details"] = "Пустая ссылка на поток"
            verified_rows[index] = row
        else:
            await queue.put((index, row))

    total = queue.qsize()
    log(f"FFmpeg-проверка потоков: total={total}, concurrency={STREAM_CHECK_CONCURRENCY}")

    if total == 0:
        return [row for row in verified_rows if row is not None]

    progress_state = {"started": 0}
    progress_lock = asyncio.Lock()

    async def worker(worker_id):
        while True:
            try:
                index, row = await queue.get()
            except Exception:
                return

            try:
                async with progress_lock:
                    progress_state["started"] += 1
                    number = progress_state["started"]

                await progress_callback(
                    "Проверка потоков",
                    f"Проверка камеры {number}/{total}: {row.get('address', '')}",
                )

                status, details = await check_single_stream_with_retry(
                    browser,
                    row.get("stream_url", ""),
                    row.get("address", ""),
                )

                row["camera_status"] = status
                row["stream_check_details"] = details

                log(
                    f"[worker={worker_id}] "
                    f"address={row.get('address', '')}, "
                    f"status={status}, details={details[:1800]}"
                )

            except Exception as e:
                row["camera_status"] = "unknown"
                row["stream_check_details"] = f"exception:{type(e).__name__}:{e}"

            finally:
                verified_rows[index] = row
                queue.task_done()

    workers = [
        asyncio.create_task(worker(i))
        for i in range(1, STREAM_CHECK_CONCURRENCY + 1)
    ]

    await queue.join()

    for task in workers:
        task.cancel()

    await asyncio.gather(*workers, return_exceptions=True)

    return [row for row in verified_rows if row is not None]


async def check_single_stream_with_retry(browser, stream_url, address, attempts=STREAM_RETRY_ATTEMPTS):
    first = await check_single_stream(
        browser,
        stream_url,
        address,
        attempt=1,
        deep_mode=False,
    )

    if first["confident"] or attempts <= 1:
        return first["status"], f"attempt_1={first['status']}[{first['details']}]"

    second = await check_single_stream(
        browser,
        stream_url,
        address,
        attempt=2,
        deep_mode=True,
    )

    if first["status"] == "working" or second["status"] == "working":
        final_status = "working"
    elif first["status"] == "not_working" and second["status"] == "not_working":
        final_status = "not_working"
    else:
        final_status = "unknown"

    return final_status, (
        f"attempt_1={first['status']}[{first['details']}] | "
        f"attempt_2={second['status']}[{second['details']}]"
    )


async def check_single_stream(browser, stream_url, address, attempt=1, deep_mode=False):
    debug_name = build_stream_debug_name(address, f"{stream_url}_attempt_{attempt}")

    if is_direct_ffmpeg_url(stream_url):
        headers_info = build_direct_ffmpeg_headers(stream_url)

        direct_candidate = {
            "url": stream_url,
            "source": "direct_ffmpeg_url_before_browser",
            "content_type": "",
            "score": 999,
        }

        ffmpeg_result = await ffmpeg_check_candidates(
            media_candidates=[direct_candidate],
            original_url=stream_url,
            headers_info=headers_info,
            debug_name=debug_name,
            deep_mode=deep_mode,
        )

        if ffmpeg_result["status"] == "working":
            return result_working(
                f"direct_ffmpeg_working;ffmpeg={compact_json(ffmpeg_result)}",
                confident=True,
            )

        if ffmpeg_result["status"] == "not_working":
            return result_not_working(
                f"direct_ffmpeg_not_working;ffmpeg={compact_json(ffmpeg_result)}",
                confident=True,
            )

        return result_unknown(
            f"direct_ffmpeg_unknown;ffmpeg={compact_json(ffmpeg_result)}",
            confident=deep_mode,
        )

    context = await browser.new_context(
        viewport={"width": 1280, "height": 900},
        ignore_https_errors=True,
        user_agent=default_user_agent(),
    )

    page = await context.new_page()

    collector = MediaCandidateCollector()
    attach_media_collectors(page, collector)

    main_response_info = {
        "status": None,
        "content_type": "",
        "final_url": "",
    }

    try:
        goto_error = None
        response = None

        try:
            response = await page.goto(
                stream_url,
                wait_until="domcontentloaded",
                timeout=STREAM_GOTO_TIMEOUT_MS + (7000 if deep_mode else 0),
            )

            if response:
                headers = response.headers or {}
                main_response_info["status"] = response.status
                main_response_info["content_type"] = headers.get("content-type", "")
                main_response_info["final_url"] = response.url

        except Exception as e:
            goto_error = f"{type(e).__name__}:{e}"
            log(
                "goto не завершился, продолжаем проверку через network/ffmpeg: "
                f"{mask_sensitive_url(stream_url)} -> {goto_error[:400]}"
            )

            try:
                await page.evaluate("() => window.stop && window.stop()")
            except Exception:
                pass

            try:
                await save_stream_debug(page, f"{debug_name}_goto_nonfatal")
            except Exception:
                pass

        await page.wait_for_timeout(STREAM_POST_GOTO_WAIT_MS + (2500 if deep_mode else 0))

        try:
            await page.wait_for_load_state(
                "networkidle",
                timeout=STREAM_NETWORKIDLE_TIMEOUT_MS + (3000 if deep_mode else 0),
            )
        except Exception:
            pass

        current_url = page.url or stream_url
        content_type = (main_response_info.get("content_type") or "").lower()
        http_status = main_response_info.get("status")

        if http_status is not None and http_status >= 400:
            await save_stream_debug(page, f"{debug_name}_http_{http_status}")
            return result_not_working(
                f"http_status={http_status};content_type={content_type};url={current_url}",
                confident=True,
            )

        title_before = await safe_title(page)
        text_before = await safe_page_text(page)
        text_problem_before = detect_problem_text(text_before, title_before)
        restricted_before = await inspect_restricted_access_overlay(page)
        visual_problem_before = await detect_visual_problem_marker(page)

        before_play_problem = first_definite_before_play_problem(
            text_problem_before,
            restricted_before,
            visual_problem_before,
        )

        if before_play_problem:
            await save_stream_debug(page, f"{debug_name}_definite_before_play")
            return result_not_working(
                (
                    f"definite_before_play={before_play_problem};"
                    f"text={text_problem_before};visual={visual_problem_before};"
                    f"restricted={restricted_before};title={title_before};url={current_url};"
                    f"goto_error={goto_error}"
                ),
                confident=True,
            )

        play_details = await try_start_playback_with_total_timeout(page, deep_mode=deep_mode)

        await page.wait_for_timeout(PLAY_SETTLE_WAIT_MS + (2500 if deep_mode else 0))
        await page.wait_for_timeout(MEDIA_COLLECT_WAIT_DEEP_MS if deep_mode else MEDIA_COLLECT_WAIT_MS)

        title_after = await safe_title(page)
        text_after = await safe_page_text(page)
        text_problem = detect_problem_text(text_after, title_after)
        restricted_problem = await inspect_restricted_access_overlay(page)
        visual_problem = await detect_visual_problem_marker(page)

        if (
            is_hard_failure_text(text_problem)
            or is_hard_failure_text(restricted_problem)
            or is_hard_failure_text(visual_problem)
        ):
            await save_stream_debug(page, f"{debug_name}_hard_text_after_play")
            return result_not_working(
                (
                    f"hard_text_after_play:"
                    f"text={text_problem};visual={visual_problem};restricted={restricted_problem};"
                    f"play={compact_json(play_details)};title={title_after};url={current_url};"
                    f"goto_error={goto_error}"
                ),
                confident=True,
            )

        dom_candidates = await collect_dom_media_candidates(page)
        for item in dom_candidates:
            collector.add(
                item.get("url"),
                source=item.get("source", "dom"),
                content_type=item.get("content_type", ""),
            )

        if looks_like_direct_media_url(stream_url, content_type):
            collector.add(stream_url, source="original_direct_url", content_type=content_type)

        media_candidates = collector.top_candidates(MAX_MEDIA_CANDIDATES)
        headers_info = await build_ffmpeg_headers(context, current_url or stream_url)

        ffmpeg_result = await ffmpeg_check_candidates(
            media_candidates=media_candidates,
            original_url=stream_url,
            headers_info=headers_info,
            debug_name=debug_name,
            deep_mode=deep_mode,
        )

        if ffmpeg_result["status"] == "working":
            return result_working(
                (
                    f"ffmpeg_working;"
                    f"goto_error={goto_error};"
                    f"ffmpeg={compact_json(ffmpeg_result)};"
                    f"media_candidates={compact_json(media_candidates)}"
                ),
                confident=True,
            )

        if ffmpeg_result["status"] == "not_working":
            await save_stream_debug(page, f"{debug_name}_ffmpeg_not_working")
            return result_not_working(
                (
                    f"ffmpeg_not_working;"
                    f"goto_error={goto_error};"
                    f"ffmpeg={compact_json(ffmpeg_result)};"
                    f"media_candidates={compact_json(media_candidates)}"
                ),
                confident=True,
            )

        browser_fallback = None

        if BROWSER_SCREENSHOT_FALLBACK:
            browser_fallback = await browser_visual_fallback(page, debug_name, deep_mode=deep_mode)

            if browser_fallback["status"] == "working":
                return result_working(
                    (
                        f"browser_fallback_working;"
                        f"goto_error={goto_error};"
                        f"ffmpeg={compact_json(ffmpeg_result)};"
                        f"browser={compact_json(browser_fallback)};"
                        f"media_candidates={compact_json(media_candidates)}"
                    ),
                    confident=deep_mode,
                )

            if browser_fallback["status"] == "not_working":
                await save_stream_debug(page, f"{debug_name}_browser_black")
                return result_not_working(
                    (
                        f"browser_fallback_not_working;"
                        f"goto_error={goto_error};"
                        f"ffmpeg={compact_json(ffmpeg_result)};"
                        f"browser={compact_json(browser_fallback)};"
                        f"media_candidates={compact_json(media_candidates)}"
                    ),
                    confident=True,
                )

        await save_stream_debug(page, f"{debug_name}_unknown")

        return result_unknown(
            (
                f"unknown_no_strong_evidence;"
                f"goto_error={goto_error};"
                f"ffmpeg={compact_json(ffmpeg_result)};"
                f"browser={compact_json(browser_fallback)};"
                f"text={text_problem};visual={visual_problem};restricted={restricted_problem};"
                f"play={compact_json(play_details)};"
                f"media_candidates={compact_json(media_candidates)};"
                f"title={title_after};url={current_url}"
            ),
            confident=deep_mode,
        )

    finally:
        await context.close()


def result_working(details, confident=True):
    return {
        "status": "working",
        "details": mask_sensitive_url(details),
        "confident": confident,
    }


def result_not_working(details, confident=True):
    return {
        "status": "not_working",
        "details": mask_sensitive_url(details),
        "confident": confident,
    }


def result_unknown(details, confident=False):
    return {
        "status": "unknown",
        "details": mask_sensitive_url(details),
        "confident": confident,
    }


class MediaCandidateCollector:
    def __init__(self):
        self.items = {}

    def add(self, url, source="", content_type="", status=None):
        if not url:
            return

        url = str(url).strip()

        if not url.startswith(("http://", "https://", "rtsp://", "rtmp://", "rtmps://")):
            return

        if is_noise_url(url):
            return

        score = score_media_candidate(url, content_type)

        if score <= 0:
            return

        key = normalize_candidate_url(url)

        item = {
            "url": url,
            "source": source,
            "content_type": content_type or "",
            "status": status,
            "score": score,
        }

        old = self.items.get(key)

        if old is None or item["score"] > old.get("score", 0):
            self.items[key] = item

    def top_candidates(self, limit):
        values = list(self.items.values())
        values.sort(key=lambda x: x.get("score", 0), reverse=True)
        return values[:limit]


def attach_media_collectors(page, collector):
    def on_request(request):
        try:
            collector.add(
                request.url,
                source=f"request:{request.resource_type}",
                content_type="",
            )
        except Exception:
            pass

    async def handle_response(response):
        try:
            headers = response.headers or {}
            collector.add(
                response.url,
                source="response",
                content_type=headers.get("content-type", ""),
                status=response.status,
            )
        except Exception:
            pass

    def on_response(response):
        asyncio.create_task(handle_response(response))

    page.on("request", on_request)
    page.on("response", on_response)


def is_noise_url(url):
    low = (url or "").lower()

    noise = [
        "google-analytics",
        "googletagmanager",
        "metrika",
        "yandex",
        "favicon",
        ".css",
        ".js",
        ".woff",
        ".woff2",
        ".ttf",
        ".svg",
        "/api/",
    ]

    return any(marker in low for marker in noise)


def score_media_candidate(url, content_type=""):
    low = (url or "").lower()
    ct = (content_type or "").lower()
    score = 0

    if low.startswith(("rtsp://", "rtmp://", "rtmps://")):
        score += 1000

    if ".m3u8" in low or "mpegurl" in ct:
        score += 200

    strong_url_markers = [
        "/hls/",
        "/live/",
        "/stream/",
        "/video/",
        "/mse/",
        ".mp4",
        ".flv",
        ".mjpeg",
        ".mjpg",
        "mjpeg",
        "mjpg",
    ]

    medium_url_markers = [
        ".ts",
        "/chunk",
        "/segment",
        "/media",
        "/camera",
        "/cam",
        "/snapshot",
        ".jpg",
        ".jpeg",
        ".png",
    ]

    content_markers = [
        "application/vnd.apple.mpegurl",
        "application/x-mpegurl",
        "mpegurl",
        "video/",
        "multipart/x-mixed-replace",
        "image/jpeg",
        "image/png",
    ]

    if any(marker in low for marker in strong_url_markers):
        score += 80

    if any(marker in low for marker in medium_url_markers):
        score += 30

    if any(marker in ct for marker in content_markers):
        score += 100

    if ".ts" in low:
        score -= 20

    if low.startswith("blob:"):
        score = 0

    return score


def normalize_candidate_url(url):
    return re.sub(r"([?&])_=\d+", "", url or "")


async def collect_dom_media_candidates(page):
    try:
        return await page.evaluate(
            """
            () => {
                const result = [];

                const add = (url, source, content_type='') => {
                    if (!url || typeof url !== 'string') return;
                    if (
                        !url.startsWith('http://') &&
                        !url.startsWith('https://') &&
                        !url.startsWith('rtsp://')
                    ) return;

                    result.push({url, source, content_type});
                };

                for (const video of Array.from(document.querySelectorAll('video'))) {
                    add(video.currentSrc || video.src || '', 'dom_video_src');

                    for (const source of Array.from(video.querySelectorAll('source'))) {
                        add(source.src || '', 'dom_video_source', source.type || '');
                    }
                }

                for (const source of Array.from(document.querySelectorAll('source'))) {
                    add(source.src || '', 'dom_source', source.type || '');
                }

                for (const img of Array.from(document.querySelectorAll('img'))) {
                    add(img.currentSrc || img.src || '', 'dom_img');
                }

                for (const iframe of Array.from(document.querySelectorAll('iframe'))) {
                    add(iframe.src || '', 'dom_iframe');
                }

                for (const objectEl of Array.from(document.querySelectorAll('object'))) {
                    add(objectEl.data || '', 'dom_object');
                }

                for (const embed of Array.from(document.querySelectorAll('embed'))) {
                    add(embed.src || '', 'dom_embed');
                }

                return result;
            }
            """
        )
    except Exception:
        return []


async def build_ffmpeg_headers(context, referer_url):
    try:
        cookies = await context.cookies()
    except Exception:
        cookies = []

    cookie_line = "; ".join(
        f"{item.get('name')}={item.get('value')}"
        for item in cookies
        if item.get("name") and item.get("value") is not None
    )

    user_agent = default_user_agent()
    origin = get_origin(referer_url)

    header_lines = [
        f"Referer: {referer_url}",
        f"Origin: {origin}",
        f"User-Agent: {user_agent}",
        "Accept: */*",
        "Connection: keep-alive",
    ]

    if cookie_line:
        header_lines.append(f"Cookie: {cookie_line}")

    return {
        "referer": referer_url,
        "origin": origin,
        "user_agent": user_agent,
        "cookie": cookie_line,
        "headers": "\r\n".join(header_lines) + "\r\n",
    }


def build_direct_ffmpeg_headers(referer_url=""):
    user_agent = default_user_agent()
    origin = get_origin(referer_url)

    header_lines = [
        f"Referer: {referer_url}",
        f"Origin: {origin}",
        f"User-Agent: {user_agent}",
        "Accept: */*",
        "Connection: keep-alive",
    ]

    return {
        "referer": referer_url,
        "origin": origin,
        "user_agent": user_agent,
        "cookie": "",
        "headers": "\r\n".join(header_lines) + "\r\n",
    }


def default_user_agent():
    return (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )


def get_origin(url):
    try:
        parsed = urlparse(url)
        if parsed.scheme and parsed.netloc:
            return f"{parsed.scheme}://{parsed.netloc}"
    except Exception:
        pass

    return ""


async def ffmpeg_check_candidates(media_candidates, original_url, headers_info, debug_name, deep_mode=False):
    if not FFMPEG_ENABLED:
        return {
            "status": "unknown",
            "reason": "ffmpeg_disabled",
            "checked": [],
        }

    if not PIL_OK:
        return {
            "status": "unknown",
            "reason": "pil_missing",
            "checked": [],
        }

    candidates = list(media_candidates or [])

    if is_direct_ffmpeg_url(original_url):
        candidates.insert(
            0,
            {
                "url": original_url,
                "source": "original_direct_url",
                "content_type": "",
                "score": 999,
            },
        )

    if not candidates:
        return {
            "status": "unknown",
            "reason": "no_media_candidates",
            "checked": [],
        }

    checked = []
    seen = set()

    async with FFMPEG_SEMAPHORE:
        for index, candidate in enumerate(candidates[:MAX_MEDIA_CANDIDATES], start=1):
            url = candidate.get("url")

            if not url or url in seen:
                continue

            seen.add(url)

            candidate_debug = f"{debug_name}_candidate_{index}"

            probe = await run_ffprobe(url, headers_info)
            frames = await extract_frames_with_ffmpeg(url, headers_info, candidate_debug, deep_mode=deep_mode)
            frame_analysis = analyze_extracted_frames(frames.get("frames", []))

            item = {
                "candidate": candidate,
                "probe": probe,
                "frames": frames,
                "analysis": frame_analysis,
            }

            decision = decide_ffmpeg_candidate(probe, frames, frame_analysis, candidate)
            item["decision"] = decision
            checked.append(item)

            if decision["status"] == "working":
                return {
                    "status": "working",
                    "reason": decision["reason"],
                    "winner": item,
                    "checked": checked,
                }

            if decision["status"] == "not_working" and decision.get("confident"):
                return {
                    "status": "not_working",
                    "reason": decision["reason"],
                    "winner": item,
                    "checked": checked,
                }

    if checked and all((item.get("decision") or {}).get("status") == "not_working" for item in checked):
        return {
            "status": "not_working",
            "reason": "all_candidates_failed_or_black",
            "checked": checked,
        }

    return {
        "status": "unknown",
        "reason": "no_candidate_confirmed",
        "checked": checked,
    }


async def run_ffprobe(url, headers_info):
    if not shutil.which(FFPROBE_BIN):
        return {
            "ok": False,
            "error": f"ffprobe_not_found:{FFPROBE_BIN}",
        }

    input_options = []

    if (url or "").lower().startswith("rtsp://"):
        input_options.extend(["-rtsp_transport", FFMPEG_RTSP_TRANSPORT])

    cmd = [
        FFPROBE_BIN,
        "-v",
        "error",
        "-rw_timeout",
        str(FFMPEG_RW_TIMEOUT_US),
        *input_options,
        "-user_agent",
        headers_info.get("user_agent", ""),
        "-headers",
        headers_info.get("headers", ""),
        "-show_streams",
        "-show_format",
        "-of",
        "json",
        url,
    ]

    code, stdout, stderr, timeout = await run_process(cmd, timeout=FFPROBE_TIMEOUT_SEC)

    result = {
        "ok": code == 0 and bool(stdout.strip()),
        "returncode": code,
        "timeout": timeout,
        "stderr": stderr[-1200:],
    }

    if stdout.strip():
        try:
            data = json.loads(stdout)
            result["json"] = data
            result["video_streams"] = [
                s for s in data.get("streams", []) if s.get("codec_type") == "video"
            ]
        except Exception as e:
            result["json_error"] = f"{type(e).__name__}:{e}"

    return result


async def extract_frames_with_ffmpeg(url, headers_info, debug_name, deep_mode=False):
    if not shutil.which(FFMPEG_BIN):
        return {
            "ok": False,
            "error": f"ffmpeg_not_found:{FFMPEG_BIN}",
            "frames": [],
        }

    out_dir = FFMPEG_DEBUG_DIR / sanitize_filename(debug_name)
    out_dir.mkdir(parents=True, exist_ok=True)

    for old in out_dir.glob("frame_*.jpg"):
        try:
            old.unlink()
        except Exception:
            pass

    frame_pattern = str(out_dir / "frame_%03d.jpg")
    frame_count = FFMPEG_FRAME_COUNT + (3 if deep_mode else 0)

    input_options = []

    if (url or "").lower().startswith("rtsp://"):
        input_options.extend(["-rtsp_transport", FFMPEG_RTSP_TRANSPORT])

    cmd = [
        FFMPEG_BIN,
        "-y",
        "-nostdin",
        "-hide_banner",
        "-loglevel",
        "error",
        "-rw_timeout",
        str(FFMPEG_RW_TIMEOUT_US),
        *input_options,
        "-user_agent",
        headers_info.get("user_agent", ""),
        "-headers",
        headers_info.get("headers", ""),
        "-i",
        url,
        "-an",
        "-vf",
        "fps=1,scale='min(640,iw)':-2",
        "-frames:v",
        str(frame_count),
        frame_pattern,
    ]

    code, stdout, stderr, timeout = await run_process(
        cmd,
        timeout=FFMPEG_TIMEOUT_SEC + (8 if deep_mode else 0),
    )

    frames = sorted(str(p) for p in out_dir.glob("frame_*.jpg"))

    return {
        "ok": code == 0 and len(frames) > 0,
        "returncode": code,
        "timeout": timeout,
        "stderr": stderr[-1800:],
        "frames": frames,
        "frame_count": len(frames),
        "out_dir": str(out_dir),
    }


async def run_process(cmd, timeout):
    try:
        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )

        try:
            stdout_b, stderr_b = await asyncio.wait_for(process.communicate(), timeout=timeout)
            return (
                process.returncode,
                stdout_b.decode("utf-8", errors="ignore"),
                stderr_b.decode("utf-8", errors="ignore"),
                False,
            )
        except asyncio.TimeoutError:
            try:
                process.kill()
            except Exception:
                pass

            try:
                stdout_b, stderr_b = await process.communicate()
            except Exception:
                stdout_b, stderr_b = b"", b""

            return (
                -999,
                stdout_b.decode("utf-8", errors="ignore"),
                stderr_b.decode("utf-8", errors="ignore") + f"\nTIMEOUT>{timeout}s",
                True,
            )

    except Exception as e:
        return -998, "", f"{type(e).__name__}:{e}", False


def analyze_extracted_frames(frame_paths):
    if not PIL_OK or not frame_paths:
        return {
            "ok": False,
            "reason": "no_frames_or_pil_missing",
            "frames": [],
        }

    stats = []

    for path in frame_paths:
        stats.append(analyze_frame_pixels(path))

    diffs = []

    for i in range(len(frame_paths) - 1):
        diffs.append(compare_frame_motion(frame_paths[i], frame_paths[i + 1]))

    summary = summarize_frame_stats(stats, diffs)

    return {
        "ok": True,
        "frames": stats,
        "diffs": diffs,
        "summary": summary,
    }


def analyze_frame_pixels(path):
    try:
        image = Image.open(path).convert("RGB")
    except Exception as e:
        return {
            "path": path,
            "error": f"{type(e).__name__}:{e}",
            "black_ratio": 1,
            "center_black_ratio": 1,
            "bright_ratio": 0,
            "luma_mean": 0,
            "luma_std": 0,
            "entropy": 0,
            "colorfulness": 0,
            "hard_black": True,
            "mostly_black": True,
            "usable_picture": False,
        }

    width, height = image.size

    crop = image.crop(
        (
            int(width * 0.02),
            int(height * 0.03),
            int(width * 0.98),
            int(height * 0.97),
        )
    )

    cw, ch = crop.size
    center = crop.crop(
        (
            int(cw * 0.15),
            int(ch * 0.15),
            int(cw * 0.85),
            int(ch * 0.85),
        )
    )

    full = compute_pixel_stats(crop)
    center_stats = compute_pixel_stats(center)

    hard_black = (
        full["black_ratio"] >= FRAME_BLACK_RATIO
        and center_stats["black_ratio"] >= FRAME_CENTER_BLACK_RATIO
        and full["luma_mean"] <= FRAME_MAX_MEAN_LUMA
        and full["bright_ratio"] <= FRAME_MAX_BRIGHT_RATIO
        and full["entropy"] <= FRAME_MAX_ENTROPY
    )

    mostly_black = (
        full["black_ratio"] >= 0.78
        and center_stats["black_ratio"] >= 0.82
        and full["luma_mean"] <= 34
        and full["bright_ratio"] <= 0.065
    )

    usable_picture = (
        full["luma_mean"] >= 30
        and full["bright_ratio"] >= 0.035
        and full["entropy"] >= 2.0
        and full["black_ratio"] < 0.88
    ) or (
        full["entropy"] >= 3.0
        and full["black_ratio"] < 0.82
    )

    return {
        "path": path,
        "width": width,
        "height": height,
        "black_ratio": round(full["black_ratio"], 4),
        "dark_ratio": round(full["dark_ratio"], 4),
        "center_black_ratio": round(center_stats["black_ratio"], 4),
        "bright_ratio": round(full["bright_ratio"], 4),
        "luma_mean": round(full["luma_mean"], 4),
        "luma_std": round(full["luma_std"], 4),
        "entropy": round(full["entropy"], 4),
        "colorfulness": round(full["colorfulness"], 4),
        "hard_black": hard_black,
        "mostly_black": mostly_black,
        "usable_picture": usable_picture,
    }


def compute_pixel_stats(image):
    max_side = 420
    width, height = image.size

    if width > max_side:
        new_width = max_side
        new_height = max(1, int(height * new_width / width))
        image = image.resize((new_width, new_height))

    rgb = image.convert("RGB")
    gray = ImageOps.grayscale(rgb)

    rgb_pixels = list(rgb.getdata())
    gray_pixels = list(gray.getdata())
    total = max(1, len(gray_pixels))

    black = 0
    dark = 0
    bright = 0
    color_sum = 0

    for idx, y in enumerate(gray_pixels):
        if y <= 22:
            black += 1
        if y <= 42:
            dark += 1
        if y >= 70:
            bright += 1

        r, g, b = rgb_pixels[idx]
        color_sum += max(r, g, b) - min(r, g, b)

    stat = ImageStat.Stat(gray)

    return {
        "black_ratio": black / total,
        "dark_ratio": dark / total,
        "bright_ratio": bright / total,
        "luma_mean": float(stat.mean[0]) if stat.mean else 0.0,
        "luma_std": float(stat.stddev[0]) if stat.stddev else 0.0,
        "entropy": compute_gray_entropy(gray_pixels),
        "colorfulness": color_sum / total,
    }


def compute_gray_entropy(gray_pixels):
    if not gray_pixels:
        return 0.0

    counts = [0] * 256

    for value in gray_pixels:
        counts[int(value)] += 1

    total = len(gray_pixels)
    entropy = 0.0

    for count in counts:
        if count <= 0:
            continue

        p = count / total
        entropy -= p * math.log2(p)

    return entropy


def compare_frame_motion(path1, path2):
    try:
        img1 = Image.open(path1).convert("RGB")
        img2 = Image.open(path2).convert("RGB")

        width = min(img1.size[0], img2.size[0])
        height = min(img1.size[1], img2.size[1])

        img1 = img1.crop((0, 0, width, height))
        img2 = img2.crop((0, 0, width, height))

        if width > 420:
            new_width = 420
            new_height = max(1, int(height * new_width / width))
            img1 = img1.resize((new_width, new_height))
            img2 = img2.resize((new_width, new_height))

        g1 = ImageOps.grayscale(img1)
        g2 = ImageOps.grayscale(img2)
        diff = ImageChops.difference(g1, g2)

        pixels = list(diff.getdata())
        total = max(1, len(pixels))

        mean_diff = sum(pixels) / total
        changed_ratio = sum(1 for p in pixels if p > 9) / total

        return {
            "mean_diff": round(mean_diff, 4),
            "changed_ratio": round(changed_ratio, 6),
        }

    except Exception as e:
        return {
            "error": f"{type(e).__name__}:{e}",
            "mean_diff": 0,
            "changed_ratio": 0,
        }


def summarize_frame_stats(stats, diffs):
    def avg(key):
        values = []

        for item in stats:
            try:
                values.append(float(item.get(key) or 0))
            except Exception:
                values.append(0.0)

        return sum(values) / max(1, len(values))

    frame_count = len(stats)
    hard_black_frames = sum(1 for s in stats if s.get("hard_black"))
    mostly_black_frames = sum(1 for s in stats if s.get("mostly_black"))
    usable_frames = sum(1 for s in stats if s.get("usable_picture"))

    mean_motion = sum(float(d.get("mean_diff") or 0) for d in diffs) / max(1, len(diffs))
    changed_motion = sum(float(d.get("changed_ratio") or 0) for d in diffs) / max(1, len(diffs))

    static = (
        mean_motion <= FRAME_STATIC_MEAN_DIFF_MAX
        and changed_motion <= FRAME_STATIC_CHANGED_RATIO_MAX
    )

    return {
        "frame_count": frame_count,
        "hard_black_frames": hard_black_frames,
        "mostly_black_frames": mostly_black_frames,
        "usable_frames": usable_frames,
        "avg_black_ratio": round(avg("black_ratio"), 4),
        "avg_center_black_ratio": round(avg("center_black_ratio"), 4),
        "avg_bright_ratio": round(avg("bright_ratio"), 4),
        "avg_luma_mean": round(avg("luma_mean"), 4),
        "avg_luma_std": round(avg("luma_std"), 4),
        "avg_entropy": round(avg("entropy"), 4),
        "avg_colorfulness": round(avg("colorfulness"), 4),
        "mean_motion": round(mean_motion, 4),
        "changed_motion": round(changed_motion, 6),
        "static": static,
    }


def decide_ffmpeg_candidate(probe, frames, analysis, candidate):
    stderr = f"{probe.get('stderr', '')} {frames.get('stderr', '')}".lower()

    hard_error_markers = [
        "403",
        "404",
        "401",
        "forbidden",
        "unauthorized",
        "not found",
        "server returned 5",
        "connection refused",
        "invalid data found",
        "http error",
        "method describe failed",
        "401 unauthorized",
    ]

    if any(marker in stderr for marker in hard_error_markers):
        return {
            "status": "not_working",
            "confident": True,
            "reason": f"ffmpeg_hard_error:{first_matching(stderr, hard_error_markers)}",
        }

    if frames.get("frame_count", 0) <= 0:
        if is_direct_ffmpeg_url(candidate.get("url")):
            return {
                "status": "not_working",
                "confident": False,
                "reason": "direct_media_no_frames",
            }

        return {
            "status": "unknown",
            "confident": False,
            "reason": "no_frames",
        }

    summary = (analysis or {}).get("summary") or {}

    frame_count = int(summary.get("frame_count") or 0)
    usable_frames = int(summary.get("usable_frames") or 0)
    hard_black_frames = int(summary.get("hard_black_frames") or 0)
    mostly_black_frames = int(summary.get("mostly_black_frames") or 0)
    static = bool(summary.get("static"))

    avg_black = float(summary.get("avg_black_ratio") or 0)
    avg_center_black = float(summary.get("avg_center_black_ratio") or 0)
    avg_bright = float(summary.get("avg_bright_ratio") or 0)
    avg_mean = float(summary.get("avg_luma_mean") or 0)
    avg_entropy = float(summary.get("avg_entropy") or 0)

    if usable_frames >= max(2, min(3, frame_count)):
        return {
            "status": "working",
            "confident": True,
            "reason": (
                f"usable_decoded_frames={usable_frames}/{frame_count};"
                f"black={avg_black};bright={avg_bright};mean={avg_mean};entropy={avg_entropy}"
            ),
        }

    if frame_count >= 2 and hard_black_frames >= max(2, frame_count - 1):
        return {
            "status": "not_working",
            "confident": True,
            "reason": (
                f"ffmpeg_hard_black_frames={hard_black_frames}/{frame_count};"
                f"black={avg_black};center={avg_center_black};"
                f"bright={avg_bright};mean={avg_mean};entropy={avg_entropy}"
            ),
        }

    if (
        frame_count >= 3
        and mostly_black_frames >= frame_count - 1
        and static
        and avg_black >= 0.80
        and avg_center_black >= 0.84
        and avg_mean <= 34
        and avg_bright <= 0.06
    ):
        return {
            "status": "not_working",
            "confident": True,
            "reason": (
                f"ffmpeg_mostly_black_static={mostly_black_frames}/{frame_count};"
                f"black={avg_black};center={avg_center_black};"
                f"bright={avg_bright};mean={avg_mean};entropy={avg_entropy}"
            ),
        }

    if frame_count >= 3 and avg_black < 0.76 and avg_entropy >= 2.2 and avg_bright >= 0.025:
        return {
            "status": "working",
            "confident": True,
            "reason": (
                f"decoded_non_black_frames={frame_count};"
                f"black={avg_black};bright={avg_bright};mean={avg_mean};entropy={avg_entropy}"
            ),
        }

    return {
        "status": "unknown",
        "confident": False,
        "reason": (
            f"decoded_but_uncertain;"
            f"frames={frame_count};usable={usable_frames};black_frames={hard_black_frames};"
            f"mostly_black={mostly_black_frames};static={static};"
            f"black={avg_black};bright={avg_bright};mean={avg_mean};entropy={avg_entropy}"
        ),
    }


async def browser_visual_fallback(page, debug_name, deep_mode=False):
    if not PIL_OK:
        return {"status": "unknown", "reason": "pil_missing"}

    files = []

    try:
        clip = await find_best_player_clip(page)
        delays = [0, 1500, 3000] if deep_mode else [0, 1500]

        for i, delay in enumerate(delays, start=1):
            if delay:
                await page.wait_for_timeout(delay)

            path = STREAM_DEBUG_DIR / f"{debug_name}_browser_fallback_{i}.png"

            if clip:
                await page.screenshot(path=str(path), full_page=False, clip=clip)
            else:
                await page.screenshot(path=str(path), full_page=False)

            files.append(str(path))

        analysis = analyze_extracted_frames(files)
        summary = analysis.get("summary") or {}

        usable = int(summary.get("usable_frames") or 0)
        avg_black = float(summary.get("avg_black_ratio") or 1)
        hard_black = int(summary.get("hard_black_frames") or 0)
        mostly_black = int(summary.get("mostly_black_frames") or 0)
        static = bool(summary.get("static"))

        if usable >= 1 and avg_black < 0.80:
            return {
                "status": "working",
                "reason": "browser_visible_non_black_picture",
                "analysis": analysis,
                "files": files,
            }

        if hard_black >= max(2, len(files) - 1) or (
            mostly_black >= max(2, len(files) - 1) and static
        ):
            return {
                "status": "not_working",
                "reason": "browser_black_static_picture",
                "analysis": analysis,
                "files": files,
            }

        return {
            "status": "unknown",
            "reason": "browser_uncertain",
            "analysis": analysis,
            "files": files,
        }

    except Exception as e:
        return {
            "status": "unknown",
            "reason": f"browser_fallback_error:{type(e).__name__}:{e}",
            "files": files,
        }


async def find_best_player_clip(page):
    try:
        return await page.evaluate(
            """
            () => {
                const selectors = [
                    'video',
                    'canvas',
                    'iframe',
                    '.video-js',
                    '.jwplayer',
                    '.plyr',
                    '[class*="player"]',
                    '[class*="video"]',
                    '[id*="player"]',
                    '[id*="video"]'
                ];

                const vw = window.innerWidth || 1280;
                const vh = window.innerHeight || 900;

                let best = null;
                let bestArea = 0;

                for (const selector of selectors) {
                    const nodes = Array.from(document.querySelectorAll(selector));

                    for (const node of nodes) {
                        try {
                            const style = window.getComputedStyle(node);

                            if (!style || style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') {
                                continue;
                            }

                            const r = node.getBoundingClientRect();

                            const x = Math.max(0, r.left);
                            const y = Math.max(0, r.top);
                            const width = Math.max(0, Math.min(r.width, vw - x));
                            const height = Math.max(0, Math.min(r.height, vh - y));
                            const area = width * height;

                            if (width < 180 || height < 100 || area < 20000) {
                                continue;
                            }

                            if (area > bestArea) {
                                bestArea = area;
                                best = {
                                    x: Math.floor(x),
                                    y: Math.floor(y),
                                    width: Math.floor(width),
                                    height: Math.floor(height)
                                };
                            }
                        } catch (e) {}
                    }
                }

                if (best && best.width > 0 && best.height > 0) {
                    return best;
                }

                return {
                    x: 0,
                    y: Math.floor(vh * 0.08),
                    width: vw,
                    height: Math.floor(vh * 0.78)
                };
            }
            """
        )
    except Exception:
        return None


async def try_start_playback_with_total_timeout(page, deep_mode=False):
    if not PLAY_ENABLED:
        return {"enabled": False, "unresponsive": False}

    try:
        return await asyncio.wait_for(
            try_start_playback(page, deep_mode=deep_mode),
            timeout=PLAY_TOTAL_TIMEOUT_MS / 1000,
        )
    except asyncio.TimeoutError:
        return {
            "enabled": True,
            "total_timeout": True,
            "unresponsive": True,
            "errors": [f"play_total_timeout>{PLAY_TOTAL_TIMEOUT_MS}ms"],
        }
    except Exception as e:
        return {
            "enabled": True,
            "total_error": f"{type(e).__name__}:{e}",
            "unresponsive": True,
            "errors": [f"play_total_error:{type(e).__name__}:{e}"],
        }


async def try_start_playback(page, deep_mode=False):
    details = {
        "enabled": True,
        "js_video_count": 0,
        "js_play_ok_count": 0,
        "clicked_selectors": [],
        "clicked_center": False,
        "errors": [],
    }

    try:
        result = await page.evaluate(
            """
            async (timeoutMs) => {
                const videos = Array.from(document.querySelectorAll('video'));
                let ok = 0;
                const errors = [];

                const timeout = ms => new Promise((_, reject) => {
                    setTimeout(() => reject(new Error('play_timeout')), ms);
                });

                for (const video of videos) {
                    try {
                        video.muted = true;
                        video.volume = 0;
                        video.autoplay = true;
                        video.playsInline = true;
                        video.setAttribute('playsinline', 'true');
                        video.setAttribute('webkit-playsinline', 'true');

                        const p = video.play();

                        if (p && typeof p.then === 'function') {
                            await Promise.race([p, timeout(timeoutMs)]);
                        }

                        ok++;
                    } catch (e) {
                        errors.push(String(e && e.message ? e.message : e));
                    }
                }

                return {video_count: videos.length, ok, errors};
            }
            """,
            PLAY_JS_TIMEOUT_MS,
        )

        details["js_video_count"] = result.get("video_count", 0)
        details["js_play_ok_count"] = result.get("ok", 0)
        details["errors"].extend(result.get("errors", []))

    except Exception as e:
        details["errors"].append(f"js_play_error:{type(e).__name__}:{e}")

    await page.wait_for_timeout(700)

    selectors = [
        ".vjs-big-play-button",
        ".vjs-play-control",
        ".plyr__control[data-plyr='play']",
        ".jw-icon-playback",
        ".jw-display-icon-container",
        "button:has-text('Play')",
        "button:has-text('PLAY')",
        "button:has-text('Воспроизвести')",
        "button[aria-label*='Play']",
        "button[aria-label*='play']",
        "button[aria-label*='Воспроизвести']",
        "[class*='big-play']",
        "[class*='play-button']",
        "[class*='play']",
        "[id*='play']",
    ]

    clicked = 0

    for selector in selectors:
        if clicked >= 5:
            break

        try:
            locator = page.locator(selector).first

            if await locator.count() == 0:
                continue

            if not await locator.is_visible():
                continue

            await locator.click(timeout=1000, force=True)
            details["clicked_selectors"].append(selector)
            clicked += 1
            await page.wait_for_timeout(700)

        except Exception:
            continue

    try:
        point = await page.evaluate(
            """
            () => {
                const candidates = [
                    document.querySelector('video'),
                    document.querySelector('canvas'),
                    document.querySelector('iframe'),
                    document.querySelector('[class*="player"]'),
                    document.querySelector('[class*="video"]')
                ].filter(Boolean);

                for (const el of candidates) {
                    const r = el.getBoundingClientRect();

                    if (r.width > 120 && r.height > 80) {
                        return {x: r.left + r.width / 2, y: r.top + r.height / 2};
                    }
                }

                return {x: window.innerWidth / 2, y: window.innerHeight / 2};
            }
            """
        )

        await page.mouse.click(float(point["x"]), float(point["y"]))
        details["clicked_center"] = True
        await page.wait_for_timeout(1000)

    except Exception as e:
        details["errors"].append(f"center_click_error:{type(e).__name__}:{e}")

    return details


async def detect_visual_problem_marker(page):
    try:
        result = await page.evaluate(
            """
            () => {
                const texts = [];

                const pushText = value => {
                    const t = (value || '').toString().trim();
                    if (t) texts.push(t.toLowerCase());
                };

                const selectors = [
                    'body',
                    '[class*="error"]',
                    '[class*="offline"]',
                    '[class*="warning"]',
                    '[class*="message"]',
                    '[class*="status"]',
                    '[class*="notice"]',
                    '[class*="overlay"]',
                    '[class*="modal"]',
                    '.vjs-error-display',
                    '.jw-error',
                    '.jw-text',
                    '.plyr'
                ];

                for (const selector of selectors) {
                    const nodes = document.querySelectorAll(selector);

                    for (const node of nodes) {
                        try {
                            const style = window.getComputedStyle(node);
                            const visible = style && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';

                            if (!visible) continue;

                            pushText(node.innerText);
                            pushText(node.textContent);
                            pushText(node.getAttribute('aria-label'));
                            pushText(node.getAttribute('title'));
                            pushText(node.getAttribute('alt'));
                            pushText(node.getAttribute('data-title'));
                            pushText(node.getAttribute('data-error'));
                            pushText(node.getAttribute('data-message'));
                        } catch (e) {}
                    }
                }

                return texts.join(' | ');
            }
            """
        )

        return detect_problem_text(result)

    except Exception:
        return None


async def inspect_restricted_access_overlay(page):
    try:
        info = await page.evaluate(
            """
            () => {
                const parts = [];

                const add = value => {
                    const text = (value || '').toString().trim();
                    if (text) parts.push(text);
                };

                add(document.title || '');
                add(document.body ? (document.body.innerText || '') : '');
                add(document.body ? (document.body.textContent || '') : '');

                const nodes = document.querySelectorAll('*');

                for (const node of nodes) {
                    try {
                        const style = window.getComputedStyle(node);
                        const rect = node.getBoundingClientRect();

                        const visible =
                            style &&
                            style.display !== 'none' &&
                            style.visibility !== 'hidden' &&
                            style.opacity !== '0' &&
                            rect.width > 0 &&
                            rect.height > 0;

                        if (!visible) continue;

                        add(node.innerText);
                        add(node.textContent);
                        add(node.getAttribute('aria-label'));
                        add(node.getAttribute('title'));
                        add(node.getAttribute('alt'));
                        add(node.getAttribute('data-title'));
                        add(node.getAttribute('data-message'));
                        add(node.getAttribute('data-error'));
                    } catch (e) {}
                }

                return parts.join(' | ');
            }
            """
        )

        normalized = normalize_ocr_text(info)

        if "доступ к трансляции временно ограничен" in normalized:
            return "доступ к трансляции временно ограничен"

        if "трансляции временно ограничен" in normalized:
            return "доступ к трансляции временно ограничен"

        if "временно ограничен" in normalized and "трансляции" in normalized:
            return "доступ к трансляции временно ограничен"

        if "временно ограничен" in normalized and "баланс" in normalized:
            return "доступ к трансляции временно ограничен / пополнить баланс"

        if "пополнить баланс" in normalized:
            return "пополнить баланс"

        if "пополните баланс" in normalized:
            return "пополнить баланс"

        if "личном кабинете" in normalized and "пополнить" in normalized and "баланс" in normalized:
            return "подробнее в личном кабинете / пополнить баланс"

        return None

    except Exception:
        return None


def normalize_text(value):
    return re.sub(r"\s+", " ", (value or "").strip().lower())


def normalize_ocr_text(value):
    text = normalize_text(value).replace("ё", "е")
    text = text.replace("|", " ")
    text = text.replace(".", " ")
    text = re.sub(r"[^a-zа-я0-9\s]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def detect_problem_text(page_text_norm, title_norm=""):
    combined = normalize_ocr_text(f"{page_text_norm} | {title_norm}")

    markers = [
        "доступ к трансляции временно ограничен",
        "трансляции временно ограничен",
        "доступ временно ограничен",
        "временно ограничен",
        "пополнить баланс",
        "пополните баланс",
        "stream unavailable",
        "access denied",
        "forbidden",
        "not found",
        "404",
        "403",
        "401",
        "502",
        "504",
        "bad gateway",
        "gateway timeout",
        "unable to play",
        "failed to load",
        "media could not be loaded",
        "no signal",
        "no video",
        "camera offline",
        "stream offline",
        "source offline",
        "device offline",
        "offline",
        "service unavailable",
        "камера недоступна",
        "поток недоступен",
        "нет сигнала",
        "нет видео",
        "видео недоступно",
        "источник недоступен",
        "канал недоступен",
        "stream not found",
        "unable to connect",
        "cannot connect",
        "unable to connect to camera",
        "check that the camera is turned on",
        "check camera connection",
        "невозможно подключиться к камере",
        "проверьте что она включена",
        "подключена к интернету",
        "проверьте подключение камеры к интернету",
        "camera is not available",
        "camera unavailable",
        "unauthorized",
        "access is denied",
        "invalid password",
    ]

    for marker in markers:
        if marker in combined:
            return marker

    return None


def is_hard_failure_text(value):
    if not value:
        return False

    value = normalize_ocr_text(value)

    hard_markers = [
        "доступ к трансляции временно ограничен",
        "трансляции временно ограничен",
        "доступ временно ограничен",
        "временно ограничен",
        "пополнить баланс",
        "пополните баланс",
        "невозможно подключиться к камере",
        "проверьте что она включена",
        "подключена к интернету",
        "проверьте подключение камеры к интернету",
        "камера недоступна",
        "поток недоступен",
        "источник недоступен",
        "канал недоступен",
        "нет сигнала",
        "нет видео",
        "видео недоступно",
        "unable to connect to camera",
        "cannot connect",
        "unable to connect",
        "check that the camera is turned on",
        "check camera connection",
        "no video",
        "no signal",
        "camera offline",
        "stream offline",
        "device offline",
        "source offline",
        "offline",
        "camera unavailable",
        "stream unavailable",
        "stream not found",
        "media could not be loaded",
        "failed to load",
        "access denied",
        "forbidden",
        "unauthorized",
        "service unavailable",
        "bad gateway",
        "gateway timeout",
    ]

    return any(marker in value for marker in hard_markers)


def is_definite_before_play_text(value):
    if not value:
        return False

    value = normalize_ocr_text(value)

    definite_markers = [
        "доступ к трансляции временно ограничен",
        "трансляции временно ограничен",
        "доступ временно ограничен",
        "временно ограничен",
        "пополнить баланс",
        "пополните баланс",
        "невозможно подключиться к камере",
        "нет видео",
        "нет сигнала",
        "no video",
        "no signal",
        "unable to connect to camera",
    ]

    return any(marker in value for marker in definite_markers)


def first_definite_before_play_problem(*values):
    for value in values:
        if is_definite_before_play_text(value):
            return value

    return None


def looks_like_direct_media_url(stream_url, content_type=""):
    url = (stream_url or "").lower().strip()
    ct = (content_type or "").lower().strip()

    if url.startswith(("rtsp://", "rtmp://", "rtmps://", "srt://", "udp://", "tcp://")):
        return True

    media_markers = [
        ".mjpg",
        ".mjpeg",
        ".jpg",
        ".jpeg",
        ".png",
        ".mp4",
        ".m3u8",
        ".ts",
        ".flv",
        "/video",
        "/stream",
        "/live",
        "/snapshot",
        "mjpeg",
        "mjpg",
        "hls",
    ]

    if any(marker in url for marker in media_markers):
        return True

    if any(
        marker in ct
        for marker in [
            "video/",
            "image/",
            "multipart/x-mixed-replace",
            "application/vnd.apple.mpegurl",
            "application/x-mpegurl",
        ]
    ):
        return True

    return False


def is_direct_ffmpeg_url(url):
    url = (url or "").lower().strip()

    if url.startswith(("rtsp://", "rtmp://", "rtmps://", "srt://", "udp://", "tcp://")):
        return True

    return looks_like_direct_media_url(url, "")


def first_matching(text, markers):
    for marker in markers:
        if marker in text:
            return marker

    return None


def compact_json(value, limit=2500):
    try:
        text = json.dumps(value, ensure_ascii=False, default=str)
    except Exception:
        text = str(value)

    text = mask_sensitive_url(text)

    if len(text) > limit:
        return text[:limit] + "...<cut>"

    return text


def mask_sensitive_url(value):
    if not value:
        return value

    text = str(value)

    text = re.sub(
        r"(rtsp://)([^:/@\s]+):([^@\s]+)@",
        r"\1\2:***@",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"(rtmp://)([^:/@\s]+):([^@\s]+)@",
        r"\1\2:***@",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"(https?://)([^:/@\s]+):([^@\s]+)@",
        r"\1\2:***@",
        text,
        flags=re.IGNORECASE,
    )

    return text


def normalize_address(value):
    value = (value or "").strip().lower().replace("ё", "е")
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"[\"'`]", "", value)
    value = re.sub(r"\bул\.\b", "улица", value)
    value = re.sub(r"\bпр-кт\b", "проспект", value)
    value = re.sub(r"\bд\.\b", "дом", value)
    return value.strip()


def load_addresses_from_tsv(file_path):
    if not file_path.exists():
        log(f"Файл таблицы адресов не найден: {file_path}")
        return []

    addresses = []
    seen = set()

    with file_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter="\t")
        rows = list(reader)

    if not rows:
        return []

    header = [str(cell).strip().lower() for cell in rows[0]]
    address_col_index = None

    for i, name in enumerate(header):
        if name in {"адрес", "address"}:
            address_col_index = i
            break

    start_index = 1 if address_col_index is not None else 0

    for row in rows[start_index:]:
        if not row:
            continue

        raw = str(row[address_col_index]).strip() if address_col_index is not None and address_col_index < len(row) else str(row[0]).strip()
        normalized = normalize_address(raw)

        if not normalized or normalized in {"адрес", "address"} or normalized in seen:
            continue

        seen.add(normalized)
        addresses.append(raw)

    return addresses


def merge_site_rows_with_table(site_rows, table_addresses):
    merged = list(site_rows)
    site_map = {}

    for row in site_rows:
        normalized = normalize_address(row.get("address", ""))

        if normalized:
            site_map[normalized] = row

    for address in table_addresses:
        normalized = normalize_address(address)

        if not normalized or normalized in site_map:
            continue

        merged.append(
            {
                "address": address,
                "municipality": "",
                "contract": "",
                "owner": "",
                "work_type": "",
                "stream_url": "",
                "last_check": "",
                "responsible": "",
                "camera_status": "not_connected",
                "stream_check_details": "Адрес отсутствует в таблице камер",
            }
        )

    return merged


async def find_first_visible(page, selectors, timeout=10000):
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            await locator.wait_for(state="visible", timeout=timeout)
            return locator
        except PlaywrightTimeoutError:
            continue
        except Exception:
            continue

    return None


async def safe_page_text(page):
    try:
        return await page.locator("body").inner_text()
    except Exception:
        try:
            return await page.evaluate("() => document.body ? (document.body.innerText || '') : ''")
        except Exception:
            return ""


async def safe_title(page):
    try:
        return await page.title()
    except Exception:
        return ""


def build_stream_debug_name(address, stream_url):
    raw = f"{address} | {stream_url}"
    digest = hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]
    base = sanitize_filename(address)[:50] or "camera"
    return f"{base}_{digest}"


def sanitize_filename(value):
    value = (value or "").strip()
    value = mask_sensitive_url(value)
    value = re.sub(r'[\\/:*?"<>|@]+', "_", value)
    value = re.sub(r"\s+", "_", value)
    return value


async def save_debug(page, name):
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        file_base = DEBUG_DIR / name

        await page.screenshot(path=str(file_base.with_suffix(".png")), full_page=True)

        html = await page.content()
        file_base.with_suffix(".html").write_text(html, encoding="utf-8")

    except Exception:
        pass


async def save_stream_debug(page, name):
    try:
        STREAM_DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        file_base = STREAM_DEBUG_DIR / sanitize_filename(name)

        await page.screenshot(path=str(file_base.with_suffix(".png")), full_page=True)

        html = await page.content()
        file_base.with_suffix(".html").write_text(html, encoding="utf-8")

    except Exception:
        pass


# =========================
# CLI entrypoint for Unified Dashboard
# =========================

CAMERAS_STATE_FILE = CAMERAS_DATA_DIR / "state" / "dashboard_state.json"


def save_dashboard_state(results):
    CAMERAS_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)

    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    counts = {
        "working": 0,
        "not_working": 0,
        "not_connected": 0,
        "unknown": 0,
    }

    for item in results:
        camera_status = item.get("camera_status") or "unknown"
        counts[camera_status] = counts.get(camera_status, 0) + 1

    payload = {
        "ok": True,
        "status": "success",
        "updated_at": checked_at,
        "checked_at": checked_at,
        "count": len(results),
        "summary": counts,
        "items": results,
        "results": results,
        "rows": results,
    }

    CAMERAS_STATE_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )

    print("STAGE: Сохранение результата", flush=True)
    print(f"Результат камер сохранён: {CAMERAS_STATE_FILE}", flush=True)
    print(f"Всего камер: {len(results)}", flush=True)


async def main():
    async def progress_callback(stage, message="", results=None):
        print(f"STAGE: {stage}", flush=True)

        if message:
            print(message, flush=True)

        if results is not None:
            save_dashboard_state(results)

    results = await run_camera_check(progress_callback)

    if results is not None:
        save_dashboard_state(results)


if __name__ == "__main__":
    asyncio.run(main())

