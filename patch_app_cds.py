from pathlib import Path

app_file = Path("app.py")
content = app_file.read_text(encoding="utf-8")

# 1. Добавляем "/cds": "cds" в PATH_MODULE_MAP
if '"/cds": "cds"' not in content:
    content = content.replace(
        '"/appeals": "appeals",',
        '"/appeals": "appeals",\n    "/cds": "cds",',
    )
    print("[OK] Added /cds to PATH_MODULE_MAP")
else:
    print("[SKIP] /cds already in PATH_MODULE_MAP")

# 2. Блок CDS роутов
CDS_BLOCK = '''
# =========================
# CDS / ДИСПЕТЧЕРСКАЯ ЖКХ
# =========================

CDS_DATA_DIR = DATA_DIR / "cds"
CDS_RESULT_FILE = CDS_DATA_DIR / "result.json"

run_status["cds"] = {
    "running": False,
    "stage": "Ожидание запуска",
    "message": "Система готова к выгрузке из CDS.",
    "last_error": "",
}


@app.get("/cds", response_class=HTMLResponse)
async def cds_page(request: Request):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}

    return templates.TemplateResponse(
        request,
        "cds.html",
        {
            "request": request,
            "user_role": user.get("role", ""),
            "user_username": user.get("username", ""),
        },
    )


@app.post("/api/cds/export")
async def api_cds_export(payload: dict):
    if run_status["cds"]["running"]:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "Выгрузка уже выполняется."},
        )

    date_from = payload.get("date_from", "")
    date_to = payload.get("date_to", "")

    if not date_from or not date_to:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "Не указан период (date_from, date_to)."},
        )

    run_status["cds"]["running"] = True
    run_status["cds"]["stage"] = "Запуск"
    run_status["cds"]["message"] = "Начинаем выгрузку..."
    run_status["cds"]["last_error"] = ""

    try:
        from services.cds.auth import login_to_cds
        from services.cds.navigate import go_to_obrasheniya
        from services.cds.export import set_period_and_export

        pw, browser, context, page = await login_to_cds(headless=True)

        try:
            run_status["cds"]["stage"] = "Авторизация"
            run_status["cds"]["message"] = "Авторизация в 1С выполнена."
            await page.wait_for_timeout(3000)

            run_status["cds"]["stage"] = "Навигация"
            run_status["cds"]["message"] = "Переход в раздел Обращения..."

            nav_ok = await go_to_obrasheniya(page)
            if not nav_ok:
                run_status["cds"]["running"] = False
                run_status["cds"]["stage"] = "Ошибка"
                run_status["cds"]["message"] = "Не удалось перейти в раздел Обращения."
                run_status["cds"]["last_error"] = "Navigation failed"
                return {"success": False, "error": "Не удалось перейти в раздел Обращения."}

            await page.wait_for_timeout(2000)

            run_status["cds"]["stage"] = "Экспорт"
            msg = "Выгрузка за период " + date_from + " - " + date_to + "..."
            run_status["cds"]["message"] = msg

            export_ok = await set_period_and_export(page)

            if not export_ok:
                run_status["cds"]["running"] = False
                run_status["cds"]["stage"] = "Ошибка"
                run_status["cds"]["message"] = "Ошибка экспорта данных."
                run_status["cds"]["last_error"] = "Export failed"
                return {"success": False, "error": "Ошибка экспорта данных из 1С."}

            result_file = CDS_DATA_DIR / "result.json"
            result_data = load_json_file(result_file, default={})
            rows = result_data.get("rows", []) if isinstance(result_data, dict) else []

            run_status["cds"]["running"] = False
            run_status["cds"]["stage"] = "Готово"
            run_status["cds"]["message"] = "Выгружено " + str(len(rows)) + " обращений."

            def count_by_status(rows, statuses):
                return sum(1 for r in rows if (r.get("status") or r.get("Состояние", "")).lower() in statuses)

            return {
                "success": True,
                "total": len(rows),
                "count_new": count_by_status(rows, ("новое", "new", "открыто")),
                "count_in_work": count_by_status(rows, ("в работе", "in_work", "выполняется")),
                "count_done": count_by_status(rows, ("завершено", "done", "закрыто", "выполнено")),
                "rows": rows[:200],
                "download_url": "/data/cds/result.json" if rows else None,
            }

        finally:
            await browser.close()
            await pw.stop()

    except Exception as e:
        import traceback
        run_status["cds"]["running"] = False
        run_status["cds"]["stage"] = "Ошибка"
        run_status["cds"]["message"] = "Ошибка: " + str(e)
        run_status["cds"]["last_error"] = str(e)
        print("[CDS] Export error: " + traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e)},
        )


@app.get("/cds/run-status")
async def cds_run_status():
    return run_status["cds"]


'''

APPEALS_MARKER = "# =========================\n# APPEALS / ОБРАЩЕНИЯ\n# ========================="

if "async def cds_page" not in content:
    if APPEALS_MARKER in content:
        content = content.replace(APPEALS_MARKER, CDS_BLOCK + APPEALS_MARKER)
        print("[OK] Inserted CDS routes before APPEALS section")
    else:
        print("[ERROR] Could not find APPEALS marker in app.py!")
else:
    print("[SKIP] CDS routes already exist in app.py")

# 3. Создаём директорию data/cds
cds_dir = Path("data/cds")
cds_dir.mkdir(parents=True, exist_ok=True)
print("[OK] Directory data/cds ensured")

# 4. Записываем файл
app_file.write_text(content, encoding="utf-8")
print("[DONE] app.py patched successfully")
