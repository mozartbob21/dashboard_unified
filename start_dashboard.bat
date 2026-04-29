@echo off
chcp 65001 >nul
title Нейрона ИИ — Unified Dashboard

setlocal

REM Папка проекта определяется автоматически по месту нахождения этого bat-файла
set "PROJECT_DIR=%~dp0"
set "VENV_DIR=%PROJECT_DIR%.venv"
set "URL=http://127.0.0.1:8000/"

cd /d "%PROJECT_DIR%"

echo.
echo ==========================================
echo   Нейрона ИИ — запуск локальной платформы
echo ==========================================
echo.
echo Папка проекта:
echo %PROJECT_DIR%
echo.

REM Проверяем виртуальное окружение
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [ОШИБКА] Не найдено виртуальное окружение:
    echo %VENV_DIR%
    echo.
    echo Создай его командой:
    echo python -m venv .venv
    echo .venv\Scripts\activate
    echo pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM Активируем виртуальное окружение
call "%VENV_DIR%\Scripts\activate.bat"

echo Виртуальное окружение активировано.
echo.

REM Открываем браузер, когда сервер станет доступен
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "$url='%URL%'; for($i=0; $i -lt 40; $i++){ try { Invoke-WebRequest $url -UseBasicParsing -TimeoutSec 1 | Out-Null; Start-Process $url; exit } catch { Start-Sleep -Seconds 1 } }; Start-Process $url"

echo Запускаю сервер...
echo.
echo Адрес: %URL%
echo.
echo Чтобы остановить сервер — закрой это окно или нажми Ctrl+C.
echo.

REM ВАЖНО:
REM Если у тебя другой entrypoint, замени main:app на свой.
python -m uvicorn app:app --host 127.0.0.1 --port 8000

echo.
echo Сервер остановлен.
pause