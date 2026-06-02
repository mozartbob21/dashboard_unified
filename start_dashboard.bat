das@echo off
chcp 65001 >nul
title Нейрона ИИ — Unified Dashboard

setlocal

REM Папка проекта определяется автоматически по месту нахождения этого bat-файла
set "PROJECT_DIR=%~dp0"
set "VENV_DIR=%PROJECT_DIR%.venv"
set "URL=http://127.0.0.1:8000/"

REM === Зеркала PyPI (обход без VPN) ===
set "MIRROR1=https://pypi.tuna.tsinghua.edu.cn/simple"
set "HOST1=pypi.tuna.tsinghua.edu.cn"
set "MIRROR2=https://mirrors.aliyun.com/pypi/simple/"
set "HOST2=mirrors.aliyun.com"

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

REM === Установка зависимостей через зеркала (обход без VPN) ===
if exist "%PROJECT_DIR%requirements.txt" (
    echo Устанавливаю зависимости из requirements.txt через зеркало Tsinghua...
    python -m pip install -r "%PROJECT_DIR%requirements.txt" ^
        -i %MIRROR1% --trusted-host %HOST1% ^
        --timeout 120 --retries 10 ^
        --disable-pip-version-check
    if errorlevel 1 (
        echo [ВНИМАНИЕ] Tsinghua не сработала. Пробую Aliyun...
        python -m pip install -r "%PROJECT_DIR%requirements.txt" ^
            -i %MIRROR2% --trusted-host %HOST2% ^
            --timeout 120 --retries 10 ^
            --disable-pip-version-check
        if errorlevel 1 (
            echo [ВНИМАНИЕ] Не удалось установить зависимости через оба зеркала.
            echo Продолжаю запуск с уже установленными пакетами...
        )
    )
    echo.
)

REM === Дополнительно: точно ставим bcrypt / passlib / jose / multipart ===
echo Проверяю дополнительные пакеты (bcrypt, passlib, python-jose, python-multipart)...
python -m pip install bcrypt passlib python-jose python-multipart ^
    -i %MIRROR1% --trusted-host %HOST1% ^
    --timeout 120 --retries 10 ^
    --disable-pip-version-check
if errorlevel 1 (
    echo [ВНИМАНИЕ] Tsinghua не сработала для доп. пакетов. Пробую Aliyun...
    python -m pip install bcrypt passlib python-jose python-multipart ^
        -i %MIRROR2% --trusted-host %HOST2% ^
        --timeout 120 --retries 10 ^
        --disable-pip-version-check
)
echo.

REM Открываем браузер, когда сервер станет доступен
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "$url='%URL%'; for($i=0; $i -lt 40; $i++){ try { Invoke-WebRequest $url -UseBasicParsing -TimeoutSec 1 | Out-Null; Start-Process $url; exit } catch { Start-Sleep -Seconds 1 } }; Start-Process $url"

echo Запускаю сервер...
echo.
echo Адрес: %URL%
echo.
echo Чтобы остановить сервер — закрой это окно или нажми Ctrl+C.
echo.

REM ВАЖНО: если у тебя другой entrypoint, замени app:app на свой
python -m uvicorn app:app --host 127.0.0.1 --port 8000 -- reload

echo.
echo Сервер остановлен.
pause