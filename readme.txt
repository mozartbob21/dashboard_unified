UNIFIED DASHBOARD
ПОЛНАЯ ИНСТРУКЦИЯ ПО УСТАНОВКЕ И ЗАПУСКУ

============================================================
1. ОПИСАНИЕ
============================================================

Unified Dashboard — локальная веб-панель на Python/FastAPI для работы с модулями:

- обращения граждан;
- камеры;
- ЕДО;
- просроченные задачи;
- УТНКР;
- водоконтроль;
- муниципальные отчёты;
- генерация документов;
- работа с DOCX, PDF, XLSX;
- локальная обработка данных через веб-интерфейс.

После запуска проект открывается в браузере по адресу:

http://127.0.0.1:8000


============================================================
2. ЧТО ДОЛЖНО БЫТЬ В ПРОЕКТЕ
============================================================

В корне проекта должны быть примерно такие файлы и папки:

unified_dashboard/
    app.py
    requirements.txt
    .env
    start_dashboard.bat
    start_dashboard.command
    data/
    generated/
    services/
    static/
    templates/
    templates_prescriptions/

Главный файл запуска:

app.py


============================================================
3. УСТАНОВКА НА MACOS
============================================================

------------------------------------------------------------
3.1. Открыть Terminal
------------------------------------------------------------

Откройте приложение Terminal.

Перейдите в папку проекта.

Например, если проект лежит на рабочем столе:

cd ~/Desktop/unified_dashboard

Если проект лежит в другом месте, укажите свой путь.


------------------------------------------------------------
3.2. Установить Homebrew
------------------------------------------------------------

Если Homebrew уже установлен, этот шаг можно пропустить.

/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

Проверить Homebrew:

brew --version


------------------------------------------------------------
3.3. Установить Python
------------------------------------------------------------

brew install python

Проверить Python:

python3 --version

Проверить pip:

python3 -m pip --version


------------------------------------------------------------
3.4. Установить Git
------------------------------------------------------------

brew install git

Проверить Git:

git --version


------------------------------------------------------------
3.5. Установить FFmpeg
------------------------------------------------------------

brew install ffmpeg

Проверить FFmpeg:

ffmpeg -version


------------------------------------------------------------
3.6. Создать виртуальное окружение
------------------------------------------------------------

Выполнять из корня проекта:

python3 -m venv .venv

Активировать окружение:

source .venv/bin/activate

Если всё хорошо, в начале строки терминала появится:

(.venv)


------------------------------------------------------------
3.7. Обновить pip
------------------------------------------------------------

python3 -m pip install --upgrade pip


------------------------------------------------------------
3.8. Установить зависимости из requirements.txt
------------------------------------------------------------

pip install -r requirements.txt


------------------------------------------------------------
3.9. Дополнительно установить обязательные пакеты
------------------------------------------------------------

Выполнить команды по очереди:

pip install python-multipart

pip install python-docx pypdf

python3 -m pip install reportlab openpyxl

Также можно установить основные зависимости одной командой:

pip install fastapi uvicorn jinja2 python-multipart python-docx pypdf reportlab openpyxl pandas xlrd requests beautifulsoup4 lxml python-dotenv aiofiles httpx pydantic itsdangerous email-validator playwright


------------------------------------------------------------
3.10. Установить браузеры Playwright
------------------------------------------------------------

playwright install

Если команда playwright не найдена:

python3 -m playwright install


------------------------------------------------------------
3.11. Создать рабочие папки
------------------------------------------------------------

mkdir -p data data/appeals data/edo data/edo/screenshots data/overdue data/cameras data/utnkr data/watercontrol screenshots generated


------------------------------------------------------------
3.12. Создать базовые JSON-файлы
------------------------------------------------------------

Скопируйте и выполните команду целиком:

python3 <<'PY'
from pathlib import Path
import json

files = [
    "data/storage_state.json",
    "data/municipality_registry.json",
    "data/edo/demo_data.json",
    "data/edo/last_result.json",
    "data/edo/latest_result.json",
    "data/edo/responsibles_by_municipality.json",
    "data/edo/result.json",
    "data/appeals/appeals.json",
]

for file in files:
    path = Path(file)
    path.parent.mkdir(parents=True, exist_ok=True)

    if not path.exists():
        if file.endswith("appeals.json"):
            data = []
        else:
            data = {}

        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        print(f"Создан файл: {path}")
    else:
        print(f"Уже существует: {path}")

print("Готово.")
PY


------------------------------------------------------------
3.13. Запустить проект на macOS
------------------------------------------------------------

python3 -m uvicorn app:app --reload --port 8000

Если команда python3 не подходит, можно попробовать:

python -m uvicorn app:app --reload --port 8000

После запуска открыть в браузере:

http://127.0.0.1:8000


------------------------------------------------------------
3.14. Если порт 8000 занят
------------------------------------------------------------

Остановить старый uvicorn:

pkill -f "uvicorn" || true

Запустить заново:

python3 -m uvicorn app:app --reload --port 8000

Или запустить на другом порту:

python3 -m uvicorn app:app --reload --port 8001

Тогда открыть:

http://127.0.0.1:8001


============================================================
4. КНОПКА ЗАПУСКА НА MACOS
============================================================

Создайте файл в корне проекта:

start_dashboard.command

Вставьте в него текст:

#!/bin/bash

cd "$(dirname "$0")"

echo "Starting Unified Dashboard..."

if [ ! -f ".venv/bin/activate" ]; then
    echo "Virtual environment not found."
    echo "Creating virtual environment..."
    python3 -m venv .venv
fi

source .venv/bin/activate

echo "Installing requirements..."
pip install -r requirements.txt
pip install python-multipart
pip install python-docx pypdf
python3 -m pip install reportlab openpyxl

echo "Installing Playwright browsers..."
python3 -m playwright install

echo "Creating folders..."
mkdir -p data data/appeals data/edo data/edo/screenshots data/overdue data/cameras data/utnkr data/watercontrol screenshots generated

echo "Starting server..."
python3 -m uvicorn app:app --reload --port 8000

read -p "Press Enter to exit..."

После создания файла дать права на запуск:

chmod +x start_dashboard.command

Теперь проект можно запускать двойным кликом по файлу:

start_dashboard.command


============================================================
5. УСТАНОВКА НА WINDOWS
============================================================

------------------------------------------------------------
5.1. Установить Python
------------------------------------------------------------

Скачать Python:

https://www.python.org/downloads/

ВАЖНО:
При установке обязательно поставить галочку:

Add Python to PATH

Проверить в PowerShell:

python --version

Проверить pip:

pip --version


------------------------------------------------------------
5.2. Установить Git
------------------------------------------------------------

Вариант 1 — скачать с сайта:

https://git-scm.com/download/win

Вариант 2 — через winget:

winget install -e --id Git.Git

Проверить:

git --version


------------------------------------------------------------
5.3. Установить FFmpeg
------------------------------------------------------------

Через winget:

winget install -e --id Gyan.FFmpeg

Проверить:

ffmpeg -version

Если команда не находится, перезапустите PowerShell или компьютер.


------------------------------------------------------------
5.4. Перейти в папку проекта
------------------------------------------------------------

Например, если проект на рабочем столе:

cd Desktop\unified_dashboard

Если проект в другом месте, укажите свой путь.


------------------------------------------------------------
5.5. Разрешить запуск скриптов PowerShell
------------------------------------------------------------

Если Windows блокирует активацию виртуального окружения:

Set-ExecutionPolicy Bypass -Scope Process

Команда действует только в текущем окне PowerShell.


------------------------------------------------------------
5.6. Создать виртуальное окружение
------------------------------------------------------------

python -m venv .venv

Активировать виртуальное окружение:

.venv\Scripts\activate

Если всё хорошо, в начале строки появится:

(.venv)


------------------------------------------------------------
5.7. Обновить pip
------------------------------------------------------------

python -m pip install --upgrade pip


------------------------------------------------------------
5.8. Установить зависимости из requirements.txt
------------------------------------------------------------

pip install -r requirements.txt


------------------------------------------------------------
5.9. Дополнительно установить обязательные пакеты
------------------------------------------------------------

Выполнить команды по очереди:

pip install python-multipart

pip install python-docx pypdf

python -m pip install reportlab openpyxl

Также можно установить основные зависимости одной командой:

pip install fastapi uvicorn jinja2 python-multipart python-docx pypdf reportlab openpyxl pandas xlrd requests beautifulsoup4 lxml python-dotenv aiofiles httpx pydantic itsdangerous email-validator playwright


------------------------------------------------------------
5.10. Установить браузеры Playwright
------------------------------------------------------------

playwright install

Если команда playwright не найдена:

python -m playwright install


------------------------------------------------------------
5.11. Создать рабочие папки
------------------------------------------------------------

New-Item -ItemType Directory -Force -Path .\data, .\data\appeals, .\data\edo, .\data\edo\screenshots, .\data\overdue, .\data\cameras, .\data\utnkr, .\data\watercontrol, .\screenshots, .\generated | Out-Null


------------------------------------------------------------
5.12. Создать базовые JSON-файлы
------------------------------------------------------------

В PowerShell иногда неудобно выполнять многострочный Python через консоль.
Поэтому проще создать файл:

init_files.py

Вставить в него:

from pathlib import Path
import json

files = [
    "data/storage_state.json",
    "data/municipality_registry.json",
    "data/edo/demo_data.json",
    "data/edo/last_result.json",
    "data/edo/latest_result.json",
    "data/edo/responsibles_by_municipality.json",
    "data/edo/result.json",
    "data/appeals/appeals.json",
]

for file in files:
    path = Path(file)
    path.parent.mkdir(parents=True, exist_ok=True)

    if not path.exists():
        if file.endswith("appeals.json"):
            data = []
        else:
            data = {}

        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        print(f"Создан файл: {path}")
    else:
        print(f"Уже существует: {path}")

print("Готово.")

Потом выполнить:

python init_files.py


------------------------------------------------------------
5.13. Запустить проект на Windows
------------------------------------------------------------

python -m uvicorn app:app --reload --port 8000

После запуска открыть в браузере:

http://127.0.0.1:8000


------------------------------------------------------------
5.14. Если порт 8000 занят
------------------------------------------------------------

Посмотреть процессы Python:

tasklist | findstr python

Закрыть процессы Python:

taskkill /F /IM python.exe

Запустить проект снова:

python -m uvicorn app:app --reload --port 8000

Или запустить на другом порту:

python -m uvicorn app:app --reload --port 8001

Тогда открыть:

http://127.0.0.1:8001


============================================================
6. КНОПКА ЗАПУСКА НА WINDOWS
============================================================

Создайте файл в корне проекта:

start_dashboard.bat

Вставьте в него:

@echo off
cd /d "%~dp0"

echo Starting Unified Dashboard...

if not exist ".venv\Scripts\activate.bat" (
    echo Virtual environment not found.
    echo Creating virtual environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

echo Installing requirements...
pip install -r requirements.txt
pip install python-multipart
pip install python-docx pypdf
python -m pip install reportlab openpyxl

echo Installing Playwright browsers...
python -m playwright install

echo Creating folders...
if not exist data mkdir data
if not exist data\appeals mkdir data\appeals
if not exist data\edo mkdir data\edo
if not exist data\edo\screenshots mkdir data\edo\screenshots
if not exist data\overdue mkdir data\overdue
if not exist data\cameras mkdir data\cameras
if not exist data\utnkr mkdir data\utnkr
if not exist data\watercontrol mkdir data\watercontrol
if not exist screenshots mkdir screenshots
if not exist generated mkdir generated

echo Starting server...
python -m uvicorn app:app --reload --port 8000

pause

После этого проект можно запускать двойным кликом по файлу:

start_dashboard.bat


============================================================
7. ФАЙЛ .ENV
============================================================

В корне проекта может быть файл:

.env

Если файла нет, создайте его вручную.

Минимальный пример:

APP_ENV=local
APP_DEBUG=true

Если проект использует логины, пароли, токены или служебные адреса,
они также могут храниться в .env.

ВАЖНО:
Файл .env нельзя публиковать в открытый доступ.


============================================================
8. REQUIREMENTS.TXT
============================================================

Если нужно создать или обновить requirements.txt:

pip freeze > requirements.txt

Рекомендуемый минимальный набор зависимостей:

fastapi
uvicorn
jinja2
python-multipart
python-docx
pypdf
reportlab
openpyxl
pandas
xlrd
requests
beautifulsoup4
lxml
python-dotenv
aiofiles
httpx
pydantic
itsdangerous
email-validator
playwright


============================================================
9. БЫСТРЫЙ СТАРТ НА MACOS
============================================================

Если Python, Homebrew и Git уже установлены:

cd ~/Desktop/unified_dashboard
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
pip install -r requirements.txt
pip install python-multipart
pip install python-docx pypdf
python3 -m pip install reportlab openpyxl
python3 -m playwright install
mkdir -p data data/appeals data/edo data/edo/screenshots data/overdue data/cameras data/utnkr data/watercontrol screenshots generated
python3 -m uvicorn app:app --reload --port 8000

Открыть:

http://127.0.0.1:8000


============================================================
10. БЫСТРЫЙ СТАРТ НА WINDOWS
============================================================

Если Python и Git уже установлены:

cd Desktop\unified_dashboard
python -m venv .venv
Set-ExecutionPolicy Bypass -Scope Process
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install python-multipart
pip install python-docx pypdf
python -m pip install reportlab openpyxl
python -m playwright install
New-Item -ItemType Directory -Force -Path .\data, .\data\appeals, .\data\edo, .\data\edo\screenshots, .\data\overdue, .\data\cameras, .\data\utnkr, .\data\watercontrol, .\screenshots, .\generated | Out-Null
python -m uvicorn app:app --reload --port 8000

Открыть:

http://127.0.0.1:8000


============================================================
11. ПРОВЕРКА, ЧТО СЕРВЕР РАБОТАЕТ
============================================================

После запуска в терминале должно быть сообщение примерно такого вида:

Uvicorn running on http://127.0.0.1:8000

После этого открыть браузер:

http://127.0.0.1:8000


============================================================
12. КАК ОСТАНОВИТЬ СЕРВЕР
============================================================

В терминале, где запущен сервер, нажмите:

Ctrl + C


============================================================
13. ЧАСТЫЕ ОШИБКИ
============================================================

------------------------------------------------------------
Ошибка: ModuleNotFoundError: No module named 'fastapi'
------------------------------------------------------------

Решение:

pip install fastapi

Или:

pip install -r requirements.txt


------------------------------------------------------------
Ошибка: No module named multipart
------------------------------------------------------------

Решение:

pip install python-multipart


------------------------------------------------------------
Ошибка: No module named docx
------------------------------------------------------------

Решение:

pip install python-docx


------------------------------------------------------------
Ошибка: No module named pypdf
------------------------------------------------------------

Решение:

pip install pypdf


------------------------------------------------------------
Ошибка: No module named reportlab
------------------------------------------------------------

Решение:

pip install reportlab


------------------------------------------------------------
Ошибка: No module named openpyxl
------------------------------------------------------------

Решение:

pip install openpyxl


------------------------------------------------------------
Ошибка: uvicorn не найден
------------------------------------------------------------

Решение:

pip install uvicorn

Запуск:

python -m uvicorn app:app --reload --port 8000

На macOS:

python3 -m uvicorn app:app --reload --port 8000


------------------------------------------------------------
Ошибка: порт уже занят
------------------------------------------------------------

На macOS:

pkill -f "uvicorn" || true

Потом:

python3 -m uvicorn app:app --reload --port 8000

На Windows:

taskkill /F /IM python.exe

Потом:

python -m uvicorn app:app --reload --port 8000


============================================================
14. ОБНОВЛЕНИЕ ПРОЕКТА
============================================================

Если проект хранится в Git:

git pull

После обновления рекомендуется установить зависимости:

pip install -r requirements.txt

Дополнительно:

pip install python-multipart
pip install python-docx pypdf
python -m pip install reportlab openpyxl

На macOS можно так:

python3 -m pip install reportlab openpyxl

Запустить сервер:

python -m uvicorn app:app --reload --port 8000

На macOS:

python3 -m uvicorn app:app --reload --port 8000


============================================================
15. ПОЛНАЯ УСТАНОВКА ОДНИМ БЛОКОМ ДЛЯ MACOS
============================================================

Выполнить из корня проекта:

python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
pip install -r requirements.txt
pip install python-multipart
pip install python-docx pypdf
python3 -m pip install reportlab openpyxl
python3 -m playwright install
mkdir -p data data/appeals data/edo data/edo/screenshots data/overdue data/cameras data/utnkr data/watercontrol screenshots generated
python3 -m uvicorn app:app --reload --port 8000


============================================================
16. ПОЛНАЯ УСТАНОВКА ОДНИМ БЛОКОМ ДЛЯ WINDOWS POWERSHELL
============================================================

Выполнить из корня проекта:

python -m venv .venv
Set-ExecutionPolicy Bypass -Scope Process
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install python-multipart
pip install python-docx pypdf
python -m pip install reportlab openpyxl
python -m playwright install
New-Item -ItemType Directory -Force -Path .\data, .\data\appeals, .\data\edo, .\data\edo\screenshots, .\data\overdue, .\data\cameras, .\data\utnkr, .\data\watercontrol, .\screenshots, .\generated | Out-Null
python -m uvicorn app:app --reload --port 8000


============================================================
17. ГЛАВНОЕ НА КАЖДЫЙ ЗАПУСК
============================================================

MACOS:

cd ~/Desktop/unified_dashboard
source .venv/bin/activate
python3 -m uvicorn app:app --reload --port 8000

WINDOWS:

cd Desktop\unified_dashboard
.venv\Scripts\activate
python -m uvicorn app:app --reload --port 8000

Открыть:

http://127.0.0.1:8000


============================================================
КОНЕЦ ИНСТРУКЦИИ
============================================================