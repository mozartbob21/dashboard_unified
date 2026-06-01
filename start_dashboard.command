#!/bin/bash

cd "$(dirname "$0")"

echo ""
echo "=========================================="
echo "  Нейрона ИИ — запуск локальной платформы"
echo "=========================================="
echo ""
echo "Папка проекта:"
echo "$(pwd)"
echo ""

# --- Активация виртуального окружения ---
if [ -d ".venv" ]; then
    source .venv/bin/activate
    echo "Виртуальное окружение .venv активировано."
elif [ -d "venv" ]; then
    source venv/bin/activate
    echo "Виртуальное окружение venv активировано."
else
    echo "[ОШИБКА] Виртуальное окружение не найдено."
    echo ""
    echo "Создай его командой:"
    echo "  python3 -m venv .venv"
    echo "  source .venv/bin/activate"
    echo "  pip install -r requirements.txt"
    echo ""
    read -p "Нажмите Enter для выхода..."
    exit 1
fi
echo ""

# --- НОВОЕ: тихая проверка и установка зависимостей ---
if [ -f "requirements.txt" ]; then
    echo "Проверяю зависимости..."
    pip install -r requirements.txt --quiet --disable-pip-version-check 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "[ВНИМАНИЕ] Не удалось обновить зависимости (возможно, нет интернета)."
        echo "Продолжаю запуск с уже установленными пакетами..."
    else
        echo "Зависимости актуальны."
    fi
    echo ""
fi

# --- Авто-открытие браузера через 2 секунды ---
(sleep 2 && open "http://127.0.0.1:8000/") &

echo "Запускаю сервер..."
echo ""
echo "Адрес: http://127.0.0.1:8000/"
echo ""
echo "Чтобы остановить сервер — закрой это окно или нажми Ctrl+C."
echo ""

python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload

echo ""
echo "Работа завершена. Нажмите Enter для закрытия окна."
read