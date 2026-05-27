#!/bin/bash

cd "$(dirname "$0")"

echo "Запуск Unified Dashboard..."

if [ -d ".venv" ]; then
    source .venv/bin/activate
elif [ -d "venv" ]; then
    source venv/bin/activate
fi

python -m uvicorn --reload app:app --port 8000

echo ""
echo "Работа завершена. Нажмите Enter для закрытия окна."
read