#!/bin/bash

echo "🌟 Установка Комета - Накладные"
echo "================================"

# Проверяем Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 не установлен. Пожалуйста, установите Python 3.8+"
    exit 1
fi

echo "✅ Python найден: $(python3 --version)"

# Создаем виртуальное окружение
if [ ! -d "venv" ]; then
    echo ""
    echo "📦 Создаем виртуальное окружение..."
    python3 -m venv venv
    echo "✅ Виртуальное окружение создано"
else
    echo "✅ Виртуальное окружение уже существует"
fi

# Активируем окружение
echo ""
echo "🔌 Активируем окружение..."
source venv/bin/activate

# Обновляем pip
echo ""
echo "📦 Обновляем pip..."
pip install --upgrade pip

# Устанавливаем зависимости
echo ""
echo "📦 Устанавливаем зависимости..."
pip install -r requirements.txt

echo ""
echo "================================"
echo "✅ Установка завершена!"
echo "================================"
echo ""
echo "Для запуска программы выполните:"
echo "source venv/bin/activate"
echo "python main.py"
echo ""
