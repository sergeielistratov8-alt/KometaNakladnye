#!/bin/bash
# Запускает main.py из директории проекта.
# Автоматически использует виртуальное окружение, если оно находится в каталоге проекта.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Если явно задано — используем указанный интерпретатор
if [ -n "$PYTHON_EXEC" ]; then
	PYTHON_EXEC="$PYTHON_EXEC"
else
	# Попробуем найти venv внутри проекта в распространённых местах
	CANDIDATES=("$SCRIPT_DIR/venv/bin/python" "$SCRIPT_DIR/.venv/bin/python" "$SCRIPT_DIR/env/bin/python" "$SCRIPT_DIR/.env/bin/python")
	PYTHON_EXEC=""
	for p in "${CANDIDATES[@]}"; do
		if [ -x "$p" ]; then
			PYTHON_EXEC="$p"
			break
		fi
	done
	# Если не нашли в проекте, возьмём системный python3
	if [ -z "$PYTHON_EXEC" ]; then
		PYTHON_EXEC="/usr/bin/env python3"
	fi
fi

# Запускаем в текущей пользовательской сессии (GUI требует привязки к сессии)
# Если установлен Homebrew Qt/PyQt, добавим его site-packages в PYTHONPATH
BREW_PY_SITE="/opt/homebrew/lib/python3.11/site-packages"
if [ -d "$BREW_PY_SITE" ]; then
	export PYTHONPATH="$BREW_PY_SITE${PYTHONPATH+:$PYTHONPATH}"
fi

exec "$PYTHON_EXEC" main.py
