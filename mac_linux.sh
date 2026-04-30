#!/usr/bin/env bash

set -euo pipefail

VENV_DIR="venv"
PYTHON_BIN="python3"

if [[ ! -x "$VENV_DIR/bin/python" ]]; then
	"$PYTHON_BIN" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

if ! python -c "import xlsxwriter" >/dev/null 2>&1; then
	python -m pip install --upgrade pip >/dev/null
	python -m pip install -r requirements.txt
fi

python .dbcConverter.py

deactivate