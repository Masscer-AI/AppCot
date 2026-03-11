#!/usr/bin/env bash
set -euo pipefail

if [[ ! -d ".venv" ]]; then
  python -m venv .venv
fi

source ".venv/Scripts/activate"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo ""
echo "Setup complete."
echo "Activate env: source .venv/Scripts/activate"
echo "Run API:      python main.py"
