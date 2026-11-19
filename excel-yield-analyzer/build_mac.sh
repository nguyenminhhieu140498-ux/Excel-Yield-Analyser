#!/usr/bin/env bash
set -euo pipefail

python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pyinstaller --noconfirm --clean ExcelYieldAnalyzer.spec
echo "App bundle created at dist/Excel Yield Analyzer.app"
