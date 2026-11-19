from __future__ import annotations

import py_compile

FILES = (
    "engine.py",
    "engineering_charts.py",
    "engineering_chart_analyzer.py",
)

for file_name in FILES:
    py_compile.compile(file_name, doraise=True)

print("Python syntax OK")
