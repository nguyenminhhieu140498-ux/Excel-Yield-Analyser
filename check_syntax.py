import py_compile

py_compile.compile("engineering_chart_analyzer.py", doraise=True)
py_compile.compile("engineering_charts.py", doraise=True)
py_compile.compile("engine.py", doraise=True)
print("Python syntax OK")
