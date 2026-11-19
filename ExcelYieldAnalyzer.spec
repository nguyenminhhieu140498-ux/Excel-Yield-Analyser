from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

for package in ("PySide6", "pandas", "openpyxl", "matplotlib"):
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hiddenimports

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Excel Yield Analyzer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="Excel Yield Analyzer",
)
app = BUNDLE(
    coll,
    name="Excel Yield Analyzer.app",
    bundle_identifier="com.example.excelyieldanalyzer",
    info_plist={
        "CFBundleName": "Excel Yield Analyzer",
        "LSMinimumSystemVersion": "13.0",
    },
)