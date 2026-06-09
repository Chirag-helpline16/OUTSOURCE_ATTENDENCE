# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_all, copy_metadata


datas = [
    (str(Path("app.py")), "."),
    (str(Path("src")), "src"),
    (str(Path(".streamlit") / "config.toml"), ".streamlit"),
]
binaries = []
hiddenimports = []


def runtime_submodule(module_name):
    parts = module_name.split(".")
    return not any(part in {"tests", "testing", "conftest"} or part.startswith("test_") for part in parts)


def collect_package(package_name):
    try:
        package_datas, package_binaries, package_hiddenimports = collect_all(
            package_name,
            filter_submodules=runtime_submodule,
        )
    except Exception:
        return
    datas.extend(package_datas)
    binaries.extend(package_binaries)
    hiddenimports.extend(package_hiddenimports)


for package in [
    "streamlit",
    "altair",
    "pydeck",
    "openpyxl",
    "pymongo",
    "dns",
]:
    collect_package(package)

for distribution in [
    "streamlit",
    "pandas",
    "openpyxl",
    "pymongo",
    "dnspython",
    "altair",
    "pydeck",
]:
    try:
        datas.extend(copy_metadata(distribution))
    except Exception:
        pass


a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pytest",
        "pandas.tests",
        "matplotlib",
        "IPython",
        "jupyter",
        "notebook",
        "scipy",
        "sklearn",
        "tensorflow",
        "torch",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="OutsourceAttendance",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="OutsourceAttendance",
)
