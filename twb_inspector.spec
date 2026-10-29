# -*- mode: python ; coding: utf-8 -*-
# Build the TWB Inspector Windows app (run from the project folder):
#     pip install pyinstaller
#     pyinstaller twb_inspector.spec --clean --noconfirm --workpath "%TEMP%\pyi_build"
# Result: dist\TWB_Inspector.exe

from PyInstaller.utils.hooks import collect_data_files

datas = [
    ('twb_inspector/vis-network.min.js', 'twb_inspector'),
    ('twb_inspector/assets/twb_inspector.ico', 'twb_inspector/assets'),
    ('twb_inspector/assets/twb_inspector.png', 'twb_inspector/assets'),
    ('twb_inspector/assets/twb_inspector_header.png', 'twb_inspector/assets'),
    ('twb_inspector/assets/favicon.png', 'twb_inspector/assets'),
]
datas += collect_data_files('customtkinter')   # themes and fonts of the window
# tkinterdnd2 (drag and drop) ships its own PyInstaller hook.

# Large packages the app never uses; excluded to keep the .exe small when they
# happen to be installed in the Python used for the build.
excludes = [
    'matplotlib', 'scipy', 'sklearn', 'IPython', 'jupyter', 'notebook', 'nbformat',
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'torch', 'tensorflow', 'pytest', 'sphinx',
    'sqlalchemy', 'pyarrow', 'numba', 'bokeh', 'plotly', 'dask',
]

a = Analysis(
    ['twb_inspector_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['PIL._tkinter_finder'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='TWB_Inspector',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='twb_inspector/assets/twb_inspector.ico',
)
