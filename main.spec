# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# List all data files
datas = [
    ('frontend/index.html', 'frontend'),
    ('frontend/app.js', 'frontend'),
    ('frontend/style.css', 'frontend'),
]

# Hidden imports
hiddenimports = [
    'backend',
    'pandas',
    'openpyxl',
    'reportlab',
    'reportlab.lib',
    'reportlab.pdfbase._fontdata',
    'reportlab.pdfbase.ttfonts',
    'webview.platforms.winforms',
    'webview.platforms.win32',
    'webview.platforms.edgechromium',
    'numpy',
    'numpy._distributor_init',
    'numpy.core._multiarray_umath',
]

# Platform-specific DLLs (if needed)
binaries = []

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    a.datas,
    [],
    name='ListadoEscritos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ListadoEscritos',
)