# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates',               'templates'),
        ('static',                  'static'),
        ('institution_lookup.xlsx', '.'),
    ],
    hiddenimports=[
        'flask',
        'jinja2',
        'jinja2.ext',
        'markupsafe',
        'werkzeug',
        'werkzeug.routing',
        'werkzeug.serving',
        'pandas',
        'openpyxl',
        'openpyxl.cell.rich_text',
        'openpyxl.cell.text',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.writer.excel',
        'et_xmlfile',
        'processing',
        'server',
        'webview',
        'webview.platforms.edgechromium',
        'webview.platforms.mshtml',
        'clr_loader',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'PIL', 'PyQt5'],
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
    exclude_binaries=True,
    name='Rate Sheet Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Rate Sheet Generator',
)
