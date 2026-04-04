# -*- mode: python ; coding: utf-8 -*-

import os


spec_dir = globals().get('SPECPATH')
project_root = os.path.abspath(os.path.join(spec_dir, os.pardir)) if spec_dir else os.path.abspath(os.getcwd())

a = Analysis(
    [os.path.join(project_root, 'main.py')],
    pathex=[project_root],
    binaries=[],
    datas=[
        (os.path.join(project_root, 'VERSION'), '.'),
        (os.path.join(project_root, 'index.html'), '.'),
        (os.path.join(project_root, 'styles.css'), '.'),
        (os.path.join(project_root, 'script.js'), '.'),
        (os.path.join(project_root, '.env'), '.'),
        (os.path.join(project_root, 'assets', 'acies.png'), '.'),
        (os.path.join(project_root, 'scripts', 'merge_pdfs.py'), 'scripts'),
        (os.path.join(project_root, 'scripts', 'detect_pdf_size.py'), 'scripts'),
        (os.path.join(project_root, 'scripts', 'PlotDWGs.ps1'), 'scripts'),
        (os.path.join(project_root, 'scripts', 'FreezeLayersDWGs.ps1'), 'scripts'),
        (os.path.join(project_root, 'scripts', 'ThawLayersDWGs.ps1'), 'scripts'),
        (os.path.join(project_root, 'scripts', 'removeXREFPaths.ps1'), 'scripts'),
        (os.path.join(project_root, 'scripts', 'StripRefPaths.dll'), 'scripts'),
        (os.path.join(project_root, 'templates'), 'templates'),
        (os.path.join(project_root, 'WireSizerApplication', 'dist'), 'WireSizerApplication\\dist')
    ],
    hiddenimports=['pillow_heif', '_pillow_heif'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ACIES Scheduler',
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
    icon=[os.path.join(project_root, 'assets', 'acies.ico')],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ACIES Scheduler',
)
