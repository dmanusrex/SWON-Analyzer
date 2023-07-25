# -*- mode: python ; coding: utf-8 -*-


# Generated with:
# 
# pipenv run pyi-makespec --onefile --noconsole --add-data media\swon-analyzer.ico;media --icon media\swon-analyzer.ico --name swon-analyzer --splash media\swon-splash.png --version-file swon-analyzer.fileinfo club_analyze.py


block_cipher = None


a = Analysis(
    ['club_analyze.py'],
    pathex=[],
    binaries=[],
    datas=[('media\\swon-analyzer.ico', 'media')],
    hiddenimports=[],
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
splash = Splash(
    'media\\swon-splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(10,20),
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    splash,
    splash.binaries,
    [],
    name='swon-analyzer',
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
    icon=['media\\swon-analyzer.ico'],
    version='swon-analyzer.fileinfo',
)
