# -*- mode: python ; coding: utf-8 -*-

# pyi-makespec --noconsole --add-data media\swon-analyzer.ico;media --icon media\swon-analyzer.ico --name swon-analyzer
#    --splash media\swon-splash.png --version-file swon-analyzer.fileinfo club_analyze.py


block_cipher = None

added_files = [ ('media\\swon-analyzer.ico', 'media'),
   ('CTkMessagebox\\icons', 'CTKMessagebox\\icons')]

a = Analysis(
    ['club_analyze.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
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
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    [],
    exclude_binaries=True,
    name='swon-analyzer',
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
    version='swon-analyzer.fileinfo',
    icon=['media\\swon-analyzer.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    splash.binaries,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='swon-analyzer',
)
