# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['meeting_reminder.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pystray._win32', 'PIL', 'PIL.Image', 'PIL.ImageDraw', 'winreg'],
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
    a.binaries,
    a.datas,
    [],
    name='會議提醒',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 不顯示黑色主控台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
