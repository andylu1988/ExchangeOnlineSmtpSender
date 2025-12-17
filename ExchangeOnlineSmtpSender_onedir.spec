# -*- mode: python ; coding: utf-8 -*-

# Fast-start onedir build (folder-based distribution)


a = Analysis(
    ['smtp_sender_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('CHANGELOG.md', '.'),
        ('README.md', '.'),
        ('README_CN.md', '.'),
        ('assets/smtp_tool.ico', 'assets'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ExchangeOnlineSmtpSender_v1.0.0_onedir',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    icon='assets/smtp_tool.ico',
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ExchangeOnlineSmtpSender_v1.0.0_onedir',
)
