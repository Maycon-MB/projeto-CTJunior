# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gerenciar_alunos.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('APIs/ct-junior-oficial.json', 'APIs/'), ('imgs/frente_carteirinha600.jpg', 'imgs/'), ('imgs/verso_carteirinha600.jpg', 'imgs/')],
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
    a.binaries,
    a.datas,
    [],
    name='GerenciarAlunos',
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
)
