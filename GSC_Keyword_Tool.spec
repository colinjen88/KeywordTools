# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('allKeyWord_normalized.csv', '.'), ('gsc_keyword_report_sample.csv', '.'), ('gsc_keyword_report.py', '.')]
binaries = []
hiddenimports = ['ttkbootstrap', 'tkinter', 'tkinter.ttk', 'pandas', 'openpyxl', 'gsc_keyword_report', 'google.oauth2', 'googleapiclient', 'google_auth_oauthlib', 'google.auth', 'googleapiclient.discovery', 'googleapiclient.errors', 'google.oauth2.service_account', 'google.auth.transport.requests']
hiddenimports += collect_submodules('googleapiclient')
hiddenimports += collect_submodules('google.oauth2')
hiddenimports += collect_submodules('google.auth')
tmp_ret = collect_all('googleapiclient')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('google.oauth2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('google.auth')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['run_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='GSC_Keyword_Tool',
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
