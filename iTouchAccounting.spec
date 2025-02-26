# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

for folder in ['logs', 'exports']:
    if not os.path.exists(folder):
        os.makedirs(folder)

datas = [
    ('plan_codes.txt', '.'), 
    ('C:/Users/Administrator/anaconda3/Library/lib/tcl8.6', 'tcl8.6'),
    ('C:/Users/Administrator/anaconda3/Library/lib/tk8.6', 'tk8.6'),
    ('dependencies/icon.ico', '.')
]


binaries = [
    ('venv/Lib/site-packages/selenium', 'selenium'),
    ('dependencies/dlls/ffi.dll', '.'),
    ('dependencies/dlls/libbz2.dll', '.'),
    ('dependencies/dlls/libcrypto-3-x64.dll', '.'),
    ('dependencies/dlls/libexpat.dll', '.'),
    ('dependencies/dlls/liblzma.dll', '.'),
    ('dependencies/dlls/libssl-3-x64.dll', '.'),
    ('dependencies/dlls/_tkinter.pyd', '.'),
    ('dependencies/dlls/tk86t.dll', '.'),
    ('dependencies/dlls/tcl86t.dll', '.')
]

hiddenimports = [
    'tkinter',
    '_tkinter',
    'tkinter.ttk',
    'selenium',
    'selenium.webdriver',
    'selenium.webdriver.chrome.service',
    'selenium.webdriver.common.by',
    'selenium.webdriver.support',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',
    'bs4',
    'pandas',
    'numpy',
    'keyring',
    'xlsxwriter',
    'xml.parsers.expat',
    'pkg_resources.py2_warn',
    'pkg_resources'
]

# 收集相關套件
for pkg in ['selenium', 'pandas']:
    tmp_ret = collect_all(pkg)
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas.tests', 'pytest'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='iTouch會計帳目自動抓取程式',
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
    icon=r'D:\Project\Professor_Program\113-02_ItouchAccounting\dependencies\icon.ico'
)