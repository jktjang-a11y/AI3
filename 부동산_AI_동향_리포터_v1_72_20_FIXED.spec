# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [
    ('report_generator_unpacked.py', '.'),
    ('api_collector_internal.py', '.'),
    ('land_price_monthly_pipeline.py', '.'),
    ('development_projects.json', '.'),
    ('housing_policy_events.json', '.'),
    ('전국 법정동 지역.xls', '.'),
    ('local_db/apt_move_in_schedule.sqlite3', 'local_db'),
    ('local_db/region_codes.sqlite3', 'local_db'),
]
binaries = []
hiddenimports = ['pandas', 'numpy', 'openpyxl', 'xlrd', 'requests', 'urllib3', 'bs4', 'jinja2', 'dateutil']
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('numpy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('bs4')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('jinja2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['realestate_api_report_unified_v1_72_20_rent_compare_raw_first_FIXED.py'],
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
    name='부동산 AI 동향 리포터',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='realestate_ai_report_icon.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
