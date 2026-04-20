# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('openpyxl')

hiddenimports = [
    'src.config',
    'src.ingestion.fetcher',
    'src.processing.transformer',
    'src.export.exporter',
    'src.ui.app',
    # pandas internals needed at runtime
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.timestamps',
    'pandas._libs.tslibs.offsets',
    'pandas._libs.tslibs.period',
    'pandas._libs.sparse',
    'pandas._libs.indexing',
    'pandas._libs.join',
    'pandas._libs.lib',
    'pandas._libs.missing',
    'pandas._libs.ops',
    'pandas._libs.reduction',
    'pandas._libs.reshape',
    'pandas._libs.writers',
    'pandas.core.arrays.integer',
    'pandas.core.arrays.floating',
    'pandas.core.arrays.masked',
    'pandas.io.formats.excel',
    # openpyxl
    'openpyxl.cell._writer',
    # stdlib
    'queue',
    'threading',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'PIL', 'IPython', 'jupyter', 'notebook',
        'pytest', 'pandas.tests', 'openpyxl.tests',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='InputDevice_Dashboard_Refresh',
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
    icon=None,
)

