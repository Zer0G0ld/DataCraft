# -*- mode: python ; coding: utf-8 -*-

import os
import sys

# Obtém o caminho do diretório atual
current_dir = os.getcwd()
venv_site_packages = os.path.join(current_dir, 'venv', 'Lib', 'site-packages')

# Adiciona o site-packages ao path se existir
if os.path.exists(venv_site_packages):
    sys.path.insert(0, venv_site_packages)

a = Analysis(
    ['main.py'],  # <--- MUDAR para o nome do seu arquivo principal
    pathex=[],
    binaries=[],
    datas=[
        ('voto.ico', '.'),  # Inclui o ícone no executável
    ],
    hiddenimports=[
        # Bibliotecas principais
        'pandas',
        'openpyxl',
        'xml.etree.ElementTree',
        're',
        'datetime',
        'threading',
        'tkinter',
        
        # Dependências do pandas
        'numpy',
        'numpy._core',
        'numpy._core._multiarray_umath',
        'numpy.random',
        'numpy.linalg',
        'numpy.fft',
        'numpy.polynomial',
        'numpy.ma',
        'numpy.ctypeslib',
        'numpy._globals',
        'numpy._typing',
        
        # Dependências do openpyxl
        'et_xmlfile',
        '_openpyxl',
        
        # Dependências de data
        'dateutil',
        'dateutil.tz',
        'pytz',
        'tzdata',
        
        # Dependências do pandas (módulos internos)
        'pandas._libs',
        'pandas._libs.tslibs',
        'pandas._libs.window',
        'pandas._libs.algos',
        'pandas._libs.hashtable',
        'pandas._libs.indexing',
        'pandas._libs.internals',
        'pandas._libs.interval',
        'pandas._libs.join',
        'pandas._libs.json',
        'pandas._libs.lib',
        'pandas._libs.missing',
        'pandas._libs.ops',
        'pandas._libs.parsers',
        'pandas._libs.reshape',
        
        # Outras dependências
        'six',
        'sqlite3'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'PIL',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'virtualenv'
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
    name='DataCraft',  # <--- MUDAR para o nome do seu app
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,  # False para evitar warnings
    upx=False,    # False se não tiver UPX
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False = GUI mode
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='voto.ico'  # Seu ícone
)

# Opcional: Gerar uma pasta com vários arquivos (descomente se preferir)
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.datas,
#     strip=False,
#     upx=False,
#     name='DataCraft'
# )
