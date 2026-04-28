# DataCraft.spec
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
    ['DataCraft.py'],  # Seu arquivo principal (pode ser DataCraft.py ou main.py)
    pathex=[],
    binaries=[],
    datas=[
        ('voto.ico', '.'),  # Ícone
        ('README.md', '.'),  # Incluir README
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
        'json',
        'csv',
        'yaml',
        'sqlalchemy',
        'markdown',
        'jinja2',
        
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
        
        # Módulos do DataCraft
        'core',
        'core.engine',
        'core.reader_factory',
        'core.writer_factory',
        'readers',
        'readers.base_reader',
        'readers.xml_reader',
        'readers.json_reader',
        'readers.csv_reader',
        'readers.yaml_reader',
        'readers.excel_reader',
        'readers.sql_reader',
        'readers.parquet_reader',
        'writers',
        'writers.base_writer',
        'writers.excel_writer',
        'writers.json_writer',
        'writers.csv_writer',
        'writers.html_writer',
        'writers.markdown_writer',
        'writers.sql_writer',
        'writers.parquet_writer',
        'transformers',
        'transformers.flatten',
        'transformers.mapper',
        'transformers.aggregator',
        'gui',
        'gui.main_window',
        'gui.mapper_dialog',
        'gui.batch_processor',
        
        # Outras dependências
        'six',
        'sqlite3',
        'pathlib',
        'typing',
        'abc'
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
        'virtualenv',
        'tkinter.test',
        'unittest'
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
    name='DataCraft',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Mude para True se tiver UPX instalado
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False = sem console (recomendado)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='voto.ico'
)

# Se quiser gerar uma pasta com vários arquivos (mais rápido para iniciar)
# Descomente abaixo:
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.datas,
#     strip=False,
#     upx=False,
#     name='DataCraft'
# )