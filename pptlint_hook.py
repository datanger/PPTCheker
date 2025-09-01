"""
PyInstaller hook for pptlint package
This file helps PyInstaller correctly identify all dependencies
"""
from PyInstaller.utils.hooks import collect_all

# Collect all submodules and data files
datas, binaries, hiddenimports = collect_all('pptlint')

# Add additional hidden imports that might be missed
additional_hiddenimports = [
    'pptx',
    'pptx.util',
    'pptx.enum',
    'pptx.dml',
    'pptx.oxml',
    'pptx.oxml.ns',
    'pptx.oxml.xmlchemy',
    'pptx.oxml.parts',
    'pptx.oxml.shapes',
    'pptx.oxml.slide',
    'pptx.oxml.presentation',
    'pptx.oxml.theme',
    'pptx.oxml.styles',
    'pptx.oxml.table',
    'pptx.oxml.chart',
    'pptx.oxml.drawing',
    'pptx.oxml.picture',
    'pptx.oxml.media',
    'pptx.oxml.notes',
    'pptx.oxml.handout',
    'pptx.oxml.comments',
    'pptx.oxml.relationships',
    'pptx.oxml.shared',
    'pptx.oxml.simpletypes',
    'pptx.oxml.text',
    'pptx.oxml.vml',
    'pptx.oxml.worksheet',
    'pptx.oxml.workbook',
    'PIL',
    'PIL.Image',
    'PIL.ImageDraw',
    'PIL.ImageFont',

    'regex',
    'jinja2',
    'streamlit',
    'yaml',
    'json',
    'urllib.request',
    'urllib.error',
    'threading',
    'subprocess',
    'platform',
    'pathlib',
    'datetime',
    'collections',
    're',
    'tempfile',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
]

hiddenimports.extend(additional_hiddenimports)
