"""
PyInstaller hook for python-pptx package
This file ensures all pptx modules are properly included
"""
from PyInstaller.utils.hooks import collect_all, collect_submodules

# Collect all pptx modules
pptx_datas, pptx_binaries, pptx_hiddenimports = collect_all('pptx')

# Collect all submodules
pptx_submodules = collect_submodules('pptx')

# Additional pptx modules that might be missed
additional_pptx_modules = [
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
    'pptx.enum.dml',
    'pptx.enum.shapes',
    'pptx.enum.text',
]

# Combine all hidden imports
hiddenimports = list(set(pptx_hiddenimports + pptx_submodules + additional_pptx_modules))

# Add pptx to datas if not already present
if 'pptx' not in [item[0] for item in pptx_datas]:
    pptx_datas.append(('pptx', 'pptx'))

print(f"pptx hook: Found {len(hiddenimports)} hidden imports and {len(pptx_datas)} data files")
