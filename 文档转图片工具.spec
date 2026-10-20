# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['win32com.client', 'win32com.client.gencache', 'pythoncom', 'fitz', 'docx', 'PIL', 'PIL.Image', 'PIL.ImageTk', 'app', 'app.overview', 'app.overview.document_to_overview', 'app.images', 'app.images.document_to_images', 'app.pdf_render', 'app.logger_config', 'app.ppt', 'app.ppt.ppt_to_overview']
hiddenimports += collect_submodules('win32com')
hiddenimports += collect_submodules('win32comext')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('PIL')


a = Analysis(
    ['ppt_to_overview_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
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
    name='文档转图片工具',
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
