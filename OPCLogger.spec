# -*- mode: python ; coding: utf-8 -*-
# DVMonOPCLogger.spec

import PyInstaller.config

block_cipher = None

a = Analysis(
    ['OPCLogger.py'],        
    pathex=['.'],                 
    binaries=[],
    datas=[],
    hiddenimports=['win32timezone'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

# Use upx if you want to compress. If you don't have UPX or don't want compression, set upx=False or remove it.
exe = EXE(
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OPCLogger',       
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,               
    icon='icon.ico',         
    version='versioninfo.txt'   
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='OPCLogger'      
)
