# -*- mode: python ; coding: utf-8 -*-
# DVMonOPCLogger.spec

import PyInstaller.config

block_cipher = None

a = Analysis(
    ['DVMonOPCLogger.py'],         # Your main Python script
    pathex=['.'],                  # Paths PyInstaller will search
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    name='DVMonOPCLogger',       # Base name of the output
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,               # True -> console application; False -> windowed app
    icon='star.ico',         # Replace with your actual .ico file
    version='versioninfo.txt'   # Embeds version info from the file below
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DVMonOPCLogger'       # The folder name in the 'dist' directory
)
