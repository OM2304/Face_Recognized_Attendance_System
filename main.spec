# -*- mode: python ; coding: utf-8 -*-
import os
import face_recognition_models
import customtkinter

# 1. Locate Internal Libraries
face_models_data = os.path.join(os.path.dirname(face_recognition_models.__file__), 'models')
ctk_root = os.path.dirname(customtkinter.__file__)

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        (face_models_data, 'face_recognition_models/models'), # AI Neural Weights
        (ctk_root, 'customtkinter/')                         # UI Theme & Assets
    ],
    hiddenimports=['pandas', 'openpyxl', 'matplotlib', 'PIL.ImageTk'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 2. Configure Splash Screen
splash = Splash(
    'splash_img.png',  # Make sure this image exists in your folder!
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

# 3. Configure the Executable
exe = EXE(
    pyz,
    a.scripts,
    splash,               # Includes Splash logic
    splash.binaries,      # Includes Splash data
    a.binaries,
    a.datas,
    [],
    name='NEURAL_SCAN_FINAL',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # Set to False to hide the black terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)