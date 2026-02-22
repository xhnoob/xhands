# -*- mode: python ; coding: utf-8 -*-
"""
xHands PyInstaller 打包配置
"""

import os
import sys

block_cipher = None

current_dir = os.getcwd()

a = Analysis(
    ['xhands_main.py'],
    pathex=[current_dir],
    binaries=[],
    datas=[
        ('screenShot', 'screenShot'),
        ('flowScripts', 'flowScripts'),
        ('flowPY', 'flowPY'),
        ('config.json', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        'cv2',
        'numpy',
        'pyautogui',
        'pyperclip',
        'pytesseract',
        'PIL',
        'PIL.Image',
        'psutil',
        'win32gui',
        'win32process',
        'win32con',
        'win32api',
    ],
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='xhands',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
