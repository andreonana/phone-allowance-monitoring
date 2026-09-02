# -*- mode: python ; coding: utf-8 -*-
#
# Empaquette l'application en un seul .exe Windows sans console (double-clic
# direct, aucun terminal requis). Compilé automatiquement par
# .github/workflows/build-windows.yml sur un runner Windows (PyInstaller ne
# fait pas de cross-compilation : ce .spec doit être exécuté sous Windows).
#
# Build local (sous Windows, avec le venv du projet actif) :
#     pyinstaller PhoneAllowanceMonitoring.spec
# -> dist/PhoneAllowanceMonitoring.exe

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=["matplotlib.backends.backend_tkagg"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PhoneAllowanceMonitoring",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,        # --noconsole : pas de fenêtre terminal noire au lancement
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
