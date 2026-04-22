# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('ui/resources/styles.qss', 'ui/resources'),
        ('ui/resources/theme_midnight.qss', 'ui/resources'),
        ('ui/resources/theme_oceanic.qss', 'ui/resources'),
        ('ui/resources/theme_arctic.qss', 'ui/resources'),
        ('ui/resources/theme_ember.qss', 'ui/resources'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Trim provably-unused stdlib tools.
    # IMPORTANT: Do NOT add email, html, http, xmlrpc, or difflib here.
    # PyInstaller's pyi_rth_setuptools.py runtime hook imports pkg_resources,
    # which top-level imports email.parser — excluding email silently breaks
    # module resolution for lxml and python-docx at runtime in the EXE.
    excludes=[
        'tkinter', '_tkinter',
        'unittest',
        'pydoc', 'doctest',
        'curses',
    ],
    noarchive=False,
    # optimize=1 strips docstrings -> smaller .pyc -> faster bytecode load
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BluebookManager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    # Exclude Qt DLLs from UPX: decompressing large Qt binaries at every
    # launch costs more startup time than the size saving is worth.
    upx_exclude=[
        'Qt6Core.dll', 'Qt6Gui.dll', 'Qt6Widgets.dll',
        'Qt6Network.dll', 'Qt6Svg.dll', 'Qt6OpenGL.dll',
        'Qt6DBus.dll', 'Qt6PrintSupport.dll',
        'qwindows.dll', 'qwindowsvistastyle.dll',
        'python3*.dll', 'vcruntime*.dll', 'msvcp*.dll',
    ],
    name='BluebookManager',
)
