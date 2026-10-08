# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('resources', 'resources')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'IPython', 'jedi', 'parso', 'astroid', 'docutils', 'sphinx', 'babel', 'notebook', 'jupyter', 'jupyter_client', 'jupyter_core', 'tornado', 'zmq', 'nbconvert', 'nbformat', 'pytest', 'scipy', 'numexpr', 'tables', 'lib2to3', 'pygments', 'pandas.tests', 'numpy.tests', 'jinja2', 'markupsafe', 'matplotlib_inline', 'pyarrow', 'bottleneck', 'numba', 'llvmlite', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'PIL.ImageQt'],
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
    name='SensorMeasureTrace',
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
    icon=['C:/Users/ghk2si/AI_Tools/SensorCalFinder/resources/logo.ico'],
)
