# -- mode python ; coding utf-8 --
from PyInstaller.utils.hooks import collect_all, copy_metadata

datas = [('app.py', '.'), ('Lector_Jsons.py', '.')]
binaries = []
hiddenimports = ['openpyxl']

# Recolectar todo el código estático (frontend) de Streamlit
st_datas, st_binaries, st_hiddenimports = collect_all('streamlit')
datas += st_datas
binaries += st_binaries
hiddenimports += st_hiddenimports

# ¡CLAVE PARA EVITAR EL ERROR 404! 
# Obligamos a Pyinstaller a empaquetar los metadatos de Streamlit
datas += copy_metadata('streamlit')

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
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
    name='ProcesadorDTE',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False para ocultar la ventana negra
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
