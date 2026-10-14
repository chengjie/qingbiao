# Single executable; runtime libraries unpack into a temporary directory.
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(['launch.py'], pathex=['src'],
             datas=[('src/excel_compressor/assets', 'excel_compressor/assets')] + collect_data_files('PySide6', includes=['**/LICENSE*', '**/COPYING*']),
             hiddenimports=['PIL.MpoImagePlugin', 'PIL.JpegImagePlugin', 'PIL.PngImagePlugin'],
             excludes=['PySide6.QtNetwork', 'PySide6.QtQml', 'PySide6.QtQuick'],
             noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='QingBiao',
          debug=False, strip=False, upx=False, console=False, icon='src/excel_compressor/assets/logo.ico')
