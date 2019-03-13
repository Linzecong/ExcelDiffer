# -*- mode: python -*-

block_cipher = None


a = Analysis(['MainWindow.py'],
             pathex=['.\\Algorithm.py', '.\\ChangeColorWidget.py', '.\\DiffWidget.py', '.\\ExcelWidget.py', '.\\MainWindow.py', '.\\ViewWidget.py', 'C:\\Users\\think\\Documents\\GitHub\\ExcelDiffer'],
             binaries=[],
             datas=[('./icon', 'icon'), ('style.qss', '.')],
             hiddenimports=['xlrd', 'sys', 'math', 'hashlib', 'datetime'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='MainWindow',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='icon\\opennew.ico')
