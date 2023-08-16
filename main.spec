# -*- coding: utf-8 -*-

block_cipher = None

a = Analysis(['main.py'],
             pathex=['C:/Users/defuziq/PycharmProjects/pythonProject'],
             binaries=[],
             datas=[('static/*.png', 'static'),
             ('static/*.ui', 'static'),
             ('AuthWindow.py', '.'),
             ('MainWindow.py', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='main',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False, icon='C:/Users/defuziq/PycharmProjects/pythonProject/static/icon.ico')