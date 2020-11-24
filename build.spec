# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['src\\main.py'],
             pathex=['src/'],
             binaries=[],
             datas=[('src/gui/assets/word_icon.png', 'assets'), ('src/gui/assets/excel_icon.png', 'assets'),
                    ('src/gui/assets/small_logo.ico', 'assets'), ('src/gui/assets/big_logo.ico', 'assets')],
             hiddenimports=['pkg_resources.py2_warn', 'pyodbc'],
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
          [],
          exclude_binaries=True,
          name='autoreport',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False)
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='autoreport')
