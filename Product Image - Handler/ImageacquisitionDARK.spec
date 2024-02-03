# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['ImageacquisitionDark.py'],
             pathex=[],
             binaries=[],
             datas=[
    ("D:\\Users\\ForeignDude\\Desktop\\VBA 1 49\\Database refresh\\lib\\site-packages\\customtkinter\\assets", "customtkinter/assets"),
    ("D:\\Users\\ForeignDude\\Desktop\\VBA 1 49\\Database refresh\\Lib\\site-packages\\tkinterdnd2", "tkinterdnd2"),
    ("D:\\Users\\ForeignDude\\PycharmProjects\\Database refresh\\Work\\Imaging\\Capture.PNG", "."),
    ("D:\\Users\\ForeignDude\\PycharmProjects\\Database refresh\\Work\\Imaging\\discover_button.PNG", ".")
    ("D:\\Users\\ForeignDude\\PycharmProjects\\Database refresh\\Work\\Imaging\\dropforfront.png", ".")
    ("D:\\Users\\ForeignDude\\PycharmProjects\\Database refresh\\Work\\Imaging\\dropforside.png", ".")
    ("D:\\Users\\ForeignDude\\PycharmProjects\\Database refresh\\Work\\Imaging\\dropfortop.png", ".")
    ("D:\\Users\\ForeignDude\\PycharmProjects\\Database refresh\\Work\\Imaging\\instructions.png", ".")
],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='ImageacquisitionDark',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False)

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='ImageacquisitionDark')
