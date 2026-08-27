# -*- mode: python ; coding: utf-8 -*-
"""ExcelMerge PyInstaller 打包配置。

关键点：
- ``src`` 作为数据目录一并打入，保证包内 import 路径解析正常。
- ``packaging/excelmerge-diff.sh`` 打入 Resources，供 Git difftool /
  SourceTree 配置使用（scripts/install-cli.sh 依赖该路径）。
- onedir 模式：启动快，便于 macOS 生成 .app / DMG。
"""

import os

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("src", "src"),
        ("resources", "resources"),
        ("packaging/excelmerge-diff.sh", "."),
    ],
    hiddenimports=["openpyxl"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ExcelMerge",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # GUI 应用，不显示终端
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="resources/icons/icon.icns" if os.path.exists("resources/icons/icon.icns") else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ExcelMerge",
)

# macOS 生成 .app 包
app = BUNDLE(
    coll,
    name="ExcelMerge.app",
    icon="resources/icons/icon.icns" if os.path.exists("resources/icons/icon.icns") else None,
    bundle_identifier="com.excelmerge.app",
    version="1.0.4",
)
