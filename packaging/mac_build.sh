#!/usr/bin/env bash
# macOS 打包脚本：生成 .app 并制作 DMG
# 用法: bash packaging/mac_build.sh
set -e
cd "$(dirname "$0")/.."   # 切到项目根目录

# 清理旧产物
rm -rf build dist
# 将 PyInstaller 缓存重定向到项目内，避免沙箱/权限问题
export PYINSTALLER_CONFIG_DIR="$PWD/.pyinstaller_cache"

echo "==> 开始 PyInstaller 打包..."
python3 -m PyInstaller \
    --noconfirm \
    --windowed \
    --onefile \
    --name "Excel Merge" \
    --add-data "src:src" \
    --hidden-import openpyxl \
    main.py

echo "✅ .app 构建完成: dist/Excel Merge.app (二进制: dist/Excel Merge)"

# 如已安装 create-dmg 则生成 DMG
if command -v create-dmg >/dev/null 2>&1; then
    echo "==> 生成 DMG 安装包..."
    create-dmg --volname "Excel Merge" --window-size 600 400 --icon-size 100 \
        --app-drop-link 425 200 "dist/Excel Merge.dmg" "dist/Excel Merge.app" 2>/dev/null || \
    create-dmg --volname "Excel Merge" "dist/Excel Merge.dmg" "dist/Excel Merge.app"
    echo "✅ DMG 构建完成: dist/Excel Merge.dmg"
else
    echo "⚠️  未安装 create-dmg，跳过 DMG 生成。可用 'brew install create-dmg' 安装后重试。"
fi
