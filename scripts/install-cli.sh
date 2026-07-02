#!/usr/bin/env bash
# 安装 ExcelMerge 命令行工具到 /usr/local/bin
# 用法: bash scripts/install-cli.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET="/usr/local/bin/excelmerge"

echo "==> 安装 ExcelMerge 命令行工具..."

if [ ! -d "/Applications/ExcelMerge.app" ] && [ ! -d "$SCRIPT_DIR/../dist/ExcelMerge.app" ]; then
    echo "❌ 未找到 ExcelMerge.app"
    echo "   请先运行 bash packaging/mac_build.sh 打包并复制到 /Applications"
    exit 1
fi

# 优先打包内置的脚本
if [ -d "/Applications/ExcelMerge.app" ]; then
    SOURCE_SCRIPT="/Applications/ExcelMerge.app/Contents/Resources/excelmerge-diff.sh"
elif [ -d "$SCRIPT_DIR/../dist/ExcelMerge.app" ]; then
    SOURCE_SCRIPT="$SCRIPT_DIR/../dist/ExcelMerge.app/Contents/Resources/excelmerge-diff.sh"
fi

sudo cp "$SOURCE_SCRIPT" "$TARGET"
sudo chmod +x "$TARGET"

echo "✅ 安装完成！"
echo "   命令行工具已安装到: $TARGET"
echo ""
echo "SourceTree 配置方法："
echo "  1. 打开 SourceTree → 偏好设置 → Diff"
echo "  2. External Diff: Custom"
echo "  3. Diff Command: $TARGET"
echo "  4. Arguments: \$LOCAL \$REMOTE"
echo ""
echo "或者（更简单，和 BeyondCompare 一样）："
echo "  Diff Command: /Applications/ExcelMerge.app/Contents/MacOS/ExcelMerge"
echo "  Arguments: \$LOCAL \$REMOTE"
echo ""
echo "Git 命令行配置（可选）："
echo "  git config --global difftool.excelmerge.cmd '$TARGET \"\$LOCAL\" \"\$REMOTE\"'"
