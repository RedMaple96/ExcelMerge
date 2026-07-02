#!/bin/bash
# ExcelMerge 命令行启动脚本
# 此脚本打包在 ExcelMerge.app/Contents/Resources/ 目录下
# 用于 Git difftool / SourceTree 配置
# 用法: excelmerge-diff <left-file> <right-file>

# 获取脚本所在目录（.app/Contents/Resources/）
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
# 二进制在同级的 MacOS 目录下
APP_MACOS_DIR="$(cd "$SCRIPT_DIR/../MacOS" && pwd)"

if [ -x "$APP_MACOS_DIR/ExcelMerge" ]; then
    # PyInstaller onedir 模式：必须在 MacOS 目录下执行才能找到 Frameworks
    cd "$APP_MACOS_DIR"
    exec "./ExcelMerge" "$@"
else
    echo "❌ 找不到 ExcelMerge 二进制文件" >&2
    echo "   请确认 ExcelMerge.app 完整性" >&2
    exit 1
fi
