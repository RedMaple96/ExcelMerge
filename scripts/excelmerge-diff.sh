#!/bin/bash
# ExcelMerge 命令行启动脚本（开发备用）
# 用于 Git difftool / SourceTree 配置
# 用法: excelmerge-diff <left-file> <right-file>

# 查找应用程序位置：优先项目dist目录，其次/Applications
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_MACOS_DIR=""

# 优先使用项目目录（开发/测试用）
if [ -d "$SCRIPT_DIR/../dist/ExcelMerge.app/Contents/MacOS" ]; then
    APP_MACOS_DIR="$(cd "$SCRIPT_DIR/../dist/ExcelMerge.app/Contents/MacOS" && pwd)"
elif [ -d "/Applications/ExcelMerge.app/Contents/MacOS" ]; then
    APP_MACOS_DIR="/Applications/ExcelMerge.app/Contents/MacOS"
fi

if [ -n "$APP_MACOS_DIR" ]; then
    cd "$APP_MACOS_DIR"
    exec "./ExcelMerge" "$@"
else
    # 开发模式：直接用python运行
    cd "$SCRIPT_DIR/.."
    exec python3 main.py "$@"
fi
