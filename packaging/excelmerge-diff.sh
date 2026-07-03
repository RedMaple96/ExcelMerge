#!/bin/bash
# ExcelMerge 命令行启动脚本
# 此脚本打包在 ExcelMerge.app/Contents/Resources/ 目录下
# 用于 Git difftool / SourceTree 配置
# 用法: excelmerge-diff <left-file> <right-file>

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# 从脚本目录向上查找 .app 包路径
find_app_path() {
    local dir="$SCRIPT_DIR"
    while [ "$dir" != "/" ]; do
        # 检查当前目录是否是 .app 包根目录（包含 Contents/MacOS/ExcelMerge）
        if [ -x "$dir/Contents/MacOS/ExcelMerge" ]; then
            echo "$dir"
            return 0
        fi
        dir="$(dirname "$dir")"
    done
    return 1
}

APP_PATH="$(find_app_path)"

if [ -n "$APP_PATH" ] && [ -d "$APP_PATH" ]; then
    # 使用 open 命令启动，这是 macOS 上正确启动 GUI 应用的方式
    # -n: 启动新实例（即使已有一个在运行）
    # --args: 后面的参数传递给应用
    exec open -n "$APP_PATH" --args "$@"
else
    # 回退：尝试直接在脚本同级或上级目录找二进制（开发环境）
    for try_dir in "$SCRIPT_DIR" "$SCRIPT_DIR/../MacOS" "$SCRIPT_DIR/.."; do
        if [ -x "$try_dir/ExcelMerge" ]; then
            cd "$try_dir"
            exec "./ExcelMerge" "$@"
        fi
    done
    echo "❌ 找不到 ExcelMerge 应用程序" >&2
    echo "   请确认 ExcelMerge.app 完整性" >&2
    exit 1
fi
