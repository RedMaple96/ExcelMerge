"""ExcelMerge — 应用程序入口。

负责初始化 QApplication、加载主窗口并启动事件循环。
支持命令行参数：可直接传入两个 .xlsx 文件路径进行比对，
格式：ExcelMerge <left-file> <right-file>
或兼容 Git difftool 参数格式（自动忽略 -L 等标签选项）。
"""

import os
import sys
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication

from src.gui.main_window import MainWindow
from src.gui.themes import apply_theme


def _set_macos_process_name(name: str) -> None:
    """macOS 上修改进程名，使菜单栏左侧应用菜单显示自定义名称。

    macOS 应用菜单标题取自 NSProcessInfo 的 processName（即进程名），
    默认为解释器名（如 Python）。Qt 的 setApplicationDisplayName 无法覆盖。
    通过 libc 的 setprogname 在 QApplication 创建前修改进程名即可生效。
    """
    if sys.platform != "darwin":
        return
    try:
        import ctypes

        libc = ctypes.CDLL("libc.dylib")
        libc.setprogname.restype = None
        libc.setprogname.argtypes = [ctypes.c_char_p]
        libc.setprogname(name.encode("utf-8"))
    except Exception:  # noqa: BLE001
        pass


def _activate_macos_app() -> None:
    """macOS 上强制激活应用为前台应用，确保菜单栏正确切换。

    当通过另一个应用（如 SourceTree）启动时，直接执行 Unix 二进制不会自动
    激活为前台应用，菜单栏仍然显示启动方的菜单。此函数通过 AppleScript
    调用 System Events 来激活 ExcelMerge 应用到前台。
    """
    if sys.platform != "darwin":
        return
    try:
        import subprocess
        # 使用AppleScript激活应用，这是macOS上最可靠的方式
        script = 'tell application "System Events" to set frontmost of the first process whose unix id is {pid} to true'.format(
            pid=os.getpid()
        )
        subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            timeout=5,
        )
    except Exception:  # noqa: BLE001
        pass


def _parse_cli_args(argv: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """解析命令行参数，提取左右两个文件路径。

    支持两种调用方式：
    1. 直接传入：ExcelMerge left.xlsx right.xlsx
    2. Git difftool 格式（含 -L 标签选项）：
       ExcelMerge -L "左标签" -L "右标签" /tmp/left.xlsx /tmp/right.xlsx

    返回 (left_path, right_path)，未提供时为 None。
    """
    args = argv[1:]
    files: List[str] = []
    i = 0
    while i < len(args):
        arg = args[i]
        # 跳过 -L 标签参数（Git difftool 格式）
        if arg == "-L":
            i += 2
            continue
        # 跳过其他以 - 开头的选项参数
        if arg.startswith("-") and len(arg) > 1:
            i += 1
            continue
        # 收集看起来像文件路径的参数（存在的文件或包含路径分隔符/扩展名）
        if os.path.exists(arg) or os.sep in arg or ("." in arg and not arg.startswith(".")):
            files.append(arg)
        i += 1

    left = files[0] if len(files) >= 1 else None
    right = files[1] if len(files) >= 2 else None
    return left, right


def main() -> int:
    """启动应用程序。"""
    # macOS：在创建 QApplication 前修改进程名，使应用菜单显示 "ExcelMerge"
    _set_macos_process_name("ExcelMerge")

    # 解析命令行参数
    left_path, right_path = _parse_cli_args(sys.argv)

    # macOS 高分屏支持
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("ExcelMerge")
    app.setApplicationDisplayName("ExcelMerge")
    app.setOrganizationName("ExcelMerge")

    # 默认字体：macOS 使用 SF Pro，其它平台交由 Qt 决定（Task 11.3）
    if sys.platform == "darwin":
        app.setFont(QFont("SF Pro", 13))

    # 应用浅色/深色主题（Task 6.4 / Task 11.2）
    apply_theme(app)

    window = MainWindow()

    # 若命令行传入了文件，在窗口显示后自动加载
    if left_path or right_path:
        def _load_files() -> None:
            if left_path and os.path.exists(left_path):
                window.load_file("left", left_path)
            if right_path and os.path.exists(right_path):
                window.load_file("right", right_path)
        QTimer.singleShot(0, _load_files)

    window.show()

    # macOS: 通过其他应用（如SourceTree）启动时，延迟激活应用到前台，确保菜单栏切换
    if sys.platform == "darwin":
        QTimer.singleShot(100, _activate_macos_app)

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
