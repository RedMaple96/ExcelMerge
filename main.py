"""Excel Merge — 应用程序入口。

负责初始化 QApplication、加载主窗口并启动事件循环。
"""

import sys

from PySide6.QtCore import Qt
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


def main() -> int:
    """启动应用程序。"""
    # macOS：在创建 QApplication 前修改进程名，使应用菜单显示 "Excel Merge"
    _set_macos_process_name("Excel Merge")

    # macOS 高分屏支持
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("Excel Merge")
    app.setApplicationDisplayName("Excel Merge")
    app.setOrganizationName("ExcelMerge")

    # 默认字体：macOS 使用 SF Pro，其它平台交由 Qt 决定（Task 11.3）
    if sys.platform == "darwin":
        app.setFont(QFont("SF Pro", 13))

    # 应用浅色/深色主题（Task 6.4 / Task 11.2）
    apply_theme(app)

    window = MainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
