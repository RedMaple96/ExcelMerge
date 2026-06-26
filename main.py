"""表格比较与合并工具 — 应用程序入口。

负责初始化 QApplication、加载主窗口并启动事件循环。
"""

import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication

from src.gui.main_window import MainWindow
from src.gui.themes import apply_theme


def main() -> int:
    """启动应用程序。"""
    # macOS 高分屏支持
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("ExcelCompare")
    app.setApplicationDisplayName("表格比较")
    app.setOrganizationName("ExcelCompareOrg")

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
