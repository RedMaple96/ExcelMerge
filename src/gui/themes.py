"""QSS 主题模块 — 浅色/深色样式表与系统外观检测。

提供 Beyond-Compare 风格的简洁样式，自动跟随 macOS 系统外观。
对应需求：FR-03（Task 6.4）与 Task 11（macOS 暗色模式适配）。
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication, QPalette


# ---------------------------------------------------------------------- #
# 浅色主题：白底深字、浅灰网格线
# ---------------------------------------------------------------------- #
LIGHT_QSS = """
QMainWindow, QWidget {
    background-color: #ffffff;
    color: #222222;
    font-size: 13px;
}

QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #f7f7f7;
    color: #222222;
    gridline-color: #d0d0d0;
    selection-background-color: #cde6f7;
    selection-color: #111111;
    border: 1px solid #d0d0d0;
}
QTableWidget::item {
    padding: 2px 4px;
}
QHeaderView::section {
    background-color: #f0f0f0;
    color: #222222;
    padding: 4px 6px;
    border: none;
    border-right: 1px solid #d0d0d0;
    border-bottom: 1px solid #d0d0d0;
    font-weight: 600;
}

QMenuBar {
    background-color: #f0f0f0;
    color: #222222;
    border-bottom: 1px solid #d0d0d0;
}
QMenuBar::item {
    padding: 4px 10px;
    background: transparent;
}
QMenuBar::item:selected {
    background-color: #cde6f7;
}
QMenu {
    background-color: #ffffff;
    color: #222222;
    border: 1px solid #d0d0d0;
}
QMenu::item:selected {
    background-color: #cde6f7;
}

QToolBar {
    background-color: #f0f0f0;
    border: none;
    border-bottom: 1px solid #d0d0d0;
    spacing: 2px;
    padding: 2px;
}
QToolBar::separator {
    width: 1px;
    background-color: #d0d0d0;
    margin: 4px 4px;
}

QStatusBar {
    background-color: #f0f0f0;
    color: #333333;
    border-top: 1px solid #d0d0d0;
}
QStatusBar::item { border: none; }

QFrame#BottomBar {
    background-color: #f0f0f0;
    border-top: 1px solid #d0d0d0;
    border-bottom: 1px solid #d0d0d0;
}
QFrame#BottomBar QLabel {
    color: #333333;
    font-size: 12px;
}
QFrame#BottomBar QPushButton {
    background-color: #e4e4e4;
    color: #222222;
    border: 1px solid #c0c0c0;
    border-radius: 3px;
    padding: 1px 8px;
    font-size: 12px;
    min-height: 16px;
}
QFrame#BottomBar QPushButton:hover {
    background-color: #d8d8d8;
    border-color: #909090;
}
QFrame#BottomBar QPushButton:checked {
    background-color: #cde6f7;
    border-color: #4a90d9;
    color: #111111;
}
QFrame#BottomBar QPushButton:disabled {
    background-color: #f0f0f0;
    color: #a0a0a0;
    border-color: #d0d0d0;
}

QComboBox {
    background-color: #ffffff;
    color: #222222;
    border: 1px solid #c0c0c0;
    border-radius: 2px;
    padding: 2px 6px;
    min-height: 18px;
}
QComboBox:hover { border-color: #909090; }
QComboBox::drop-down { border: none; width: 18px; }
QComboBox QAbstractItemView {
    background-color: #ffffff;
    color: #222222;
    selection-background-color: #cde6f7;
    border: 1px solid #c0c0c0;
}

QLabel { color: #222222; }

QPushButton {
    background-color: #f5f5f5;
    color: #222222;
    border: 1px solid #c0c0c0;
    border-radius: 3px;
    padding: 4px 12px;
    min-height: 18px;
}
QPushButton:hover { background-color: #e8e8e8; }
QPushButton:pressed { background-color: #d8d8d8; }
QPushButton:default { border-color: #4a90d9; }

QFrame[frameShape="6"] {
    background-color: #fafafa;
    border: 1px solid #e0e0e0;
    border-radius: 2px;
}
"""


# ---------------------------------------------------------------------- #
# 深色主题：深底浅字、深灰网格线（类 VS Code Dark）
# ---------------------------------------------------------------------- #
DARK_QSS = """
QMainWindow, QWidget {
    background-color: #1e1e1e;
    color: #d4d4d4;
    font-size: 13px;
}

QTableWidget {
    background-color: #1e1e1e;
    alternate-background-color: #252525;
    color: #d4d4d4;
    gridline-color: #3a3a3a;
    selection-background-color: #264f78;
    selection-color: #ffffff;
    border: 1px solid #3a3a3a;
}
QTableWidget::item {
    padding: 2px 4px;
}
QHeaderView::section {
    background-color: #2d2d2d;
    color: #d4d4d4;
    padding: 4px 6px;
    border: none;
    border-right: 1px solid #3a3a3a;
    border-bottom: 1px solid #3a3a3a;
    font-weight: 600;
}

QMenuBar {
    background-color: #2d2d2d;
    color: #d4d4d4;
    border-bottom: 1px solid #3a3a3a;
}
QMenuBar::item {
    padding: 4px 10px;
    background: transparent;
}
QMenuBar::item:selected {
    background-color: #3a3a3a;
}
QMenu {
    background-color: #2d2d2d;
    color: #d4d4d4;
    border: 1px solid #3a3a3a;
}
QMenu::item:selected {
    background-color: #264f78;
}

QToolBar {
    background-color: #2d2d2d;
    border: none;
    border-bottom: 1px solid #3a3a3a;
    spacing: 2px;
    padding: 2px;
}
QToolBar::separator {
    width: 1px;
    background-color: #3a3a3a;
    margin: 4px 4px;
}

QStatusBar {
    background-color: #2d2d2d;
    color: #c0c0c0;
    border-top: 1px solid #3a3a3a;
}
QStatusBar::item { border: none; }

QFrame#BottomBar {
    background-color: #2d2d2d;
    border-top: 1px solid #3a3a3a;
    border-bottom: 1px solid #3a3a3a;
}
QFrame#BottomBar QLabel {
    color: #c0c0c0;
    font-size: 12px;
}
QFrame#BottomBar QPushButton {
    background-color: #3a3a3a;
    color: #d4d4d4;
    border: 1px solid #4a4a4a;
    border-radius: 3px;
    padding: 1px 8px;
    font-size: 12px;
    min-height: 16px;
}
QFrame#BottomBar QPushButton:hover {
    background-color: #454545;
    border-color: #6a6a6a;
}
QFrame#BottomBar QPushButton:checked {
    background-color: #264f78;
    border-color: #3a7bc8;
    color: #ffffff;
}
QFrame#BottomBar QPushButton:disabled {
    background-color: #2d2d2d;
    color: #5a5a5a;
    border-color: #3a3a3a;
}

QComboBox {
    background-color: #2d2d2d;
    color: #d4d4d4;
    border: 1px solid #4a4a4a;
    border-radius: 2px;
    padding: 2px 6px;
    min-height: 18px;
}
QComboBox:hover { border-color: #6a6a6a; }
QComboBox::drop-down { border: none; width: 18px; }
QComboBox QAbstractItemView {
    background-color: #2d2d2d;
    color: #d4d4d4;
    selection-background-color: #264f78;
    border: 1px solid #4a4a4a;
}

QLabel { color: #d4d4d4; }

QPushButton {
    background-color: #3a3a3a;
    color: #d4d4d4;
    border: 1px solid #4a4a4a;
    border-radius: 3px;
    padding: 4px 12px;
    min-height: 18px;
}
QPushButton:hover { background-color: #454545; }
QPushButton:pressed { background-color: #2d2d2d; }
QPushButton:default { border-color: #264f78; }

QFrame[frameShape="6"] {
    background-color: #252525;
    border: 1px solid #3a3a3a;
    border-radius: 2px;
}
"""


def is_dark_mode() -> bool:
    """检测系统当前是否为暗色模式。

    优先使用 PySide6 6.5+ 的 styleHints().colorScheme()（macOS 原生支持），
    回退方案：比较 QPalette.Window 与 WindowText 的亮度。
    """
    sh = QGuiApplication.styleHints()
    if sh is not None and hasattr(sh, "colorScheme"):
        try:
            return sh.colorScheme() == Qt.ColorScheme.Dark
        except Exception:  # noqa: BLE001
            pass

    app = QGuiApplication.instance()
    if app is not None:
        pal = app.palette()
        try:
            window = pal.color(QPalette.Window)
            text = pal.color(QPalette.WindowText)
            return window.lightness() < text.lightness()
        except Exception:  # noqa: BLE001
            pass
    return False


def apply_theme(app) -> None:
    """根据系统外观为 QApplication 应用对应的 QSS 样式表。"""
    if is_dark_mode():
        app.setStyleSheet(DARK_QSS)
    else:
        app.setStyleSheet(LIGHT_QSS)
