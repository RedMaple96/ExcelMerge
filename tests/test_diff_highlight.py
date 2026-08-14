"""字符级差异高亮 delegate 的冒烟回归测试。

覆盖 delegate 的真实绘制路径：曾经因缺失 QStyle 导入在 paint 中抛
NameError，进而导致 GUI 渲染崩溃。此用例在 offscreen 模式下渲染一次
携带高亮区间的表格，确保绘制路径可用。
"""

import os
import unittest

# 必须在导入/创建 QApplication 之前设置，避免需要真实显示环境
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QApplication, QTableWidget, QTableWidgetItem

from src.core.chardiff import char_diff_ranges
from src.gui.diff_highlight_delegate import (
    CELL_WRAP_ROLE,
    HIGHLIGHT_RANGES_ROLE,
    DiffHighlightDelegate,
)
from src.gui.main_window import MainWindow


def _make_app():
    app = QApplication.instance()
    if app is None:
        try:
            app = QApplication([])
        except Exception:  # noqa: BLE001
            return None
    return app


class DiffHighlightDelegateTest(unittest.TestCase):
    """delegate 绘制路径冒烟测试。"""

    @classmethod
    def setUpClass(cls):
        cls.app = _make_app()

    @classmethod
    def tearDownClass(cls):
        if cls.app is not None:
            cls.app.processEvents()

    def test_paint_with_ranges_does_not_crash(self):
        if self.app is None:
            self.skipTest("无可用 Qt 平台插件，跳过 GUI 冒烟测试")

        left_text = "12345"
        right_text = "12678"
        left_ranges, right_ranges = char_diff_ranges(left_text, right_text)

        table = QTableWidget(1, 2)
        table.setItemDelegate(DiffHighlightDelegate(table))
        left_item = QTableWidgetItem(left_text)
        right_item = QTableWidgetItem(right_text)
        left_item.setData(HIGHLIGHT_RANGES_ROLE, left_ranges)
        right_item.setData(HIGHLIGHT_RANGES_ROLE, right_ranges)
        table.setItem(0, 0, left_item)
        table.setItem(0, 1, right_item)
        table.resize(240, 40)

        pixmap = QPixmap(240, 40)
        pixmap.fill()
        table.viewport().render(pixmap)

        # 高亮区间数据仍应为 [(2,5)]，且渲染未抛异常
        self.assertEqual(
            table.item(0, 0).data(HIGHLIGHT_RANGES_ROLE), [(2, 5)]
        )
        self.assertEqual(
            table.item(0, 1).data(HIGHLIGHT_RANGES_ROLE), [(2, 5)]
        )

    def test_paint_long_multiline_text_does_not_crash(self):
        if self.app is None:
            self.skipTest("无可用 Qt 平台插件，跳过 GUI 冒烟测试")

        text = "规则说明：\n1、购买礼包可获赠道具，使用道具可打开宝箱。\n" * 20
        table = QTableWidget(1, 1)
        table.setItemDelegate(DiffHighlightDelegate(table))
        item = QTableWidgetItem(text)
        item.setData(HIGHLIGHT_RANGES_ROLE, [(3, 12)])
        table.setItem(0, 0, item)
        table.resize(300, 200)

        pixmap = QPixmap(300, 200)
        pixmap.fill()
        table.viewport().render(pixmap)

    def test_double_click_autofits_cell(self):
        if self.app is None:
            self.skipTest("无可用 Qt 平台插件，跳过 GUI 冒烟测试")

        window = MainWindow()
        text = "第一行内容\n第二行是明显更长的内容内容内容内容内容"
        for table in (window.left_table, window.right_table):
            # 用两列：目标放第 0 列，避免末列被 stretchLastSection 拉宽
            table.setColumnCount(2)
            table.setRowCount(1)
            table.setItem(0, 0, QTableWidgetItem(text))

        initial_w = window.right_table.columnWidth(0)
        initial_h = window.right_table.rowHeight(0)
        window.right_table.cellDoubleClicked.emit(0, 0)
        self.app.processEvents()

        # 内容在窗口宽度 80% 内：只加宽列，不折行，左右同步
        self.assertGreater(window.right_table.columnWidth(0), initial_w)
        self.assertIsNone(window.right_table.item(0, 0).data(CELL_WRAP_ROLE))
        self.assertEqual(
            window.left_table.columnWidth(0),
            window.right_table.columnWidth(0),
        )
        self.assertEqual(
            window.left_table.rowHeight(0),
            window.right_table.rowHeight(0),
        )

        # 再次双击恢复初始列宽/行高
        window.right_table.cellDoubleClicked.emit(0, 0)
        self.app.processEvents()
        self.assertEqual(window.right_table.columnWidth(0), initial_w)
        self.assertEqual(window.right_table.rowHeight(0), initial_h)
        window.close()

    def test_double_click_wraps_when_exceeding_window_width(self):
        if self.app is None:
            self.skipTest("无可用 Qt 平台插件，跳过 GUI 冒烟测试")

        window = MainWindow()
        text = "很长的内容" * 200
        for table in (window.left_table, window.right_table):
            table.setColumnCount(2)
            table.setRowCount(1)
            table.setItem(0, 0, QTableWidgetItem(text))

        initial_w = window.right_table.columnWidth(0)
        initial_h = window.right_table.rowHeight(0)
        window.right_table.cellDoubleClicked.emit(0, 0)
        self.app.processEvents()

        # 超过窗口宽度 80%：列宽被限制到 80%，并启用折行
        self.assertEqual(
            window.right_table.columnWidth(0), int(window.width() * 0.8)
        )
        self.assertTrue(window.right_table.item(0, 0).data(CELL_WRAP_ROLE))
        self.assertGreater(window.right_table.rowHeight(0), initial_h)

        # 再次双击恢复
        window.right_table.cellDoubleClicked.emit(0, 0)
        self.app.processEvents()
        self.assertEqual(window.right_table.columnWidth(0), initial_w)
        self.assertEqual(window.right_table.rowHeight(0), initial_h)
        window.close()

    def test_diff_cell_has_light_background_and_char_ranges(self):
        if self.app is None:
            self.skipTest("无可用 Qt 平台插件，跳过 GUI 冒烟测试")

        window = MainWindow()
        for table in (window.left_table, window.right_table):
            table.setColumnCount(2)
            table.setRowCount(1)
        window.left_table.setItem(0, 0, QTableWidgetItem("12345"))
        window.right_table.setItem(0, 0, QTableWidgetItem("12678"))

        window._apply_char_highlight(0, 0)

        left_item = window.left_table.item(0, 0)
        right_item = window.right_table.item(0, 0)
        # 整格浅色背景，保证内容被截断时也能看出差异
        self.assertEqual(left_item.background().color().name().lower(), "#ffe0e0")
        self.assertEqual(right_item.background().color().name().lower(), "#ffe0e0")
        # 字符级高亮区间仍然保留
        self.assertEqual(left_item.data(HIGHLIGHT_RANGES_ROLE), [(2, 5)])
        self.assertEqual(right_item.data(HIGHLIGHT_RANGES_ROLE), [(2, 5)])
        window.close()


if __name__ == "__main__":
    unittest.main()
