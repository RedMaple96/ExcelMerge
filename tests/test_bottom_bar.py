"""底部 Sheet 导航栏交互测试。"""

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, QTimer
from PySide6.QtWidgets import QApplication, QListWidget, QMenu

from src.gui.bottom_bar import BottomBar


class BottomBarTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance()
        if cls.app is None:
            cls.app = QApplication([])

    def test_arrow_context_menu_is_scrollable(self):
        bar = BottomBar()
        self.addCleanup(bar.deleteLater)
        bar.set_sheets([f"Sheet-{index}" for index in range(120)], set())

        def close_menu():
            menu = bar.findChild(QMenu)
            if menu is not None:
                menu.close()

        QTimer.singleShot(0, close_menu)
        bar.btn_next.customContextMenuRequested.emit(QPoint(1, 1))
        self.app.processEvents()

        menu = bar.findChild(QMenu)
        self.assertIsNotNone(menu)
        sheet_list = menu.findChild(QListWidget)
        self.assertIsNotNone(sheet_list)
        self.assertEqual(sheet_list.count(), 120)
        self.assertGreater(sheet_list.verticalScrollBar().maximum(), 0)


if __name__ == "__main__":
    unittest.main()
