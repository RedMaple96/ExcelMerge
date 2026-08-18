"""按名称同步 Sheet 的 GUI 回归测试。"""

import os
import shutil
import tempfile
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

from src.core.excel_loader import ExcelLoader
from src.gui.main_window import MainWindow


class SheetMatchingTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance()
        if cls.app is None:
            cls.app = QApplication([])

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmpdir, ignore_errors=True)
        self.window = MainWindow()
        self.addCleanup(self.window.close)

    def test_missing_right_sheet_is_not_matched_by_position(self):
        left_path = os.path.join(self.tmpdir, "left.xlsx")
        right_path = os.path.join(self.tmpdir, "right.xlsx")

        left_wb = Workbook()
        left_wb.active.title = "sheet1"
        for name in ("sheet2", "sheet3"):
            left_wb.create_sheet(name)["A1"] = name
        left_wb.save(left_path)

        right_wb = Workbook()
        right_wb.active.title = "sheet1"
        right_wb.create_sheet("sheet3")["A1"] = "sheet3"
        right_wb.save(right_path)

        self.window.left_wb = ExcelLoader.load_workbook(left_path)
        self.window.right_wb = ExcelLoader.load_workbook(right_path)
        for combo, names in (
            (self.window.left_sheet_combo, self.window.left_wb.sheetnames),
            (self.window.right_sheet_combo, self.window.right_wb.sheetnames),
        ):
            combo.blockSignals(True)
            combo.addItems(names)
            combo.blockSignals(False)
        self.window._load_sheet_data("left", "sheet1")
        self.window._load_sheet_data("right", "sheet1")

        self.window._switch_to_sheet("sheet2")

        self.assertEqual(self.window.left_sheet_name, "sheet2")
        self.assertEqual(self.window.right_sheet_name, "sheet2")
        self.assertIsNotNone(self.window.left_sheet_data)
        self.assertIsNone(self.window.right_sheet_data)
        self.assertTrue(self.window.diff_result.sheet_missing)
        self.assertEqual(self.window.diff_result.stats["left_only"], 1)
        self.assertEqual(self.window.right_sheet_combo.currentIndex(), -1)

        self.window._switch_to_sheet("sheet3")
        while (
            self.window._compare_worker is not None
            and self.window._compare_worker.isRunning()
        ):
            self.app.processEvents()
        self.assertEqual(self.window.left_sheet_name, "sheet3")
        self.assertEqual(self.window.right_sheet_name, "sheet3")
        self.assertIsNotNone(self.window.left_sheet_data)
        self.assertIsNotNone(self.window.right_sheet_data)
        self.assertEqual(self.window.right_sheet_combo.currentText(), "sheet3")

        self.window._switch_to_sheet("sheet2")
        self.window.swap_sides()

        self.assertEqual(self.window.left_sheet_name, "sheet2")
        self.assertEqual(self.window.right_sheet_name, "sheet2")
        self.assertIsNone(self.window.left_sheet_data)
        self.assertIsNotNone(self.window.right_sheet_data)
        self.assertEqual(self.window.left_sheet_combo.currentIndex(), -1)
        self.assertEqual(self.window.right_sheet_combo.currentText(), "sheet2")


if __name__ == "__main__":
    unittest.main()
