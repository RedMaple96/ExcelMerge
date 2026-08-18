"""按名称同步 Sheet 的 GUI 回归测试。"""

import os
import shutil
import tempfile
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
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
        self.addCleanup(lambda: self.window._dirty.update(left=False, right=False))

    def _wait_for_workers(self):
        while (
            self.window._compare_worker is not None
            and self.window._compare_worker.isRunning()
        ) or (
            self.window._sheet_summary_worker is not None
            and self.window._sheet_summary_worker.isRunning()
        ):
            self.app.processEvents()
        self.app.processEvents()

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

    def test_copy_missing_sheet_to_other_workbook(self):
        left_wb = Workbook()
        left_wb.active.title = "sheet1"
        source = left_wb.create_sheet("sheet2")
        source["A1"] = "whole sheet"
        source["A1"].font = Font(bold=True)
        source["A1"].fill = PatternFill(
            start_color="FFFF0000", end_color="FFFF0000", fill_type="solid"
        )
        source.merge_cells("A1:B1")

        right_wb = Workbook()
        right_wb.active.title = "sheet1"

        self.window.left_wb = left_wb
        self.window.right_wb = right_wb
        for combo, names in (
            (self.window.left_sheet_combo, left_wb.sheetnames),
            (self.window.right_sheet_combo, right_wb.sheetnames),
        ):
            combo.blockSignals(True)
            combo.addItems(names)
            combo.blockSignals(False)
        self.window._load_sheet_data("left", "sheet1")
        self.window._load_sheet_data("right", "sheet1")
        self.window._switch_to_sheet("sheet2")

        self.window._copy_whole_sheet("left")
        self._wait_for_workers()

        target = right_wb["sheet2"]
        self.assertEqual(target["A1"].value, "whole sheet")
        self.assertTrue(target["A1"].font.bold)
        self.assertIn("FF0000", target["A1"].fill.fgColor.rgb)
        self.assertIn("A1:B1", {str(rng) for rng in target.merged_cells.ranges})
        self.assertEqual(self.window.right_sheet_combo.currentText(), "sheet2")
        self.assertIsNotNone(self.window.right_sheet_data)
        self.assertTrue(self.window._dirty["right"])

    def test_delete_sheet_missing_on_other_side(self):
        left_wb = Workbook()
        left_wb.active.title = "sheet1"
        left_wb.create_sheet("sheet2")["A1"] = "extra"
        left_wb.create_sheet("sheet3")

        right_wb = Workbook()
        right_wb.active.title = "sheet1"
        right_wb.create_sheet("sheet3")

        self.window.left_wb = left_wb
        self.window.right_wb = right_wb
        for combo, names in (
            (self.window.left_sheet_combo, left_wb.sheetnames),
            (self.window.right_sheet_combo, right_wb.sheetnames),
        ):
            combo.blockSignals(True)
            combo.addItems(names)
            combo.blockSignals(False)
        self.window._load_sheet_data("left", "sheet1")
        self.window._load_sheet_data("right", "sheet1")
        self.window._switch_to_sheet("sheet2")

        self.window._delete_whole_sheet("left")
        self._wait_for_workers()

        self.assertEqual(left_wb.sheetnames, ["sheet1", "sheet3"])
        self.assertNotIn("sheet2", self.window.left_sheet_combo.currentText())
        self.assertEqual(self.window.left_sheet_name, "sheet1")
        self.assertEqual(self.window.right_sheet_name, "sheet1")
        self.assertIsNotNone(self.window.left_sheet_data)
        self.assertIsNotNone(self.window.right_sheet_data)
        self.assertFalse(self.window.diff_result.sheet_missing)
        self.assertTrue(self.window._dirty["left"])


if __name__ == "__main__":
    unittest.main()
