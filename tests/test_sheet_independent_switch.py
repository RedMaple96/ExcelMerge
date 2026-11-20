"""两侧 Sheet 独立切换的 GUI 回归测试。

交互约定：
- 手动切换单侧下拉：只影响该侧，对侧保持不动（允许两侧停在不同 Sheet）；
- 打开/替换文件：两侧对齐到同一个 Sheet；
- 点击底部标签栏：两侧同步到同一个 Sheet。
"""

import os
import shutil
import tempfile
import unittest
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication, QMessageBox

from src.core.excel_loader import ExcelLoader
from src.gui.main_window import MainWindow


class IndependentSheetSwitchTest(unittest.TestCase):
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

    # ------------------------------------------------------------------ #
    # 辅助
    # ------------------------------------------------------------------ #
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

    def _build_workbook(self, filename: str, sheet_names):
        """创建并保存一个工作簿，每个 Sheet 的 A1 写入侧别+表名便于核对。"""
        path = os.path.join(self.tmpdir, filename)
        wb = Workbook()
        wb.active.title = sheet_names[0]
        wb.active["A1"] = f"{filename}:{sheet_names[0]}"
        for name in sheet_names[1:]:
            ws = wb.create_sheet(name)
            ws["A1"] = f"{filename}:{name}"
        wb.save(path)
        return path, wb

    def _load_both(self, left_sheets, right_sheets):
        """走真实打开流程加载两侧文件（后台线程回调入口）。"""
        left_path, left_wb = self._build_workbook("left.xlsx", left_sheets)
        right_path, right_wb = self._build_workbook("right.xlsx", right_sheets)
        self.window._on_workbook_loaded("left", left_path, left_wb, None)
        self.window._on_workbook_loaded("right", right_path, right_wb, None)
        self._wait_for_workers()

    def _select_via_combo(self, side: str, name: str):
        """模拟用户在下拉框里手动切换 Sheet（触发 currentIndexChanged）。"""
        combo = (
            self.window.left_sheet_combo if side == "left" else self.window.right_sheet_combo
        )
        combo.setCurrentIndex(combo.findText(name))
        self._wait_for_workers()

    def _side_state(self, side: str):
        """返回 (当前 Sheet 名, 下拉显示文本, 是否已加载数据)。"""
        name = self.window.left_sheet_name if side == "left" else self.window.right_sheet_name
        combo = (
            self.window.left_sheet_combo if side == "left" else self.window.right_sheet_combo
        )
        data = self.window.left_sheet_data if side == "left" else self.window.right_sheet_data
        return name, combo.currentText(), data is not None

    # ------------------------------------------------------------------ #
    # 打开文件：两侧对齐到同一 Sheet
    # ------------------------------------------------------------------ #
    def test_open_aligns_both_sides_to_same_sheet(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])

        self.assertEqual(self._side_state("left")[:2], ("sheet1", "sheet1"))
        self.assertEqual(self._side_state("right")[:2], ("sheet1", "sheet1"))
        self.assertEqual(self.window.bottom_bar.get_active(), "sheet1")

    def test_open_falls_back_to_common_sheet_when_first_is_missing(self):
        # 右侧首个 Sheet 左侧没有 → 退回两侧共有的第一个 Sheet
        self._load_both(["sheet1", "sheet2"], ["sheet3", "sheet2"])

        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        self.assertEqual(self._side_state("right")[:2], ("sheet2", "sheet2"))
        self.assertFalse(self.window.diff_result.sheet_missing)

    # ------------------------------------------------------------------ #
    # 手动切换：只影响该侧
    # ------------------------------------------------------------------ #
    def test_manual_switch_left_does_not_move_right(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])

        self._select_via_combo("left", "sheet2")

        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        # 对侧保持原 Sheet，不被强制同步
        self.assertEqual(self._side_state("right")[:2], ("sheet1", "sheet1"))
        self.assertTrue(self._side_state("right")[2])

    def test_manual_switch_right_does_not_move_left(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])

        self._select_via_combo("right", "sheet2")

        self.assertEqual(self._side_state("right")[:2], ("sheet2", "sheet2"))
        self.assertEqual(self._side_state("left")[:2], ("sheet1", "sheet1"))
        self.assertTrue(self._side_state("left")[2])

    def test_manual_switch_both_sides_to_different_sheets_keeps_both_loaded(self):
        self._load_both(["sheet1", "sheet2", "sheet3"], ["sheet1", "sheet2", "sheet3"])

        self._select_via_combo("left", "sheet2")
        self._select_via_combo("right", "sheet3")

        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        self.assertEqual(self._side_state("right")[:2], ("sheet3", "sheet3"))
        self.assertTrue(self._side_state("left")[2])
        self.assertTrue(self._side_state("right")[2])
        # 两侧都在场 → 跨表比较，不是 Sheet 缺失
        self.assertFalse(self.window.diff_result.sheet_missing)

    def test_manual_switch_to_sheet_missing_on_other_side_keeps_other_side(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1"])

        self._select_via_combo("left", "sheet2")

        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        # 对侧没有该 Sheet，但不应被清空（旧行为会把对侧置为缺失）
        self.assertEqual(self._side_state("right")[:2], ("sheet1", "sheet1"))
        self.assertTrue(self._side_state("right")[2])
        self.assertFalse(self.window.diff_result.sheet_missing)

    # ------------------------------------------------------------------ #
    # 底部标签栏：两侧同步
    # ------------------------------------------------------------------ #
    def test_bottom_tab_click_syncs_both_sides(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])
        self._select_via_combo("left", "sheet2")
        self.assertEqual(self._side_state("right")[:2], ("sheet1", "sheet1"))

        self.window.bottom_bar._tabs["sheet2"].click()
        self._wait_for_workers()

        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        self.assertEqual(self._side_state("right")[:2], ("sheet2", "sheet2"))
        self.assertEqual(self.window.bottom_bar.get_active(), "sheet2")

    def test_cancel_bottom_tab_switch_restores_active_tab(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])
        self._select_via_combo("left", "sheet2")
        self.assertEqual(self.window.bottom_bar.get_active(), "sheet2")

        self.window._dirty["left"] = True
        with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.Cancel):
            self.window.bottom_bar._tabs["sheet1"].click()
        self._wait_for_workers()

        # 取消后两侧与激活标签都停在切换前
        self.assertEqual(self.window.bottom_bar.get_active(), "sheet2")
        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        self.assertEqual(self._side_state("right")[:2], ("sheet1", "sheet1"))

    def test_manual_switch_cancel_restores_combo(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])

        self.window._dirty["left"] = True
        with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.Cancel):
            self._select_via_combo("left", "sheet2")

        self.assertEqual(self._side_state("left")[:2], ("sheet1", "sheet1"))
        self.assertEqual(self.window.left_sheet_data.worksheet.title, "sheet1")

    # ------------------------------------------------------------------ #
    # 未保存检查只覆盖受影响的侧
    # ------------------------------------------------------------------ #
    def test_manual_switch_ignores_unsaved_changes_on_other_side(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])

        # 对侧有未保存修改，但本次只切换左侧：不应弹出确认框
        self.window._dirty["right"] = True
        with mock.patch.object(QMessageBox, "question") as question:
            self._select_via_combo("left", "sheet2")

        question.assert_not_called()
        self.assertEqual(self._side_state("left")[:2], ("sheet2", "sheet2"))
        self.assertTrue(self.window._dirty["right"])
        self.assertIsNotNone(self.window.right_sheet_data)

    def test_manual_switch_asks_when_switched_side_is_dirty(self):
        self._load_both(["sheet1", "sheet2"], ["sheet1", "sheet2"])

        self.window._dirty["left"] = True
        with mock.patch.object(
            QMessageBox, "question", return_value=QMessageBox.Cancel
        ) as question:
            self._select_via_combo("left", "sheet2")

        question.assert_called_once()
        self.assertEqual(self._side_state("left")[:2], ("sheet1", "sheet1"))


if __name__ == "__main__":
    unittest.main()
