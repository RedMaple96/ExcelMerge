"""Sheet 级差异定位（上一个表差异 / 下一个表差异）的 GUI 回归测试。

行为约定：
- 只在“有差异的 Sheet”之间跳转（同名 Sheet 内容不同，或仅一侧存在的 Sheet）；
- 按底部标签顺序循环，连点可逐个走遍全部差异 Sheet；
- 跳转把两侧同步切到目标 Sheet；用户取消未保存确认时不切换。
"""

import os
import shutil
import tempfile
import unittest
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtCore import Qt
from PySide6.QtGui import QShortcut
from PySide6.QtWidgets import QApplication, QMessageBox

from src.core.excel_loader import ExcelLoader
from src.gui.main_window import MainWindow


def _fill(ws, rows):
    for row in rows:
        ws.append(row)


class SheetDiffNavigationTest(unittest.TestCase):
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
        """等待比较线程与 Sheet 差异预计算线程结束并处理完信号。"""
        for _ in range(3000):
            self.app.processEvents()
            if (
                self.window._compare_worker is None
                and self.window._sheet_summary_worker is None
            ):
                return
        raise AssertionError("后台线程未在预期时间内结束")

    def _build(self, filename, sheets):
        """sheets: {name: rows}，rows 为该表数据行（含表头行）。"""
        path = os.path.join(self.tmpdir, filename)
        wb = Workbook()
        wb.remove(wb.active)
        for name, rows in sheets.items():
            _fill(wb.create_sheet(name), rows)
        wb.save(path)
        return path, wb

    def _open_pair(self, left_sheets, right_sheets):
        left_path, left_wb = self._build("left.xlsx", left_sheets)
        right_path, right_wb = self._build("right.xlsx", right_sheets)
        self.window._on_workbook_loaded("left", left_path, left_wb, None)
        self.window._on_workbook_loaded("right", right_path, right_wb, None)
        self._wait_for_workers()

    @staticmethod
    def _same(value):
        return [["ID", "名称"], [1, value]]

    def _state(self):
        return (
            self.window.left_sheet_name,
            self.window.right_sheet_name,
            self.window.bottom_bar.get_active(),
        )

    # ------------------------------------------------------------------ #
    # 正常跳转
    # ------------------------------------------------------------------ #
    def test_next_sheet_diff_skips_identical_sheets(self):
        # S1 相同、S2 不同、S3 相同、S4 不同
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b"), "S3": self._same("c"), "S4": self._same("d")},
            {"S1": self._same("a"), "S2": self._same("B"), "S3": self._same("c"), "S4": self._same("D")},
        )
        self.assertEqual(self._state(), ("S1", "S1", "S1"))

        self.window.next_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S2", "S2", "S2"))
        self.assertEqual(self.window.status_label.text(), "Sheet 差异 1/2: S2")

        # 跳过相同的 S3，直接到 S4
        self.window.next_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S4", "S4", "S4"))
        self.assertEqual(self.window.status_label.text(), "Sheet 差异 2/2: S4")

    def test_next_sheet_diff_wraps_around(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b"), "S3": self._same("c")},
            {"S1": self._same("a"), "S2": self._same("B"), "S3": self._same("C")},
        )
        self.window.next_sheet_diff()  # -> S2
        self._wait_for_workers()
        self.window.next_sheet_diff()  # -> S3
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S3", "S3", "S3"))

        self.window.next_sheet_diff()  # 循环回到 S2
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S2", "S2", "S2"))
        self.assertEqual(self.window.status_label.text(), "Sheet 差异 1/2: S2")

    def test_prev_sheet_diff_wraps_to_last_diff_sheet(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b"), "S3": self._same("c")},
            {"S1": self._same("a"), "S2": self._same("B"), "S3": self._same("C")},
        )
        # 当前 S1 自身无差异：向前应循环到最后一个差异表 S3
        self.window.prev_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S3", "S3", "S3"))
        self.assertEqual(self.window.status_label.text(), "Sheet 差异 2/2: S3")

        self.window.prev_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S2", "S2", "S2"))

    def test_next_sheet_diff_reaches_sheet_missing_on_one_side(self):
        self._open_pair(
            {"S1": self._same("a"), "独有表": self._same("only-left")},
            {"S1": self._same("a")},
        )
        self.window.next_sheet_diff()
        self._wait_for_workers()

        self.assertEqual(self.window.left_sheet_name, "独有表")
        self.assertEqual(self.window.right_sheet_name, "独有表")
        self.assertIsNone(self.window.right_sheet_data)
        self.assertTrue(self.window.diff_result.sheet_missing)
        self.assertEqual(self.window.status_label.text(), "Sheet 差异 1/1: 独有表")

    def test_jump_syncs_both_sides_from_cross_sheet_state(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b"), "S3": self._same("c")},
            {"S1": self._same("a"), "S2": self._same("B"), "S3": self._same("c")},
        )
        # 先手动把两侧切到不同 Sheet（上一项改动后的独立切换）
        combo = self.window.left_sheet_combo
        combo.setCurrentIndex(combo.findText("S3"))
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S3", "S1", "S3"))

        # 定位以底部激活标签（S3）为基准，跳到差异表 S2
        self.window.next_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S2", "S2", "S2"))

    # ------------------------------------------------------------------ #
    # 无差异 / 取消
    # ------------------------------------------------------------------ #
    def test_no_diff_sheets_reports_status_and_keeps_sheet(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b")},
            {"S1": self._same("a"), "S2": self._same("b")},
        )
        self.window.next_sheet_diff()
        self._wait_for_workers()

        self.assertEqual(self._state(), ("S1", "S1", "S1"))
        self.assertEqual(self.window.status_label.text(), "无差异 Sheet")

        self.window.prev_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S1", "S1", "S1"))
        self.assertEqual(self.window.status_label.text(), "无差异 Sheet")

    def test_cancel_unsaved_keeps_current_sheet_and_drops_pending_status(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b")},
            {"S1": self._same("a"), "S2": self._same("B")},
        )
        self.window._dirty["left"] = True
        with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.Cancel):
            self.window.next_sheet_diff()
        self._wait_for_workers()

        self.assertEqual(self._state(), ("S1", "S1", "S1"))
        self.assertEqual(self.window.status_label.text(), "就绪")

    def test_reports_progress_while_summary_still_computing(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b")},
            {"S1": self._same("a"), "S2": self._same("B")},
        )
        # 模拟预计算未完成：结果尚未回填且后台线程仍在运行
        self.window._sheet_diff_summary = {}
        running = mock.Mock()
        running.isRunning.return_value = True
        self.window._sheet_summary_worker = running
        self.addCleanup(setattr, self.window, "_sheet_summary_worker", None)

        self.window.next_sheet_diff()

        self.assertEqual(self._state(), ("S1", "S1", "S1"))
        self.assertEqual(self.window.status_label.text(), "Sheet 差异计算中，请稍候...")

    def test_sheet_diff_shortcuts_are_registered(self):
        """快捷键必须解析成真实按键：QKeySequence 不认 "PageUp" 拼写。"""
        combos = {
            sc.key().toString(): sc.key()[0]
            for sc in self.window.findChildren(QShortcut)
            if not sc.key().isEmpty()
        }
        self.assertEqual(combos["Ctrl+PgDown"].key(), Qt.Key_PageDown)
        self.assertEqual(combos["Ctrl+PgUp"].key(), Qt.Key_PageUp)

    def test_single_diff_sheet_stays_on_it(self):
        self._open_pair(
            {"S1": self._same("a"), "S2": self._same("b")},
            {"S1": self._same("a"), "S2": self._same("B")},
        )
        self.window.next_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S2", "S2", "S2"))

        self.window.next_sheet_diff()
        self._wait_for_workers()
        self.assertEqual(self._state(), ("S2", "S2", "S2"))
        self.assertEqual(self.window.status_label.text(), "Sheet 差异 1/1: S2")


if __name__ == "__main__":
    unittest.main()
