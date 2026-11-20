"""行级差异导航（上一个差异 / 下一个差异）定位精度的 GUI 回归测试。

行为约定：定位到差异行的**具体差异单元格**，而不是该行行首单元格；
整行独占（仅左/仅右，含整表缺失）时定位到该侧第一个非虚拟列。
"""

import os
import shutil
import tempfile
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

from src.core.excel_loader import ExcelLoader
from src.gui.main_window import MainWindow


class DiffNavigationTest(unittest.TestCase):
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
        for _ in range(3000):
            self.app.processEvents()
            if (
                self.window._compare_worker is None
                and self.window._sheet_summary_worker is None
            ):
                return
        raise AssertionError("后台线程未在预期时间内结束")

    def _build(self, filename, rows, sheet="S1"):
        path = os.path.join(self.tmpdir, filename)
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        for row in rows:
            ws.append(row)
        wb.save(path)
        return path, wb

    def _open_pair(self, left_rows, right_rows):
        left_path, left_wb = self._build("left.xlsx", left_rows)
        right_path, right_wb = self._build("right.xlsx", right_rows)
        self.window._on_workbook_loaded("left", left_path, left_wb, None)
        self.window._on_workbook_loaded("right", right_path, right_wb, None)
        self._wait_for_workers()

    def _current_cells(self):
        return (
            (self.window.left_table.currentRow(), self.window.left_table.currentColumn()),
            (self.window.right_table.currentRow(), self.window.right_table.currentColumn()),
        )

    # ------------------------------------------------------------------ #
    # 单元格级定位
    # ------------------------------------------------------------------ #
    def test_next_diff_lands_on_the_differing_cell(self):
        header = ["ID", "名称", "价格", "区域"]
        self._open_pair(
            [header, [1, "A", 10, "X"], [2, "B", 20, "Y"], [3, "C", 30, "Z"]],
            [header, [1, "A", 10, "X"], [2, "B改", 20, "Y"], [3, "C", 30, "Z"]],
        )
        # 只有第 2 个数据行（对齐行 2）的第 1 列（名称）不同
        self.assertEqual(self.window.diff_result.diff_row_indices, [2])

        self.window.next_diff()

        self.assertEqual(self._current_cells(), ((2, 1), (2, 1)))
        self.assertEqual(self.window.status_label.text(), "差异 1/1")

    def test_next_diff_uses_leftmost_differing_cell_of_row(self):
        header = ["ID", "名称", "价格", "区域"]
        self._open_pair(
            [header, [1, "A", 10, "X"], [2, "B", 20, "Y"]],
            [header, [1, "A", 10, "X"], [2, "B", 999, "Y改"]],
        )
        # 第 2 个数据行第 2、3 列都不同 → 定位到最靠左的差异列（2）
        self.assertEqual(self.window.diff_result.aligned_rows[2].diff_cells, [2, 3])

        self.window.next_diff()

        self.assertEqual(self._current_cells(), ((2, 2), (2, 2)))

    def test_prev_diff_lands_on_the_differing_cell(self):
        header = ["ID", "名称", "价格"]
        self._open_pair(
            [header, [1, "A", 10], [2, "B", 20], [3, "C", 30]],
            [header, [1, "A", 10], [2, "B改", 20], [3, "C改", 30]],
        )
        self.window.next_diff()  # -> 对齐行 2
        self.window.next_diff()  # -> 对齐行 3
        self.assertEqual(self.window.left_table.currentRow(), 3)

        self.window.prev_diff()  # 回到对齐行 2，仍应落在具体差异列
        self.assertEqual(self._current_cells(), ((2, 1), (2, 1)))

    def test_column_order_change_maps_to_differing_cell(self):
        # 列顺序不同：右表把"价格"放在最前，列按列名对齐后差异仍在"价格"列
        self._open_pair(
            [["ID", "名称", "价格"], [1, "A", 10]],
            [["价格", "ID", "名称"], [99, 1, "A"]],
        )
        result = self.window.diff_result
        diff_col = result.aligned_rows[result.diff_row_indices[0]].diff_cells[0]
        self.assertEqual(result.aligned_cols[diff_col].label, "价格")

        self.window.next_diff()

        row = result.diff_row_indices[0]
        self.assertEqual(self._current_cells(), ((row, diff_col), (row, diff_col)))

    def test_each_row_lands_on_its_own_differing_cell(self):
        # 不同差异行的差异列不同：应各按自己的差异列定位
        header = ["ID", "名称", "价格", "区域"]
        self._open_pair(
            [header, [1, "A", 10, "X"], [2, "B", 20, "Y"]],
            [header, [1, "A改", 10, "X"], [2, "B", 20, "Y改"]],
        )
        result = self.window.diff_result
        self.assertEqual(
            [(i, result.aligned_rows[i].diff_cells) for i in result.diff_row_indices],
            [(1, [1]), (2, [3])],
        )

        self.window.next_diff()
        self.assertEqual(self._current_cells(), ((1, 1), (1, 1)))

        self.window.next_diff()
        self.assertEqual(self._current_cells(), ((2, 3), (2, 3)))

    # ------------------------------------------------------------------ #
    # 整行独占 / 整表缺失
    # ------------------------------------------------------------------ #
    def test_right_only_row_lands_on_own_first_real_column(self):
        # 左表多一列（对右侧是虚拟空列）、右表多一行：独占行应定位到右侧自己的列
        self._open_pair(
            [["旧列", "ID"], [9, 1]],
            [["ID"], [1], [2]],
        )
        result = self.window.diff_result
        self.assertEqual(result.aligned_cols[0].status, "left_only")
        self.assertIsNone(result.aligned_cols[0].right_col)

        self.window.next_diff()

        row = result.diff_row_indices[0]
        pair = result.aligned_rows[row]
        self.assertEqual(pair.status, "right_only")
        self.assertEqual(pair.diff_cells, [])
        # 列 0 在右侧是虚拟空列，应跳到右侧第一个真实列（列 1）
        self.assertEqual(self._current_cells(), ((row, 1), (row, 1)))
        self.assertIsNotNone(self.window.right_table.item(row, 1))

    def test_left_only_row_lands_on_own_first_real_column(self):
        # 对称场景：右表多一列、左表多一行
        self._open_pair(
            [["ID"], [1], [2]],
            [["新列", "ID"], [9, 1]],
        )
        result = self.window.diff_result
        self.assertEqual(result.aligned_cols[0].status, "right_only")

        self.window.next_diff()

        row = result.diff_row_indices[0]
        self.assertEqual(result.aligned_rows[row].status, "left_only")
        self.assertEqual(self._current_cells(), ((row, 1), (row, 1)))
        self.assertIsNotNone(self.window.left_table.item(row, 1))

    def test_sheet_missing_lands_on_first_real_column(self):
        left_path, left_wb = self._build("left.xlsx", [["ID", "名称"], [1, "A"]])
        right_path, right_wb = self._build(
            "right.xlsx", [["ID", "名称"], [1, "A"]], sheet="另一个表"
        )
        self.window._on_workbook_loaded("left", left_path, left_wb, None)
        self.window._on_workbook_loaded("right", right_path, right_wb, None)
        self._wait_for_workers()
        # 两侧无同名 Sheet → 缺失态比较
        self.assertTrue(self.window.diff_result.sheet_missing)

        self.window.next_diff()

        self.assertEqual(self._current_cells(), ((0, 0), (0, 0)))
        self.assertIsNotNone(self.window.left_table.item(0, 0))

    # ------------------------------------------------------------------ #
    # 边界
    # ------------------------------------------------------------------ #
    def test_navigation_without_diff_keeps_status(self):
        header = ["ID", "名称"]
        self._open_pair([header, [1, "A"]], [header, [1, "A"]])
        before = self._current_cells()

        self.window.next_diff()

        self.assertEqual(self.window.status_label.text(), "无差异")
        self.assertEqual(self._current_cells(), before)

    def test_ignored_column_difference_is_not_navigated(self):
        header = ["ID", "名称", "备注"]
        self._open_pair(
            [header, [1, "A", "x"]],
            [header, [1, "A", "y"]],  # 仅"备注"列不同
        )
        self.window.ignore_cols = [2]
        self.window.recompare()
        self._wait_for_workers()

        self.assertEqual(self.window.diff_result.diff_row_indices, [])
        self.window.next_diff()
        self.assertEqual(self.window.status_label.text(), "无差异")


if __name__ == "__main__":
    unittest.main()
