"""边界场景验证：交换左右、无同名 Sheet、点底部标签切到单侧独有 Sheet。"""

import os
import tempfile

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

from src.core.excel_loader import ExcelLoader
from src.gui.main_window import MainWindow

TMP = tempfile.mkdtemp()


def build(path, names):
    wb = Workbook()
    wb.active.title = names[0]
    wb.active["A1"] = f"{os.path.basename(path)}:{names[0]}"
    for n in names[1:]:
        wb.create_sheet(n)["A1"] = f"{os.path.basename(path)}:{n}"
    wb.save(path)
    return path


app = QApplication.instance() or QApplication([])
win = MainWindow()
win.show()


def wait():
    while (
        win._compare_worker is not None and win._compare_worker.isRunning()
    ) or (
        win._sheet_summary_worker is not None
        and win._sheet_summary_worker.isRunning()
    ):
        app.processEvents()
    for _ in range(5):
        app.processEvents()


def state(tag):
    wait()
    print(
        f"[{tag}] 左={win.left_sheet_name!r} 右={win.right_sheet_name!r} "
        f"左数据={'有' if win.left_sheet_data else '无'} "
        f"右数据={'有' if win.right_sheet_data else '无'} "
        f"底部激活={win.bottom_bar.get_active()!r} "
        f"红点={sorted(n for n, h in win._sheet_diff_summary.items() if h)} "
        f"缺失表={bool(win.diff_result and win.diff_result.sheet_missing)}"
    )


# --- 场景 A：两侧没有同名 Sheet（保持 Sheet 缺失比较） ---
lp = build(os.path.join(TMP, "a-left.xlsx"), ["甲表", "乙表"])
rp = build(os.path.join(TMP, "a-right.xlsx"), ["丙表"])
win._on_workbook_loaded("left", lp, ExcelLoader.load_workbook(lp), None)
state("A1-仅左侧")
win._on_workbook_loaded("right", rp, ExcelLoader.load_workbook(rp), None)
state("A2-两侧无同名表")

# --- 场景 B：交换左右（两侧需先能同表） ---
lp = build(os.path.join(TMP, "b-left.xlsx"), ["S1", "S2"])
rp = build(os.path.join(TMP, "b-right.xlsx"), ["S1", "S2"])
win._on_workbook_loaded("left", lp, ExcelLoader.load_workbook(lp), None)
win._on_workbook_loaded("right", rp, ExcelLoader.load_workbook(rp), None)
win.left_sheet_combo.setCurrentIndex(win.left_sheet_combo.findText("S2"))
state("B1-手动切左侧")
win.swap_sides()
state("B2-交换后")

# --- 场景 C：点底部标签切到仅左侧存在的 Sheet ---
win.bottom_bar._tabs["甲表"].click() if "甲表" in win.bottom_bar._tabs else None
state("C1-原标签栏")
lp2 = build(os.path.join(TMP, "c-left.xlsx"), ["S1", "独有表"])
rp2 = build(os.path.join(TMP, "c-right.xlsx"), ["S1"])
win._on_workbook_loaded("left", lp2, ExcelLoader.load_workbook(lp2), None)
win._on_workbook_loaded("right", rp2, ExcelLoader.load_workbook(rp2), None)
state("C2-重新打开文件对")
win.bottom_bar._tabs["独有表"].click()
state("C3-点单侧独有标签")
