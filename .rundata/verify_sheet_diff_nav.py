"""手工验证：Sheet 级差异定位按钮（上一个表差异 / 下一个表差异）。"""

import os
import tempfile

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

from src.core.excel_loader import ExcelLoader
from src.gui.main_window import MainWindow

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shots")
os.makedirs(OUT, exist_ok=True)
TMP = tempfile.mkdtemp()


def build(path, changed):
    """changed: 有差异的 sheet 名集合。"""
    wb = Workbook()
    wb.remove(wb.active)
    for name in ("够保跑二期", "礼包配置", "格式说明", "活动排期"):
        ws = wb.create_sheet(name)
        ws.append(["ID", "名称", "数值"])
        ws.append([1001, name, 100])
        if name in changed:
            ws.append([1002, f"{name}-新增", 200])
    wb.save(path)
    return path


left_path = build(os.path.join(TMP, "081_破产礼包.xlsx"), set())
right_path = build(os.path.join(TMP, "081_破产礼包_新.xlsx"), {"礼包配置", "活动排期"})

app = QApplication.instance() or QApplication([])
win = MainWindow()
win.resize(1500, 760)
win.show()


def wait():
    for _ in range(3000):
        app.processEvents()
        if win._compare_worker is None and win._sheet_summary_worker is None:
            break
    for _ in range(5):
        app.processEvents()


def state(tag):
    wait()
    print(
        f"[{tag}] 左={win.left_sheet_name!r} 右={win.right_sheet_name!r} "
        f"底部激活={win.bottom_bar.get_active()!r} 状态栏={win.status_label.text()!r}"
    )


def shot(tag):
    win.grab().save(os.path.join(OUT, f"{tag}.png"))


win._on_workbook_loaded("left", left_path, ExcelLoader.load_workbook(left_path), None)
win._on_workbook_loaded("right", right_path, ExcelLoader.load_workbook(right_path), None)
state("打开后")
print("  差异表:", [n for n, h in win._sheet_diff_summary.items() if h])

win.next_sheet_diff()
state("下一个表差异 1 次")
print("  差异表:", [n for n, h in win._sheet_diff_summary.items() if h])
shot("6-表差异定位-第1个差异表")

win.next_sheet_diff()
state("下一个表差异 2 次")
shot("7-表差异定位-第2个差异表")

win.next_sheet_diff()
state("下一个表差异 3 次（循环）")

win.prev_sheet_diff()
state("上一个表差异 1 次")

# 工具栏整体外观（含新按钮）
shot("8-工具栏")
print("工具栏动作:", [a.text() for a in win.toolbar.actions() if a.text()])

# 快捷键注册情况
from PySide6.QtGui import QShortcut

print(
    "已注册快捷键:",
    sorted(s.key().toString() for s in win.findChildren(QShortcut)),
)
