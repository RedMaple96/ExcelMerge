"""手工验证：上一个/下一个差异 定位到具体差异单元格。"""

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

HEADER = ["礼包ID", "礼包名称", "价格", "绑定区域"]
LEFT_ROWS = [
    [1001, "破产礼包·初", 6, "够保跑二期"],
    [1002, "破产礼包·中", 30, "够保跑二期"],
    [1003, "破产礼包·高", 98, "够保跑二期"],
    [1004, "破产礼包·至尊", 648, "够保跑二期"],
]
RIGHT_ROWS = [
    [1001, "破产礼包·初", 6, "够保跑二期"],
    [1002, "破产礼包·中", 30, "够保跑二期"],
    [1003, "破产礼包·高级", 128, "够保跑二期"],   # 名称、价格都不同
    [1004, "破产礼包·至尊", 648, "够保跑二期"],
]


def build(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "够保跑二期"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


left_path = build(os.path.join(TMP, "081_破产礼包.xlsx"), LEFT_ROWS)
right_path = build(os.path.join(TMP, "081_破产礼包_新.xlsx"), RIGHT_ROWS)

app = QApplication.instance() or QApplication([])
win = MainWindow()
win.resize(1500, 620)
win.show()


def wait():
    for _ in range(3000):
        app.processEvents()
        if win._compare_worker is None and win._sheet_summary_worker is None:
            break
    for _ in range(5):
        app.processEvents()


def report(tag):
    wait()
    row = win.left_table.currentRow()
    col = win.left_table.currentColumn()
    rrow = win.right_table.currentRow()
    rcol = win.right_table.currentColumn()
    pair = win.diff_result.aligned_rows[row] if row >= 0 else None
    print(
        f"[{tag}] 左侧=({row},{col}) 右侧=({rrow},{rcol}) "
        f"底部坐标={win.bottom_bar.coord_label.text()!r} "
        f"行状态={pair.status if pair else None} 差异列={pair.diff_cells if pair else None} "
        f"状态栏={win.status_label.text()!r}"
    )


win._on_workbook_loaded("left", left_path, ExcelLoader.load_workbook(left_path), None)
win._on_workbook_loaded("right", right_path, ExcelLoader.load_workbook(right_path), None)
wait()
print("表头列:", [cp.label for cp in win.diff_result.aligned_cols])
print("差异行:", win.diff_result.diff_row_indices,
      [(i, p.diff_cells) for i, p in enumerate(win.diff_result.aligned_rows) if p.status != "same"])

win.next_diff()
report("下一个差异")
win.grab().save(os.path.join(OUT, "9-差异定位-具体单元格.png"))

win.next_diff()
report("下一个差异（循环回第 1 个）")

win.prev_diff()
report("上一个差异")
