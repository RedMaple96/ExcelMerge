"""手工验证：两侧 Sheet 独立切换 / 打开对齐 / 底部标签同步切换。

离屏渲染 Qt 窗口并截图，用真实工作簿走真实信号链路（下拉 currentIndexChanged、
底部标签 click），产出 PNG 供人工核对。
"""

import os
import tempfile

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

from src.gui.main_window import MainWindow

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shots")
os.makedirs(OUT, exist_ok=True)
TMP = tempfile.mkdtemp()


def build(path, drop_second_sheet=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "够保跑二期"
    ws.append(["礼包ID", "礼包名称", "价格", "绑定区域"])
    ws.append([1001, "破产礼包·初", 6, "够保跑二期"])
    ws.append([1002, "破产礼包·中", 30, "够保跑二期"])
    ws.append([1003, "破产礼包·高", 98, "够保跑二期"])
    ws2 = wb.create_sheet("礼包配置")
    ws2.append(["礼包ID", "道具", "数量"])
    ws2.append([1001, "金币", 10000])
    ws3 = wb.create_sheet("格式说明")
    ws3["A1"] = "本表仅说明字段含义"
    wb.save(path)
    return path


left_path = build(os.path.join(TMP, "081_破产礼包.xlsx"))
right_path = build(os.path.join(TMP, "081_破产礼包_新.xlsx"))

app = QApplication.instance() or QApplication([])
win = MainWindow()
win.resize(1700, 820)
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
    print(
        f"[{tag}] 左={win.left_sheet_name!r} 右={win.right_sheet_name!r} "
        f"左下拉={win.left_sheet_combo.currentText()!r} "
        f"右下拉={win.right_sheet_combo.currentText()!r} "
        f"底部激活={win.bottom_bar.get_active()!r} "
        f"左数据={'有' if win.left_sheet_data else '无'} "
        f"右数据={'有' if win.right_sheet_data else '无'} "
        f"缺失表={bool(win.diff_result and win.diff_result.sheet_missing)}"
    )


def shot(tag):
    wait()
    win.grab().save(os.path.join(OUT, f"{tag}.png"))
    state(tag)


# 1) 打开两侧文件
from src.core.excel_loader import ExcelLoader

win._on_workbook_loaded("left", left_path, ExcelLoader.load_workbook(left_path), None)
shot("1-仅左侧已打开")
win._on_workbook_loaded("right", right_path, ExcelLoader.load_workbook(right_path), None)
shot("2-打开两侧后自动对齐")

# 2) 手动把左侧切到「礼包配置」——右侧应保持「够保跑二期」
idx = win.left_sheet_combo.findText("礼包配置")
win.left_sheet_combo.setCurrentIndex(idx)
shot("3-手动切左侧后对侧不动")

# 3) 手动把右侧切到「格式说明」——左侧保持「礼包配置」
idx = win.right_sheet_combo.findText("格式说明")
win.right_sheet_combo.setCurrentIndex(idx)
shot("4-两侧停在不同表")

# 4) 点击底部标签「够保跑二期」——两侧同步
win.bottom_bar._tabs["够保跑二期"].click()
shot("5-点底部标签两侧同步")

print("截图目录:", OUT)
