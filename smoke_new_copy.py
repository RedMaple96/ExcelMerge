"""冒烟验证：sheet级覆盖复制 + 文件级整体替换。"""
import os
import sys
import tempfile

os.environ["QT_QPA_PLATFORM"] = "offscreen"
from openpyxl import Workbook, load_workbook
from PySide6.QtWidgets import QApplication, QMessageBox
from unittest.mock import patch

from src.gui.main_window import MainWindow

tmp = tempfile.mkdtemp()
left_path = os.path.join(tmp, "left.xlsx")
right_path = os.path.join(tmp, "right.xlsx")

wb = Workbook()
ws = wb.active
ws.title = "S1"
ws["A1"] = "L-S1-A1"
ws["B2"] = 100
wb.create_sheet("S2")["A1"] = "L-S2"
wb.save(left_path)

wb = Workbook()
ws = wb.active
ws.title = "S1"
ws["A1"] = "R-S1-A1-CHANGED"
ws["C3"] = 999
wb.create_sheet("S3")["A1"] = "R-S3"
wb.save(right_path)

app = QApplication(sys.argv)
with patch("src.utils.license.check_license", return_value={"ok": True}):
    win = MainWindow()
    win.load_file("left", left_path)
    win.load_file("right", right_path)

    def _wait_load():
        import time
        deadline = time.time() + 30
        while time.time() < deadline and (win.left_wb is None or win.right_wb is None):
            app.processEvents()
        assert win.left_wb is not None and win.right_wb is not None, "加载超时"

    _wait_load()

# --- 场景1：sheet级覆盖，右侧S1覆盖左侧S1 ---
# QMessageBox.question 自动回答 Yes
with patch("PySide6.QtWidgets.QMessageBox.question", return_value=QMessageBox.Yes):
    win._replace_whole_sheet("right")

lw = win.left_wb
assert "S1" in lw.sheetnames and "S2" in lw.sheetnames, lw.sheetnames
assert lw["S1"]["A1"].value == "R-S1-A1-CHANGED", lw["S1"]["A1"].value
assert lw["S1"]["C3"].value == 999
assert lw.sheetnames == ["S1", "S2"], lw.sheetnames  # S2 保留、顺序不变
assert win._dirty["left"] is True
print("场景1 sheet级覆盖: PASS")

# --- 场景2：文件级替换，右侧整体覆盖左侧（合并模型内对象） ---
with patch("PySide6.QtWidgets.QMessageBox.question", return_value=QMessageBox.Yes):
    win._replace_whole_file("right")

lw = win.left_wb
assert lw.sheetnames == ["S1", "S3"], lw.sheetnames  # S2 被替换掉
assert lw["S1"]["A1"].value == "R-S1-A1-CHANGED"
assert lw["S3"]["A1"].value == "R-S3"
assert win.left_path == left_path  # 路径不变
assert win._dirty["left"] is True
print("场景2 文件级替换(内存模型): PASS")

# --- 场景3：文件级替换后保存写回磁盘 ---
win._save_side("left", quiet=True); win._save_side("right", quiet=True)
rb = load_workbook(left_path)
assert rb.sheetnames == ["S1", "S3"], rb.sheetnames
assert rb["S1"]["A1"].value == "R-S1-A1-CHANGED"
assert rb["S3"]["A1"].value == "R-S3"
print("场景3 保存写回磁盘: PASS")

# --- 场景4：反向，左侧覆盖右侧（右侧路径写回） ---
with patch("PySide6.QtWidgets.QMessageBox.question", return_value=QMessageBox.Yes):
    win._replace_whole_file("left")
win._save_side("left", quiet=True); win._save_side("right", quiet=True)
rb2 = load_workbook(right_path)
assert rb2.sheetnames == ["S1", "S3"], rb2.sheetnames
print("场景4 反向文件级替换+保存: PASS")

# --- 场景5：No 取消不修改 ---
before = win.left_wb.sheetnames[:]
with patch("PySide6.QtWidgets.QMessageBox.question", return_value=QMessageBox.No):
    win._replace_whole_file("right")
assert win.left_wb.sheetnames == before
print("场景5 取消确认不修改: PASS")

print("ALL PASS")
