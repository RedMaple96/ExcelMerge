"""验证：长多行单元格只有两处差异时，仅高亮差异片段（含修复前后对比图）。"""

import os
import tempfile

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

import src.gui.main_window as mw
from src.core.chardiff import char_diff_ranges
from src.core.excel_loader import ExcelLoader
from src.gui.diff_highlight_delegate import HIGHLIGHT_RANGES_ROLE
from src.gui.main_window import MainWindow

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shots")
os.makedirs(OUT, exist_ok=True)
TMP = tempfile.mkdtemp()

BASE_LINES = [
    "说明：",
    "1.当幸运值达到100时，下次抽取（单抽或者五连抽）必获得稀有或极品奖品并扣除100幸运值。",
    "2.五连抽将会额外赠送2次机会（共抽7次）。",
    "3.每进行一次转盘幸运值+1。",
    "4.部分奖品会额外获得幸运值。",
    "5.每24小时可以获得一次免费机会。",
    "6.每日可抽奖次数上限为999次，每日0点更新次数。",
    "其它说明：",
    "1.每日一新转盘中每天会更新不同的奖品（每日0点更新）。",
    "2.今日转盘中的常规抽取概率公示如下：",
    "游戏豆x1000：18%",
    "游戏豆x6666：18%",
    "游戏豆x2500：19.5%",
    "白银弹头x2：6.45%",
    "白银弹头x3：8%",
    "黄金弹头x3：0.75%",
    "黄金弹头x3：0.8%",
    "许愿石x50：5%",
    "命运之轮（30天）：7.45%",
    "夏·莲开并蒂（永久）：0.05%",
    "狂杀卡x30：8%",
    "冰冻卡x66：8%",
    "3.今日转盘中的幸运值抽取概率公示如下：",
    "游戏豆x1000：0%",
    "游戏豆x6666：0%",
    "游戏豆x2500：0%",
    "白银弹头x2：0%",
    "白银弹头x3：0%",
    "黄金弹头x3：29.4%",
    "黄金弹头x3：26.45%",
    "许愿石x50：0%",
    "命运之轮（30天）：44.1%",
    "夏·莲开并蒂（永久）：0.05%",
    "狂杀卡x30：0%",
    "冰冻卡x66：0%",
    "温馨提示：文中所有“概率”均是在大样本（大量用户）下的统计数值，与单个玩家的少量测试数据之间可能存在一定的差异。",
]
LEFT_TEXT = "\n".join(BASE_LINES)
RIGHT_LINES = list(BASE_LINES)
RIGHT_LINES[19] = "夏·莲开并蒂（碎片X10）：0.05%"   # 第 1 处差异（第 20 行）
RIGHT_LINES[33] = "夏·莲开并蒂（碎片X10）：0.05%"   # 第 2 处差异（第 34 行）
RIGHT_TEXT = "\n".join(RIGHT_LINES)


def old_char_diff_ranges(left, right):
    """修复前的实现：只做公共前后缀切分。"""
    if left == right:
        return [], []
    max_prefix = min(len(left), len(right))
    prefix = 0
    while prefix < max_prefix and left[prefix] == right[prefix]:
        prefix += 1
    max_suffix = max_prefix - prefix
    suffix = 0
    while (
        suffix < max_suffix
        and left[len(left) - 1 - suffix] == right[len(right) - 1 - suffix]
    ):
        suffix += 1
    left_ranges, right_ranges = [], []
    if len(left) - suffix > prefix:
        left_ranges.append((prefix, len(left) - suffix))
    if len(right) - suffix > prefix:
        right_ranges.append((prefix, len(right) - suffix))
    return left_ranges, right_ranges


def build(path, text):
    wb = Workbook()
    ws = wb.active
    ws.title = "每日一新"
    ws.append(["rulestr2", None])
    ws.append(["rulestr2", None])
    ws.append([None, text])
    for _ in range(6):
        ws.append([None, None])
    wb.save(path)
    return path


left_path = build(os.path.join(TMP, "018_转盘配置.xlsx"), LEFT_TEXT)
right_path = build(os.path.join(TMP, "018_转盘配置_新.xlsx"), RIGHT_TEXT)

app = QApplication.instance() or QApplication([])
win = MainWindow()
win.resize(1500, 780)
win.show()


def wait():
    for _ in range(3000):
        app.processEvents()
        if win._compare_worker is None and win._sheet_summary_worker is None:
            break
    for _ in range(5):
        app.processEvents()


def report(tag):
    item = win.left_table.item(2, 1)
    ranges = item.data(HIGHLIGHT_RANGES_ROLE) or []
    covered = sum(e - s for s, e in ranges)
    print(
        f"[{tag}] 高亮 {len(ranges)} 段，覆盖 {covered}/{len(LEFT_TEXT)} 字符 "
        f"= {covered / len(LEFT_TEXT):.1%}"
    )
    print(f"        区间: {ranges}")


win._on_workbook_loaded("left", left_path, ExcelLoader.load_workbook(left_path), None)
win._on_workbook_loaded("right", right_path, ExcelLoader.load_workbook(right_path), None)
wait()

# 展开该长文本单元格（等价用户双击自适应），让整段内容可见
win._autofit_cell(2, 1, LEFT_TEXT, RIGHT_TEXT)
app.processEvents()

report("修复后")
win.left_table.setCurrentCell(2, 1)
win.right_table.setCurrentCell(2, 1)
app.processEvents()
win.grab().save(os.path.join(OUT, "11-长文本差异-修复后.png"))

# 临时换回旧实现，渲染同样的数据做对比
mw.char_diff_ranges = old_char_diff_ranges
win._render_diffs()
report("修复前")
app.processEvents()
win.grab().save(os.path.join(OUT, "12-长文本差异-修复前.png"))
mw.char_diff_ranges = char_diff_ranges
