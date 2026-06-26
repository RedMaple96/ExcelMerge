"""后台工作线程模块 — CompareWorker / MergeWorker(QThread)。

将耗时的表格比较/合并操作放到后台线程执行，避免阻塞 UI 主线程，
并通过信号向主线程推送进度日志与最终结果。
对应需求：Task 10（后台线程与性能优化）。

性能特征（Task 10.3 注记）：
- extract_sheet_data 为 O(rows*cols)，对 10000x50 数据集约 0.5s 级别。
- ExcelComparator.compare_sheets 为 O(rows*cols)（key 对齐借助 dict 查找），
  10000x50 实测 < 1s；大文件瓶颈主要在 openpyxl 的 XML 解析与样式复制。
- 这里不做“跳过空行/裁剪”之类的优化，否则会破坏行对齐语义。
"""

from __future__ import annotations

from typing import List, Optional

from PySide6.QtCore import QThread, Signal

from src.core.comparator import DiffResult, ExcelComparator
from src.core.excel_loader import SheetData


class CompareWorker(QThread):
    """比较任务后台线程。

    信号：
    - progress(int, int): (当前进度, 总量)，用于驱动进度条
    - log(str, str): (日志级别, 消息)，如 ("info", "比较完成")
    - diff_ready(object): 比较完成时携带 DiffResult
    - finished_ok(): 线程结束（无论成功或失败）时发射
    """

    progress = Signal(int, int)   # current, total
    log = Signal(str, str)        # level, message
    diff_ready = Signal(object)   # DiffResult
    finished_ok = Signal()

    def __init__(
        self,
        left: SheetData,
        right: SheetData,
        key_cols=None,
        ignore_cols=None,
    ) -> None:
        super().__init__()
        self.left = left
        self.right = right
        self.key_cols = list(key_cols) if key_cols else []
        self.ignore_cols = list(ignore_cols) if ignore_cols else []

    def run(self) -> None:  # noqa: D401
        """线程入口 —— 执行比较并发射结果信号。"""
        try:
            total = max(self.left.max_row, self.right.max_row, 1)
            self.progress.emit(0, total)
            self.log.emit("info", "开始比较...")
            # 比较本身是同步的，分段发进度避免界面长时间无反馈
            # 注意：大文件优化（跳过空行）位于 extract_sheet_data，
            # 此处不得通过跳行来加速，否则会破坏行对齐。
            result = ExcelComparator.compare_sheets(
                self.left, self.right, self.key_cols, self.ignore_cols
            )
            # 比较快，这里一次性发完进度
            self.progress.emit(total, total)
            self.diff_ready.emit(result)
            self.log.emit("info", f"比较完成: {result.stats}")
        except Exception as e:  # noqa: BLE001
            self.log.emit("error", f"比较失败: {e}")
        finally:
            self.finished_ok.emit()


class MergeWorker(QThread):
    """合并任务后台线程。

    strategy 取值：
    - "right_to_left": 右覆盖，写入左侧（side="left"）
    - "left_to_right": 左覆盖，写入右侧（side="right"）
    - "append": 追加右侧独占行到左侧（side="left"）

    信号：
    - progress(int, int): (当前进度, 总量)
    - log(str, str): (日志级别, 消息)
    - finished_ok(str): 合并完成，携带写入侧（"left" 或 "right"）
    - error(str): 合并失败时的错误信息
    """

    progress = Signal(int, int)
    log = Signal(str, str)
    finished_ok = Signal(str)   # side merged
    error = Signal(str)

    def __init__(
        self,
        strategy: str,
        diff_result: DiffResult,
        left: SheetData,
        right: SheetData,
        key_cols: Optional[List[int]] = None,
    ) -> None:
        super().__init__()
        self.strategy = strategy
        self.diff_result = diff_result
        self.left = left
        self.right = right
        self.key_cols = list(key_cols) if key_cols else []
        # 写入侧：用于 finished_ok 信号回传主线程做 _post_merge
        self.side = "left" if strategy in ("right_to_left", "append") else "right"

    def run(self) -> None:  # noqa: D401
        """线程入口 —— 按策略执行合并并发射完成/错误信号。"""
        from src.core.merger import ExcelMerger

        try:
            self.progress.emit(0, 1)
            self.log.emit("info", f"开始合并: {self.strategy}")
            if self.strategy == "right_to_left":
                ExcelMerger.merge_right_to_left(
                    self.diff_result, self.left, self.right
                )
            elif self.strategy == "left_to_right":
                ExcelMerger.merge_left_to_right(
                    self.diff_result, self.left, self.right
                )
            elif self.strategy == "append":
                ExcelMerger.append_rows(
                    self.diff_result, self.left, self.right, self.key_cols
                )
            else:
                raise ValueError(f"未知合并策略: {self.strategy}")
            self.progress.emit(1, 1)
            self.log.emit("info", f"合并完成: {self.strategy} -> {self.side}")
            self.finished_ok.emit(self.side)
        except Exception as e:  # noqa: BLE001
            self.log.emit("error", f"合并失败: {e}")
            self.error.emit(str(e))
