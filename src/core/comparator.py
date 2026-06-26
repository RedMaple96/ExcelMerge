"""差异比较引擎 — ExcelComparator。

比较两个 SheetData，输出对齐后的行对、差异行/单元格集合与统计信息，
供 GUI（Task 6）可视化与合并器（Task 4）使用。
对应需求：FR-02。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Set, Tuple

from src.core.excel_loader import SheetData


@dataclass
class RowPair:
    """一对对齐后的行。

    row_index 为 0-based 在各自 SheetData.values 中的索引，None 表示该侧不存在此行。

    字段：
    - left_row: 左侧行索引(0-based)，None=左侧无此行
    - right_row: 右侧行索引(0-based)，None=右侧无此行
    - status: "same" | "different" | "left_only" | "right_only"
    - diff_cells: 对于 "different" 行，记录差异列索引(0-based)；其他状态为空列表
    """

    left_row: Optional[int]
    right_row: Optional[int]
    status: str
    diff_cells: List[int]


@dataclass
class DiffResult:
    """比较结果。

    字段：
    - aligned_rows: 对齐后的行列表（按顺序）
    - diff_row_indices: aligned_rows 中 status != "same" 的索引列表
    - diff_cell_set: (aligned_row_index, col_index) 差异单元格集合
    - stats: {"same": int, "different": int, "left_only": int, "right_only": int}
    - max_col: 用于显示的最大列数
    """

    aligned_rows: List[RowPair]
    diff_row_indices: List[int]
    diff_cell_set: Set[Tuple[int, int]]
    stats: dict
    max_col: int


class ExcelComparator:
    """Excel 差异比较器，全部为静态方法。"""

    @staticmethod
    def compare_sheets(
        left: SheetData,
        right: SheetData,
        key_cols: Optional[List[int]] = None,
        ignore_cols: Optional[List[int]] = None,
    ) -> DiffResult:
        """比较两个 SheetData。

        - key_cols: 作为匹配键的列索引列表(0-based)。为空或 None 时按行顺序对齐
          （index 0 对 0、1 对 1 ...）。
        - ignore_cols: 忽略的列索引列表(0-based)，这些列不参与差异判断。
        - 比较范围列数取 max(left.max_col, right.max_col)，缺省单元格视为 ""。
        - 行对齐：用 key_cols 拼接成 key 字符串，按 key 匹配左右行。
          * 仅左侧有的 key -> left_only
          * 仅右侧有的 key -> right_only
          * 两边都有 -> 比较各列(排除 ignore_cols)，全相同 -> same，否则 -> different
        - 公式按字面量比较（values 中已是字符串）。
        """
        if key_cols is None:
            key_cols = []
        if ignore_cols is None:
            ignore_cols = []
        ignore_set: Set[int] = set(ignore_cols)

        # 比较范围：列数取两侧最大值
        max_col: int = max(left.max_col, right.max_col)
        compare_cols: range = range(max_col)

        aligned_rows: List[RowPair] = []
        diff_row_indices: List[int] = []
        diff_cell_set: Set[Tuple[int, int]] = set()
        stats: dict = {"same": 0, "different": 0, "left_only": 0, "right_only": 0}

        n_left = len(left.values)
        n_right = len(right.values)

        def append_pair(
            l_idx: Optional[int],
            r_idx: Optional[int],
            status: str,
            diff_cells: Optional[List[int]] = None,
        ) -> None:
            """追加一个对齐行，并同步更新差异索引/统计/差异单元格集合。"""
            aligned_idx = len(aligned_rows)
            aligned_rows.append(
                RowPair(
                    left_row=l_idx,
                    right_row=r_idx,
                    status=status,
                    diff_cells=list(diff_cells) if diff_cells else [],
                )
            )
            if status == "same":
                stats["same"] += 1
                return
            # 非相同行均记入差异行索引
            diff_row_indices.append(aligned_idx)
            stats[status] += 1
            if status == "different":
                # 差异单元格集合记录所有差异列（ignore_cols 已在比较时跳过，不会出现）
                for c in aligned_rows[aligned_idx].diff_cells:
                    diff_cell_set.add((aligned_idx, c))

        if not key_cols:
            # 默认按行顺序对齐：0 对 0, 1 对 1 ...
            min_len = min(n_left, n_right)
            for i in range(min_len):
                equal, diff_cells = ExcelComparator._rows_equal(
                    left.values[i], right.values[i], compare_cols, ignore_set
                )
                append_pair(
                    i, i, "same" if equal else "different",
                    [] if equal else diff_cells,
                )
            # 左侧多出的行 -> left_only
            for i in range(min_len, n_left):
                append_pair(i, None, "left_only")
            # 右侧多出的行 -> right_only
            for i in range(min_len, n_right):
                append_pair(None, i, "right_only")
        else:
            # 按关键列对齐：构建 key -> 首次出现行索引 的映射（首次出现优先）
            left_key_to_idx: dict = {}
            for i in range(n_left):
                k = ExcelComparator._build_key(left.values[i], key_cols)
                if k not in left_key_to_idx:
                    left_key_to_idx[k] = i
            right_key_to_idx: dict = {}
            for j in range(n_right):
                k = ExcelComparator._build_key(right.values[j], key_cols)
                if k not in right_key_to_idx:
                    right_key_to_idx[k] = j

            consumed_right: Set[int] = set()
            # 先遍历左侧行（保持原始顺序）：能匹配到未消费的右侧行则配对，否则 left_only
            for i in range(n_left):
                k = ExcelComparator._build_key(left.values[i], key_cols)
                j = right_key_to_idx.get(k)
                if j is not None and j not in consumed_right:
                    consumed_right.add(j)
                    equal, diff_cells = ExcelComparator._rows_equal(
                        left.values[i], right.values[j], compare_cols, ignore_set
                    )
                    append_pair(
                        i, j, "same" if equal else "different",
                        [] if equal else diff_cells,
                    )
                else:
                    # 右侧无对应 key，或该右侧行已被消费 -> 左侧独占
                    append_pair(i, None, "left_only")
            # 再遍历右侧行（保持原始顺序）：未被消费的右侧行 -> right_only
            # 对于无重复 key 的常见场景，等价于“key 未在左侧出现 -> right_only”
            for j in range(n_right):
                if j not in consumed_right:
                    append_pair(None, j, "right_only")

        return DiffResult(
            aligned_rows=aligned_rows,
            diff_row_indices=diff_row_indices,
            diff_cell_set=diff_cell_set,
            stats=stats,
            max_col=max_col,
        )

    @staticmethod
    def _build_key(values_row: List[str], key_cols: List[int]) -> str:
        """拼接 key：用 "\\x00" 连接各 key 列的值，越界列视为 ""。"""
        n = len(values_row)
        parts: List[str] = []
        for c in key_cols:
            parts.append(values_row[c] if 0 <= c < n else "")
        return "\x00".join(parts)

    @staticmethod
    def _rows_equal(
        left_row: List[str],
        right_row: List[str],
        compare_cols: range,
        ignore_cols: Set[int],
    ) -> Tuple[bool, List[int]]:
        """比较两行（在 compare_cols 范围内、排除 ignore_cols）。

        缺省单元格视为 ""。返回 (是否全等, 差异列索引列表)。
        """
        diff_cells: List[int] = []
        n_left = len(left_row)
        n_right = len(right_row)
        for c in compare_cols:
            if c in ignore_cols:
                continue
            lv = left_row[c] if c < n_left else ""
            rv = right_row[c] if c < n_right else ""
            if lv != rv:
                diff_cells.append(c)
        return (len(diff_cells) == 0), diff_cells
