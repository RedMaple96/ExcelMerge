"""差异比较引擎 — ExcelComparator。

比较两个 SheetData，输出对齐后的行对、列对、差异行/单元格集合与统计信息，
供 GUI（Task 6）可视化与合并器（Task 4）使用。
对应需求：FR-02。

对齐策略：
- 列对齐：基于标题行（第一行）的列名用 LCS 对齐。两侧都有的列 -> same；
  仅左侧有的列 -> left_only（右侧用虚拟空列填充）；仅右侧有的列 -> right_only。
  行签名与行比较只使用 same 列，保证"相同内容不错位"。
- 行对齐（无 key_cols）：基于行内容签名（same 列的值）计算最长公共子序列（LCS），
  完全相同的行作为锚点配对（same）；LCS 间隙中的行按顺序配对比较
  （same/different），多出的行标记为 left_only/right_only。
- 关键列对齐（有 key_cols）：按 key 匹配配对（首次未消费优先），
  未配对的右侧行（right_only）按其原始位置插入到对齐序列的合适位置。
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Callable, List, Optional, Set, Tuple

from src.core.excel_loader import SheetData


@dataclass
class ColPair:
    """一对对齐后的列。

    字段：
    - left_col: 左侧列索引(0-based)，None=左侧无此列（右侧独占）
    - right_col: 右侧列索引(0-based)，None=右侧无此列（左侧独占）
    - status: "same" | "left_only" | "right_only"
    - label: 列名（用于显示，取自标题行）
    """

    left_col: Optional[int]
    right_col: Optional[int]
    status: str
    label: str


@dataclass
class RowPair:
    """一对对齐后的行。

    row_index 为 0-based 在各自 SheetData.values 中的索引，None 表示该侧不存在此行。

    字段：
    - left_row: 左侧行索引(0-based)，None=左侧无此行
    - right_row: 右侧行索引(0-based)，None=右侧无此行
    - status: "same" | "different" | "left_only" | "right_only"
    - diff_cells: 对于 "different" 行，记录差异列索引(0-based，aligned_col 索引)；其他状态为空列表
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
    - aligned_cols: 对齐后的列列表（按顺序）
    - diff_row_indices: aligned_rows 中 status != "same" 的索引列表
    - diff_cell_set: (aligned_row_index, aligned_col_index) 差异单元格集合
    - stats: {"same": int, "different": int, "left_only": int, "right_only": int}
    - max_col: 用于显示的最大列数（= len(aligned_cols)）
    """

    aligned_rows: List[RowPair]
    aligned_cols: List[ColPair]
    diff_row_indices: List[int]
    diff_cell_set: Set[Tuple[int, int]]
    stats: dict
    max_col: int


class ExcelComparator:
    """Excel 差异比较器，全部为静态方法。"""

    # ------------------------------------------------------------------ #
    # 公共入口
    # ------------------------------------------------------------------ #
    @staticmethod
    def compare_sheets(
        left: SheetData,
        right: SheetData,
        key_cols: Optional[List[int]] = None,
        ignore_cols: Optional[List[int]] = None,
    ) -> DiffResult:
        """比较两个 SheetData。

        - key_cols: 作为匹配键的列索引列表(0-based, aligned_col 索引)。为空或 None 时
          按行顺序对齐（基于内容 LCS 智能对齐）。
        - ignore_cols: 忽略的列索引列表(0-based, aligned_col 索引)，这些列不参与差异判断。
        - 列对齐：基于标题行（第一行）列名 LCS 对齐。多出的列在另一侧用虚拟空列填充。
        - 行对齐：
          * 无 key_cols：基于行内容签名（same 列的值）LCS 对齐。
          * 有 key_cols：按 key 匹配配对。
        - 公式按字面量比较（values 中已是字符串）。
        """
        if key_cols is None:
            key_cols = []
        if ignore_cols is None:
            ignore_cols = []
        ignore_set: Set[int] = set(ignore_cols)

        # 1. 列对齐：基于标题行列名 LCS
        aligned_cols: List[ColPair] = ExcelComparator._align_columns(left, right)
        max_col: int = len(aligned_cols)

        # same 列信息列表：(aligned_idx, left_col, right_col)
        same_col_info: List[Tuple[int, int, int]] = [
            (idx, cp.left_col, cp.right_col)
            for idx, cp in enumerate(aligned_cols)
            if cp.status == "same"
        ]

        aligned_rows: List[RowPair] = []
        diff_row_indices: List[int] = []
        diff_cell_set: Set[Tuple[int, int]] = set()
        stats: dict = {"same": 0, "different": 0, "left_only": 0, "right_only": 0}

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
            diff_row_indices.append(aligned_idx)
            stats[status] += 1
            if status == "different":
                for c in aligned_rows[aligned_idx].diff_cells:
                    diff_cell_set.add((aligned_idx, c))

        # 2. 行对齐
        if not key_cols:
            ExcelComparator._align_by_content(
                left.values, right.values, aligned_cols, same_col_info,
                ignore_set, append_pair,
            )
        else:
            ExcelComparator._align_by_key(
                left.values, right.values, key_cols, aligned_cols,
                same_col_info, ignore_set, append_pair,
            )

        return DiffResult(
            aligned_rows=aligned_rows,
            aligned_cols=aligned_cols,
            diff_row_indices=diff_row_indices,
            diff_cell_set=diff_cell_set,
            stats=stats,
            max_col=max_col,
        )

    # ------------------------------------------------------------------ #
    # 列对齐
    # ------------------------------------------------------------------ #
    @staticmethod
    def _align_columns(
        left: SheetData, right: SheetData
    ) -> List[ColPair]:
        """基于标题行（第一行）的列名用 LCS 对齐列。

        策略：
        - 列数不同时：始终用 LCS 对齐（检测列增删/移位）。
        - 列数相同时：同时计算 LCS 和位置 1:1 配对，取匹配数多的方案。
          这样能正确处理两种情况：
          a) 列数相同但列名有数据差异（如表头单元格值不同）→ 位置 1:1 更优
          b) 列数相同但列被移位（如插入/删除导致后续列平移）→ LCS 更优
        - 如果任一侧无数据行，返回空列表。
        - 如果 LCS 完全没有匹配（列名完全不同），退化为位置 1:1 配对。
        """
        n_left = left.max_col
        n_right = right.max_col

        # 提取列名（第一行）
        left_names: List[str] = []
        right_names: List[str] = []
        if left.values:
            row0 = left.values[0]
            left_names = [row0[c] if c < len(row0) else "" for c in range(n_left)]
        if right.values:
            row0 = right.values[0]
            right_names = [row0[c] if c < len(row0) else "" for c in range(n_right)]

        # 无标题行时退化为顺序配对
        if not left_names and not right_names:
            return []

        # 列数相同：比较 LCS 和位置配对，取匹配多的
        if n_left == n_right and n_left > 0:
            positional_matches = sum(
                1 for c in range(n_left) if left_names[c] == right_names[c]
            )
            raw = ExcelComparator._lcs_backtrace(left_names, right_names)
            lcs_matches = sum(1 for item in raw if item[0] == "match")

            # 位置配对匹配数 >= LCS 匹配数 → 用位置 1:1（处理表头数据差异）
            if positional_matches >= lcs_matches:
                return [
                    ColPair(left_col=c, right_col=c, status="same",
                            label=left_names[c])
                    for c in range(n_left)
                ]
            # LCS 匹配数更多 → 列被移位，用 LCS 结果
            return ExcelComparator._build_col_pairs_from_lcs(
                raw, left_names, right_names
            )

        # 列数不同：始终用 LCS
        raw = ExcelComparator._lcs_backtrace(left_names, right_names)

        # LCS 完全无匹配 → 退化为位置配对
        match_count = sum(1 for item in raw if item[0] == "match")
        if match_count == 0 and n_left > 0 and n_right > 0:
            common = min(n_left, n_right)
            aligned_cols: List[ColPair] = []
            for c in range(common):
                aligned_cols.append(ColPair(
                    left_col=c, right_col=c, status="same",
                    label=left_names[c] if c < len(left_names) else "",
                ))
            for c in range(common, n_left):
                aligned_cols.append(ColPair(
                    left_col=c, right_col=None, status="left_only",
                    label=left_names[c] if c < len(left_names) else "",
                ))
            for c in range(common, n_right):
                aligned_cols.append(ColPair(
                    left_col=None, right_col=c, status="right_only",
                    label=right_names[c] if c < len(right_names) else "",
                ))
            return aligned_cols

        return ExcelComparator._build_col_pairs_from_lcs(
            raw, left_names, right_names
        )

    @staticmethod
    def _build_col_pairs_from_lcs(
        raw: List[Tuple],
        left_names: List[str],
        right_names: List[str],
    ) -> List[ColPair]:
        """从 LCS 回溯结果构建 ColPair 列表。"""
        aligned_cols: List[ColPair] = []
        for item in raw:
            if item[0] == "match":
                _, li, ri = item
                aligned_cols.append(ColPair(
                    left_col=li, right_col=ri, status="same",
                    label=left_names[li],
                ))
            elif item[0] == "left":
                _, li = item
                aligned_cols.append(ColPair(
                    left_col=li, right_col=None, status="left_only",
                    label=left_names[li],
                ))
            else:  # right
                _, ri = item
                aligned_cols.append(ColPair(
                    left_col=None, right_col=ri, status="right_only",
                    label=right_names[ri],
                ))
        return aligned_cols

    # ------------------------------------------------------------------ #
    # 行对齐算法
    # ------------------------------------------------------------------ #
    @staticmethod
    def _row_signature_aligned(
        row: List[str],
        same_col_info: List[Tuple[int, int, int]],
        ignore_set: Set[int],
        side: str,
    ) -> Tuple[str, ...]:
        """计算行签名：在 same 列范围内、排除忽略列后的单元格值元组。

        same_col_info: (aligned_idx, left_col, right_col) 列表
        side: "left" -> 用 left_col 取值；"right" -> 用 right_col 取值

        公式值（以 '=' 开头）统一替换为 '='，避免行号引用差异导致 LCS 错位。
        例如 '=E10' 和 '=E9' 在签名中都变为 '='，使同语义行能正确对齐。
        实际差异判定由 _rows_equal_aligned 负责（保留原始公式值比较）。
        """
        sig: List[str] = []
        for aligned_idx, left_col, right_col in same_col_info:
            if aligned_idx in ignore_set:
                continue
            col = left_col if side == "left" else right_col
            val = row[col] if col is not None and col < len(row) else ""
            # 公式归一化：统一为 '='，避免行号引用差异导致 LCS 错位
            if val and val.startswith("="):
                val = "="
            sig.append(val)
        return tuple(sig)

    @staticmethod
    def _lcs_backtrace(
        left_sigs: List, right_sigs: List
    ) -> List[Tuple]:
        """计算两个签名序列的最长公共子序列并回溯，返回原始对齐序列。

        返回列表元素为：
        - ('match', left_idx, right_idx): 签名相等的配对
        - ('left', left_idx): 仅左侧有的行
        - ('right', right_idx): 仅右侧有的行
        顺序与原始行顺序一致（已 reverse 还原）。
        """
        n_left = len(left_sigs)
        n_right = len(right_sigs)

        dp: List[List[int]] = [[0] * (n_right + 1) for _ in range(n_left + 1)]
        for i in range(1, n_left + 1):
            for j in range(1, n_right + 1):
                if left_sigs[i - 1] == right_sigs[j - 1]:
                    dp[i][j] = dp[i - 1][j - 1] + 1
                else:
                    dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

        raw: List[Tuple] = []
        i, j = n_left, n_right
        while i > 0 and j > 0:
            if left_sigs[i - 1] == right_sigs[j - 1]:
                raw.append(("match", i - 1, j - 1))
                i -= 1
                j -= 1
            elif dp[i - 1][j] >= dp[i][j - 1]:
                raw.append(("left", i - 1))
                i -= 1
            else:
                raw.append(("right", j - 1))
                j -= 1
        while i > 0:
            raw.append(("left", i - 1))
            i -= 1
        while j > 0:
            raw.append(("right", j - 1))
            j -= 1
        raw.reverse()
        return raw

    @staticmethod
    def _align_by_content(
        left_values: List[List[str]],
        right_values: List[List[str]],
        aligned_cols: List[ColPair],
        same_col_info: List[Tuple[int, int, int]],
        ignore_set: Set[int],
        append_pair: Callable,
    ) -> None:
        """默认顺序对齐：基于行内容签名（same 列的值）的 LCS + 间隙配对。

        - 等长快速路径：直接顺序配对比较（保持 different 判定，避免 LCS 开销）。
        - 行数不等：用 LCS 找完全相同行作为锚点（same）；间隙中等长部分配对
          比较（same/different），多出行标记 left_only/right_only。
        """
        n_left = len(left_values)
        n_right = len(right_values)

        # 等长快速路径
        if n_left == n_right:
            for i in range(n_left):
                equal, diff_cells = ExcelComparator._rows_equal_aligned(
                    left_values[i], right_values[i], aligned_cols, ignore_set
                )
                append_pair(
                    i, i, "same" if equal else "different",
                    [] if equal else diff_cells,
                )
            return

        # 行数不等：LCS 对齐
        left_sigs = [
            ExcelComparator._row_signature_aligned(r, same_col_info, ignore_set, "left")
            for r in left_values
        ]
        right_sigs = [
            ExcelComparator._row_signature_aligned(r, same_col_info, ignore_set, "right")
            for r in right_values
        ]
        raw = ExcelComparator._lcs_backtrace(left_sigs, right_sigs)

        k = 0
        while k < len(raw):
            if raw[k][0] == "match":
                _, li, ri = raw[k]
                # 签名匹配后仍需验证实际值（公式归一化可能导致签名相同但值不同）
                equal, diff_cells = ExcelComparator._rows_equal_aligned(
                    left_values[li], right_values[ri], aligned_cols, ignore_set
                )
                append_pair(
                    li, ri, "same" if equal else "different",
                    [] if equal else diff_cells,
                )
                k += 1
            else:
                left_block: List[int] = []
                right_block: List[int] = []
                while k < len(raw) and raw[k][0] != "match":
                    if raw[k][0] == "left":
                        left_block.append(raw[k][1])
                    else:
                        right_block.append(raw[k][1])
                    k += 1
                common = min(len(left_block), len(right_block))
                for m in range(common):
                    li = left_block[m]
                    ri = right_block[m]
                    equal, diff_cells = ExcelComparator._rows_equal_aligned(
                        left_values[li], right_values[ri], aligned_cols, ignore_set
                    )
                    append_pair(
                        li, ri, "same" if equal else "different",
                        [] if equal else diff_cells,
                    )
                for m in range(common, len(left_block)):
                    append_pair(left_block[m], None, "left_only")
                for m in range(common, len(right_block)):
                    append_pair(None, right_block[m], "right_only")

    @staticmethod
    def _align_by_key(
        left_values: List[List[str]],
        right_values: List[List[str]],
        key_cols: List[int],
        aligned_cols: List[ColPair],
        same_col_info: List[Tuple[int, int, int]],
        ignore_set: Set[int],
        append_pair: Callable,
    ) -> None:
        """关键列对齐：按 key 配对（首次未消费优先）+ right_only 按原始位置插入。

        key_cols 是 aligned_col 索引。
        """
        n_left = len(left_values)
        n_right = len(right_values)

        left_keys = [
            ExcelComparator._build_key_aligned(
                left_values[i], key_cols, aligned_cols, "left"
            )
            for i in range(n_left)
        ]
        right_keys = [
            ExcelComparator._build_key_aligned(
                right_values[j], key_cols, aligned_cols, "right"
            )
            for j in range(n_right)
        ]

        consumed_right: Set[int] = set()
        pairs: List[Tuple[int, Optional[int]]] = []
        right_key_first: dict = {}
        for j in range(n_right):
            k = right_keys[j]
            if k not in right_key_first:
                right_key_first[k] = j

        for i in range(n_left):
            k = left_keys[i]
            j = right_key_first.get(k)
            if j is not None and j not in consumed_right:
                consumed_right.add(j)
                pairs.append((i, j))
                nxt = j + 1
                while nxt < n_right and (
                    right_keys[nxt] != k or nxt in consumed_right
                ):
                    nxt += 1
                if nxt < n_right and right_keys[nxt] == k:
                    right_key_first[k] = nxt
                else:
                    right_key_first.pop(k, None)
            else:
                pairs.append((i, None))

        unconsumed_right = [j for j in range(n_right) if j not in consumed_right]

        right_by_pos: dict = defaultdict(list)
        for j in unconsumed_right:
            count = sum(
                1 for (_li, ri) in pairs if ri is not None and ri < j
            )
            right_by_pos[count].append(j)
        for v in right_by_pos.values():
            v.sort()

        for idx, (li, ri) in enumerate(pairs):
            for j in right_by_pos.get(idx, []):
                append_pair(None, j, "right_only")
            if ri is None:
                append_pair(li, None, "left_only")
            else:
                equal, diff_cells = ExcelComparator._rows_equal_aligned(
                    left_values[li], right_values[ri], aligned_cols, ignore_set
                )
                append_pair(
                    li, ri, "same" if equal else "different",
                    [] if equal else diff_cells,
                )
        for j in right_by_pos.get(len(pairs), []):
            append_pair(None, j, "right_only")

    @staticmethod
    def _build_key_aligned(
        values_row: List[str],
        key_cols: List[int],
        aligned_cols: List[ColPair],
        side: str,
    ) -> str:
        """拼接 key：用 "\\x00" 连接各 key 列的值。

        key_cols 是 aligned_col 索引。
        """
        parts: List[str] = []
        for aligned_idx in key_cols:
            if aligned_idx >= len(aligned_cols):
                parts.append("")
                continue
            cp = aligned_cols[aligned_idx]
            col = cp.left_col if side == "left" else cp.right_col
            parts.append(
                values_row[col] if col is not None and col < len(values_row) else ""
            )
        return "\x00".join(parts)

    @staticmethod
    def _rows_equal_aligned(
        left_row: List[str],
        right_row: List[str],
        aligned_cols: List[ColPair],
        ignore_set: Set[int],
    ) -> Tuple[bool, List[int]]:
        """比较两行（在 aligned_cols 范围内、排除 ignore_set）。

        - same 列：比较两侧值
        - left_only / right_only 列：不参与比较（跳过），保证"相同内容不错位"

        返回 (是否全等, 差异列的 aligned_col 索引列表)。
        """
        diff_cells: List[int] = []
        for idx, cp in enumerate(aligned_cols):
            if idx in ignore_set:
                continue
            if cp.status != "same":
                # 独占列不参与 same/different 判定
                continue
            lv = (
                left_row[cp.left_col]
                if cp.left_col is not None and cp.left_col < len(left_row)
                else ""
            )
            rv = (
                right_row[cp.right_col]
                if cp.right_col is not None and cp.right_col < len(right_row)
                else ""
            )
            if lv != rv:
                diff_cells.append(idx)
        return (len(diff_cells) == 0), diff_cells
