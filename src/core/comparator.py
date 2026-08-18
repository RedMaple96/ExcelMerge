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

import bisect
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
    - sheet_missing: 表示当前比较的 Sheet 仅存在于一侧（整表级差异）
    """

    aligned_rows: List[RowPair]
    aligned_cols: List[ColPair]
    diff_row_indices: List[int]
    diff_cell_set: Set[Tuple[int, int]]
    stats: dict
    max_col: int
    sheet_missing: bool = False

    def has_differences(self) -> bool:
        """是否真的存在任何差异（行差异 + 列结构差异）。

        仅统计行 stats 会漏掉"只新增/删除了列"的场景：此时所有行都是
        same，行差异为 0，但 left_only/right_only 列本身就是差异，
        必须在红点等 UI 判定中一并考虑。
        """
        row_diff = (
            self.stats["different"]
            + self.stats["left_only"]
            + self.stats["right_only"]
            > 0
        )
        col_diff = any(
            cp.status in ("left_only", "right_only")
            for cp in self.aligned_cols
        )
        return row_diff or col_diff or self.sheet_missing


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
        - 公式按字面量比较（values 中已是字符串）；若 SheetData 提供
          cached_values（data_only=True 缓存计算值），公式相同但计算值不同
          也判定为差异。
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
                left.cached_values, right.cached_values,
            )
        else:
            ExcelComparator._align_by_key(
                left.values, right.values, key_cols, aligned_cols,
                same_col_info, ignore_set, append_pair,
                left.cached_values, right.cached_values,
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
        """基于自动选择的表头行的列名对齐列。

        策略：
        - 表头行选择：扫描两侧共有行，选"同位置非空且值相等单元格数"最多的
          一行（同分取最早）。真实文件常有标题/摘要行，第一行未必是表头；
          只看第一行会让大量空表头参与 LCS，产生歧义并对错列（如把末尾新增列
          识别成中间新增列）。
        - 列数不同时：用 LCS 对齐（保序地检测列增删/移位）。
        - 列数相同时：在三种候选（位置 1:1、LCS、按列名的全局二分匹配）中
          取匹配数最多的方案，匹配数相同时优先保序方案：
          a) 表头值有差异（如 "Name" vs "Name "）→ 位置 1:1 更稳
          b) 列被移位但顺序大体保留 → LCS 更优
          c) 列被整体重排（如 C,B,A vs A,B,C）→ 全局匹配才能正确对齐，
             避免"同名列被当成不同列比较"造成的错位
        - 如果任一侧无数据行，返回空列表。
        - 如果 LCS 完全没有匹配（列名完全不同），退化为位置 1:1 配对。
        """
        n_left = left.max_col
        n_right = right.max_col

        # 提取列名（自动选择的表头行）
        header_row = ExcelComparator._select_header_row(left, right)
        left_names: List[str] = []
        right_names: List[str] = []
        if left.values:
            lrow = left.values[header_row]
            left_names = [lrow[c] if c < len(lrow) else "" for c in range(n_left)]
        if right.values:
            rrow = right.values[header_row]
            right_names = [rrow[c] if c < len(rrow) else "" for c in range(n_right)]

        # 无标题行时退化为顺序配对
        if not left_names and not right_names:
            return []

        # 列数相同：在位置配对 / LCS / 全局匹配中取最优
        if n_left == n_right and n_left > 0:
            positional_matches = sum(
                1 for c in range(n_left) if left_names[c] == right_names[c]
            )
            raw = ExcelComparator._lcs_backtrace(left_names, right_names)
            lcs_matches = sum(1 for item in raw if item[0] == "match")
            bip_pairs = ExcelComparator._match_columns_bipartite(
                left_names, right_names
            )

            # (匹配数, 优先级, 方案名)；匹配数相同优先保序方案
            candidates = [
                (positional_matches, 2, "positional"),
                (lcs_matches, 1, "lcs"),
                (len(bip_pairs), 0, "bipartite"),
            ]
            candidates.sort(key=lambda item: (-item[0], -item[1]))
            best = candidates[0][2]

            if best == "positional":
                return [
                    ColPair(left_col=c, right_col=c, status="same",
                            label=left_names[c])
                    for c in range(n_left)
                ]
            if best == "lcs":
                return ExcelComparator._build_col_pairs_from_lcs(
                    raw, left_names, right_names
                )
            return ExcelComparator._build_col_pairs_from_matching(
                bip_pairs, left_names, right_names
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
    def _select_header_row(left: SheetData, right: SheetData) -> int:
        """选择用于列对齐的表头行。

        扫描两侧共有的行，返回"同位置非空且值相等的单元格数"最多的一行
        （同分取最早）。表头行通常列名覆盖完整且两侧一致，得分最高；
        标题/摘要行只有部分列有值，得分低；数据行即使值相同也不会比
        更早出现的表头行优先（同分取最早），因此不会误选数据行。
        """
        n_rows = min(len(left.values), len(right.values))
        if n_rows == 0:
            return 0
        n_cols = min(left.max_col, right.max_col)
        best_row = 0
        best_score = -1
        for r in range(n_rows):
            lrow = left.values[r]
            rrow = right.values[r]
            score = 0
            for c in range(n_cols):
                lv = lrow[c] if c < len(lrow) else ""
                rv = rrow[c] if c < len(rrow) else ""
                if lv != "" and lv == rv:
                    score += 1
            if score > best_score:
                best_score = score
                best_row = r
                # 已是理论最高分，后续行不可能更优，提前退出
                if score == n_cols:
                    return best_row
        return best_row

    @staticmethod
    def _match_columns_bipartite(
        left_names: List[str], right_names: List[str]
    ) -> List[Tuple[int, int]]:
        """按列名做全局贪心匹配（允许列重排），返回 (left_idx, right_idx) 列表。

        对每个左侧列名，在尚未消费的右侧列中选同名且位置最接近的列，
        既支持整体重排，又尽量保留原有顺序。表头列名通常不重复，
        此贪心策略即最优匹配；重复列名时也能保持稳定。
        """
        unmatched = set(range(len(right_names)))
        pairs: List[Tuple[int, int]] = []
        for li, name in enumerate(left_names):
            best_ri = -1
            best_dist = None
            for ri in unmatched:
                if right_names[ri] != name:
                    continue
                dist = abs(ri - li)
                if best_dist is None or dist < best_dist:
                    best_ri = ri
                    best_dist = dist
            if best_ri >= 0:
                unmatched.discard(best_ri)
                pairs.append((li, best_ri))
        return pairs

    @staticmethod
    def _build_col_pairs_from_matching(
        pairs: List[Tuple[int, int]],
        left_names: List[str],
        right_names: List[str],
    ) -> List[ColPair]:
        """从全局匹配结果构建 ColPair 列表。

        输出顺序以左侧列序为主；未匹配的右侧列追加到末尾，标记 right_only。
        """
        pair_map = {li: ri for li, ri in pairs}
        matched_right = {ri for _li, ri in pairs}
        aligned_cols: List[ColPair] = []
        for li in range(len(left_names)):
            ri = pair_map.get(li)
            if ri is not None:
                aligned_cols.append(ColPair(
                    left_col=li, right_col=ri, status="same",
                    label=left_names[li],
                ))
            else:
                aligned_cols.append(ColPair(
                    left_col=li, right_col=None, status="left_only",
                    label=left_names[li],
                ))
        for ri in range(len(right_names)):
            if ri not in matched_right:
                aligned_cols.append(ColPair(
                    left_col=None, right_col=ri, status="right_only",
                    label=right_names[ri],
                ))
        return aligned_cols

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

        性能保护：当任一序列长度超过 _LCS_MAX_SIZE 时，退化为简单的
        贪心位置配对，避免 O(n*m) DP 内存/时间爆炸（如16000列场景）。
        """
        n_left = len(left_sigs)
        n_right = len(right_sigs)

        # 性能阈值：超过该大小时不使用标准LCS DP，避免内存爆炸
        _LCS_MAX_SIZE = 500
        if n_left > _LCS_MAX_SIZE or n_right > _LCS_MAX_SIZE:
            return ExcelComparator._greedy_align(left_sigs, right_sigs)

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
    def _greedy_align(left_sigs: List, right_sigs: List) -> List[Tuple]:
        """大序列快速对齐：贪心位置匹配 + 顺序配对。

        - 优先匹配位置相同且签名相等的项
        - 剩余项按顺序配对：短的匹配完后长的标记为独占
        - 时间/空间复杂度 O(n+m)，适合超大序列
        """
        n_left = len(left_sigs)
        n_right = len(right_sigs)
        raw: List[Tuple] = []
        i, j = 0, 0
        while i < n_left and j < n_right:
            if left_sigs[i] == right_sigs[j]:
                raw.append(("match", i, j))
                i += 1
                j += 1
            else:
                # 尝试向右找右侧匹配
                found_j = -1
                search_limit = min(j + 20, n_right)
                for jj in range(j + 1, search_limit):
                    if left_sigs[i] == right_sigs[jj]:
                        found_j = jj
                        break
                if found_j >= 0:
                    for jj in range(j, found_j):
                        raw.append(("right", jj))
                    raw.append(("match", i, found_j))
                    i += 1
                    j = found_j + 1
                    continue
                # 尝试向下找左侧匹配
                found_i = -1
                search_limit = min(i + 20, n_left)
                for ii in range(i + 1, search_limit):
                    if left_sigs[ii] == right_sigs[j]:
                        found_i = ii
                        break
                if found_i >= 0:
                    for ii in range(i, found_i):
                        raw.append(("left", ii))
                    raw.append(("match", found_i, j))
                    i = found_i + 1
                    j += 1
                    continue
                # 都没找到：左侧当前行标记为left_only，推进i
                raw.append(("left", i))
                i += 1
        while i < n_left:
            raw.append(("left", i))
            i += 1
        while j < n_right:
            raw.append(("right", j))
            j += 1
        return raw

    @staticmethod
    def _align_by_content(
        left_values: List[List[str]],
        right_values: List[List[str]],
        aligned_cols: List[ColPair],
        same_col_info: List[Tuple[int, int, int]],
        ignore_set: Set[int],
        append_pair: Callable,
        left_cached: Optional[List[List[str]]] = None,
        right_cached: Optional[List[List[str]]] = None,
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
                    left_values[i], right_values[i], aligned_cols, ignore_set,
                    left_cached[i] if left_cached else None,
                    right_cached[i] if right_cached else None,
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
                    left_values[li], right_values[ri], aligned_cols, ignore_set,
                    left_cached[li] if left_cached else None,
                    right_cached[ri] if right_cached else None,
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
                        left_values[li], right_values[ri], aligned_cols, ignore_set,
                        left_cached[li] if left_cached else None,
                        right_cached[ri] if right_cached else None,
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
        left_cached: Optional[List[List[str]]] = None,
        right_cached: Optional[List[List[str]]] = None,
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

        # 计算每个未消费右侧行的插入位置（其前已配对的右侧行数）。
        # 原实现为 O(unconsumed * pairs) 线性扫描，大文件下退化为 O(n²)；
        # 改为对已配对右侧行索引排序后用 bisect 二分查找，降为 O(n log n)。
        paired_right_sorted = sorted(
            ri for (_li, ri) in pairs if ri is not None
        )
        right_by_pos: dict = defaultdict(list)
        for j in unconsumed_right:
            count = bisect.bisect_left(paired_right_sorted, j)
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
                    left_values[li], right_values[ri], aligned_cols, ignore_set,
                    left_cached[li] if left_cached else None,
                    right_cached[ri] if right_cached else None,
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
        left_cached_row: Optional[List[str]] = None,
        right_cached_row: Optional[List[str]] = None,
    ) -> Tuple[bool, List[int]]:
        """比较两行（在 aligned_cols 范围内、排除 ignore_set）。

        - same 列：比较两侧公式（values）与缓存计算值（cached_values），
          公式或缓存值任一不同即判定为差异。无缓存值时退化为仅比较公式。
        - left_only / right_only 列：不参与比较（跳过），保证"相同内容不错位"

        返回 (是否全等, 差异列的 aligned_col 索引列表)。
        """
        has_cached = left_cached_row is not None and right_cached_row is not None
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
                continue
            # 公式相同：若两侧都有缓存计算值，进一步比较计算值
            # （检测"公式相同但计算结果不同"的差异）
            if has_cached:
                lcv = (
                    left_cached_row[cp.left_col]
                    if cp.left_col is not None and cp.left_col < len(left_cached_row)
                    else ""
                )
                rcv = (
                    right_cached_row[cp.right_col]
                    if cp.right_col is not None and cp.right_col < len(right_cached_row)
                    else ""
                )
                if lcv != rcv:
                    diff_cells.append(idx)
        return (len(diff_cells) == 0), diff_cells
