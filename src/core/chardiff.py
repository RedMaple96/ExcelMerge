"""单元格文本的字符级差异切分。

供 GUI 的差异高亮 delegate 使用：给定左右两个单元格的文本，计算各自
与对方不同的子串区间，实现"只高亮差异内容"（如 12345 vs 12678 时，
12 不变，345 与 678 分别高亮）。
"""

from __future__ import annotations

from typing import List, Tuple


def char_diff_ranges(
    left: str, right: str
) -> Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]:
    """计算左右文本的差异区间（最长公共前缀 + 最长公共后缀切分）。

    返回 (left_ranges, right_ranges)，每项为 [start, end) 半开区间，
    代表该侧文本中与对方不同的部分。两侧完全相同时返回 ([], [])。

    例如 char_diff_ranges("12345", "12678") -> ([(2, 5)], [(2, 5)])，
    即左侧高亮 "345"、右侧高亮 "678"，公共前缀 "12" 保持不变。
    """
    if left == right:
        return [], []

    # 最长公共前缀
    max_prefix = min(len(left), len(right))
    prefix = 0
    while prefix < max_prefix and left[prefix] == right[prefix]:
        prefix += 1

    # 最长公共后缀（不与前缀重叠）
    max_suffix = max_prefix - prefix
    suffix = 0
    while (
        suffix < max_suffix
        and left[len(left) - 1 - suffix] == right[len(right) - 1 - suffix]
    ):
        suffix += 1

    left_ranges: List[Tuple[int, int]] = []
    right_ranges: List[Tuple[int, int]] = []
    if len(left) - suffix > prefix:
        left_ranges.append((prefix, len(left) - suffix))
    if len(right) - suffix > prefix:
        right_ranges.append((prefix, len(right) - suffix))
    return left_ranges, right_ranges
