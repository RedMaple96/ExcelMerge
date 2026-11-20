"""单元格文本的字符级差异切分。

供 GUI 的差异高亮 delegate 使用：给定左右两个单元格的文本，计算各自
与对方不同的子串区间，实现"只高亮差异内容"（如 12345 vs 12678 时，
12 不变，345 与 678 分别高亮）。

切分策略是"行级对齐 + 行内字符级切分"两级：

- 先按行（保留行尾）用 SequenceMatcher 对齐，只有发生变化的那几行会
  进入行内切分；
- 变化的行再按位置配对做行内字符级切分，因此同一单元格里有多处互不
  相邻的差异时，中间没变的内容不会被连带标红（旧实现只做公共前后缀
  切分，会把首尾差异之间的整段内容都判为差异）。

行级对齐既贴合配置表"多行文本单元格"的常态，也比整段文本做字符级
SequenceMatcher 快两个数量级（避免在渲染差异时卡住 UI）。
"""

from __future__ import annotations

from difflib import SequenceMatcher
from typing import List, Tuple

# 单行文本超过该长度时退化为"公共前后缀"切分：字符级 SequenceMatcher
# 最坏复杂度为 O(n*m)，超长行直接跑会明显变慢，而超长单行文本很少见。
_LINE_CHAR_DIFF_LIMIT = 2000


def char_diff_ranges(
    left: str, right: str
) -> Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]:
    """计算左右文本的差异区间（多段，行级对齐 + 行内字符级切分）。

    返回 (left_ranges, right_ranges)，每项为 [start, end) 半开区间，
    代表该侧文本中与对方不同的部分。两侧完全相同时返回 ([], [])。

    例如 char_diff_ranges("12345", "12678") -> ([(2, 5)], [(2, 5)])，
    即左侧高亮 "345"、右侧高亮 "678"，公共前缀 "12" 保持不变。
    """
    if left == right:
        return [], []

    left_lines = _split_lines(left)
    right_lines = _split_lines(right)
    if len(left_lines) == 1 and len(right_lines) == 1:
        return _line_pair_ranges(left, right)

    left_offsets = _line_offsets(left_lines)
    right_offsets = _line_offsets(right_lines)
    left_ranges: List[Tuple[int, int]] = []
    right_ranges: List[Tuple[int, int]] = []

    matcher = SequenceMatcher(None, left_lines, right_lines, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        # replace：左右行数可能不等，按位置配对，多出来的行整行标记；
        # insert/delete：整块都只存在于一侧，整行标记。
        for k in range(max(i2 - i1, j2 - j1)):
            li, rj = i1 + k, j1 + k
            if li < i2 and rj < j2:
                pair_left, pair_right = _line_pair_ranges(
                    left_lines[li], right_lines[rj]
                )
                left_ranges += [
                    (left_offsets[li] + s, left_offsets[li] + e) for s, e in pair_left
                ]
                right_ranges += [
                    (right_offsets[rj] + s, right_offsets[rj] + e)
                    for s, e in pair_right
                ]
            elif li < i2:
                left_ranges.append(
                    (left_offsets[li], left_offsets[li] + len(left_lines[li]))
                )
            else:
                right_ranges.append(
                    (right_offsets[rj], right_offsets[rj] + len(right_lines[rj]))
                )

    return _merge_ranges(left_ranges), _merge_ranges(right_ranges)


def _split_lines(text: str) -> List[str]:
    """按行切分并保留行尾（含换行符），保证拼接后与原文完全一致。"""
    lines = text.splitlines(keepends=True)
    return lines if lines else [""]


def _line_offsets(lines: List[str]) -> List[int]:
    """每行首字符在整段文本中的偏移。"""
    offsets: List[int] = []
    pos = 0
    for line in lines:
        offsets.append(pos)
        pos += len(line)
    return offsets


def _line_pair_ranges(
    left: str, right: str
) -> Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]:
    """单行（或单行文本）的行内字符级差异区间。"""
    if left == right:
        return [], []
    if len(left) + len(right) > _LINE_CHAR_DIFF_LIMIT:
        return _prefix_suffix_ranges(left, right)

    left_ranges: List[Tuple[int, int]] = []
    right_ranges: List[Tuple[int, int]] = []
    matcher = SequenceMatcher(None, left, right, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag in ("replace", "delete") and i2 > i1:
            left_ranges.append((i1, i2))
        if tag in ("replace", "insert") and j2 > j1:
            right_ranges.append((j1, j2))
    return _merge_ranges(left_ranges), _merge_ranges(right_ranges)


def _prefix_suffix_ranges(
    left: str, right: str
) -> Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]:
    """超长单行的降级切分：最大公共前缀 + 最大公共后缀。"""
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

    left_ranges: List[Tuple[int, int]] = []
    right_ranges: List[Tuple[int, int]] = []
    if len(left) - suffix > prefix:
        left_ranges.append((prefix, len(left) - suffix))
    if len(right) - suffix > prefix:
        right_ranges.append((prefix, len(right) - suffix))
    return left_ranges, right_ranges


def _merge_ranges(ranges: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    """合并重叠或首尾相接的区间，并按位置排序。"""
    merged: List[Tuple[int, int]] = []
    for start, end in sorted(ranges):
        if end <= start:
            continue
        if merged and start <= merged[-1][1]:
            prev_start, prev_end = merged[-1]
            merged[-1] = (prev_start, max(prev_end, end))
        else:
            merged.append((start, end))
    return merged
