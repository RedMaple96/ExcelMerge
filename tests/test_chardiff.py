"""字符级差异切分单元测试。"""

import unittest

from src.core.chardiff import char_diff_ranges


class CharDiffTest(unittest.TestCase):
    def test_middle_difference(self):
        self.assertEqual(
            char_diff_ranges("12345", "12678"), ([(2, 5)], [(2, 5)])
        )

    def test_tail_append(self):
        self.assertEqual(char_diff_ranges("abc", "abcX"), ([], [(3, 4)]))

    def test_tail_remove(self):
        self.assertEqual(char_diff_ranges("abc", "ab"), ([(2, 3)], []))

    def test_insert_in_middle(self):
        self.assertEqual(char_diff_ranges("abc", "axbc"), ([], [(1, 2)]))

    def test_equal_text(self):
        self.assertEqual(char_diff_ranges("same", "same"), ([], []))

    def test_one_side_empty(self):
        self.assertEqual(char_diff_ranges("", "abc"), ([], [(0, 3)]))
        self.assertEqual(char_diff_ranges("abc", ""), ([(0, 3)], []))

    def test_common_prefix_and_suffix(self):
        # "a...z" 公共前后缀，"X" 与 "YZ" 是差异片段
        self.assertEqual(char_diff_ranges("aXz", "aYZz"), ([(1, 2)], [(1, 3)]))

    # ------------------------------------------------------------------ #
    # 多处差异：中间未变内容不得被连带标红
    # ------------------------------------------------------------------ #
    def test_two_separated_differences_keep_middle_unchanged(self):
        # 多行文本：仅第 2 行与第 5 行不同，中间第 3、4 行必须保持不变
        left = "1.aaa\n2.bbb\n3.ccc\n4.ddd\n5.eee\n"
        right = "1.aaa\n2.bXb\n3.ccc\n4.ddd\n5.eXe\n"

        self.assertEqual(
            char_diff_ranges(left, right), ([(9, 10), (27, 28)], [(9, 10), (27, 28)])
        )

    def test_single_line_two_separated_differences(self):
        # 单行文本同样只标出两处差异，不覆盖中间的 BBBBB
        self.assertEqual(
            char_diff_ranges("AAA X BBBBB Y CCC", "AAA Z BBBBB W CCC"),
            ([(4, 5), (12, 13)], [(4, 5), (12, 13)]),
        )

    def test_inserted_line_marks_only_that_line(self):
        self.assertEqual(
            char_diff_ranges("a\nb\nc\n", "a\nb\nNEW\nc\n"),
            ([], [(4, 8)]),
        )

    def test_deleted_line_marks_only_that_line(self):
        self.assertEqual(
            char_diff_ranges("a\nb\nOLD\nc\n", "a\nb\nc\n"),
            ([(4, 8)], []),
        )

    def test_replaced_block_pairs_lines_by_position(self):
        # 左右同为 3 行、每行都有小改动：逐行给出差异，不漏行也不并成一段
        left = "甲甲甲\n乙乙乙\n丙丙丙\n"
        right = "甲X甲\n乙乙Y\n丙丙丙\n"

        left_ranges, right_ranges = char_diff_ranges(left, right)
        self.assertEqual(left_ranges, [(1, 2), (6, 7)])
        self.assertEqual(right_ranges, [(1, 2), (6, 7)])

    def test_equal_multiline_text(self):
        text = "第一行\n第二行\n第三行"
        self.assertEqual(char_diff_ranges(text, text), ([], []))

    def test_long_single_line_falls_back_to_prefix_suffix(self):
        # 超长单行退化为公共前后缀切分，仍能精确命中中间的单点差异
        prefix = "甲" * 1500
        suffix = "乙" * 1500
        left = prefix + "X" + suffix
        right = prefix + "Y" + suffix

        self.assertEqual(
            char_diff_ranges(left, right), ([(1500, 1501)], [(1500, 1501)])
        )

    def test_line_count_difference_marks_extra_lines(self):
        left = "a\nb\n"
        right = "a\nb\nc\nd\n"

        left_ranges, right_ranges = char_diff_ranges(left, right)
        self.assertEqual(left_ranges, [])
        self.assertEqual(right_ranges, [(4, 8)])


if __name__ == "__main__":
    unittest.main()
