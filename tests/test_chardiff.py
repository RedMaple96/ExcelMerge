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


if __name__ == "__main__":
    unittest.main()
