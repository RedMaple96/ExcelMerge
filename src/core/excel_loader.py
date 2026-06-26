"""Excel 文件读写层 — ExcelLoader。

提供工作簿加载、工作表结构化数据提取、单元格样式/值复制与保存能力，
是后续比较器（Task 3）与合并器（Task 4）的基础设施。
对应需求：FR-01。
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from typing import List, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class SheetData:
    """单个工作表的提取结果。

    字段说明：
    - worksheet: 原始 openpyxl Worksheet 对象，合并器需要它就地修改单元格/样式/合并区域
    - max_row / max_col: 数据区域的最大行/列数
    - values: 字符串化的单元格值二维表（0 索引），合并区域内所有单元格均填左上角值
    - merged_ranges: 合并单元格区域列表，元素为 (min_row, min_col, max_row, max_col)，1 索引
    - header_labels: 列标签，如 ["A","B","C",...]，长度等于 max_col
    """

    worksheet: Worksheet
    max_row: int
    max_col: int
    values: List[List[str]]
    merged_ranges: List[Tuple[int, int, int, int]]
    header_labels: List[str]


class ExcelLoader:
    """Excel 读写工具类，全部为静态方法。"""

    @staticmethod
    def load_workbook(path: str) -> Workbook:
        """加载 .xlsx 工作簿。

        - data_only=False：保留公式字符串而非缓存计算值
        - keep_vba=False：不保留 VBA 宏
        - 注意：openpyxl 的 load_workbook 在当前版本没有 keep_styles 参数，
          样式默认就会被保留，因此此处无需也无法显式传入（PRD 中提到的
          keep_styles=True 并非真实参数，样式保留是默认行为）。
        - 仅支持 .xlsx；遇到 .xls 等旧格式抛出 ValueError。
        """
        if not str(path).lower().endswith(".xlsx"):
            raise ValueError(
                f"仅支持 .xlsx 格式文件，收到: {path}。"
                "如需支持旧版 .xls，请先用 Excel 另存为 .xlsx 后再处理。"
            )
        return load_workbook(path, data_only=False, keep_vba=False)

    @staticmethod
    def get_sheet_names(wb: Workbook) -> List[str]:
        """返回工作簿中所有工作表名称。"""
        return wb.sheetnames

    @staticmethod
    def get_worksheet(wb: Workbook, name: str) -> Worksheet:
        """按名称获取工作表。"""
        return wb[name]

    @staticmethod
    def get_merged_ranges(ws: Worksheet) -> List[Tuple[int, int, int, int]]:
        """返回工作表的合并单元格区域列表，元素为 (min_row, min_col, max_row, max_col)。"""
        ranges: List[Tuple[int, int, int, int]] = []
        for mr in ws.merged_cells.ranges:
            ranges.append((mr.min_row, mr.min_col, mr.max_row, mr.max_col))
        return ranges

    @staticmethod
    def extract_sheet_data(ws: Worksheet) -> SheetData:
        """从 Worksheet 提取结构化数据 SheetData。"""
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0

        merged_ranges = ExcelLoader.get_merged_ranges(ws)

        # 空表保护：行或列为 0 时返回空结构，避免后续循环越界
        if max_row <= 0 or max_col <= 0:
            return SheetData(
                worksheet=ws,
                max_row=max_row,
                max_col=max_col,
                values=[],
                merged_ranges=merged_ranges,
                header_labels=[],
            )

        # 先按单元格原值构建字符串二维表
        values: List[List[str]] = [
            ["" for _ in range(max_col)] for _ in range(max_row)
        ]
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                values[r - 1][c - 1] = ExcelLoader._stringify(ws.cell(r, c).value)

        # 合并单元格：用左上角值填充区域内所有单元格，保证比较时一致
        for mr in ws.merged_cells.ranges:
            min_r, min_c = mr.min_row, mr.min_col
            max_r, max_c = mr.max_row, mr.max_col
            top_value = ExcelLoader._stringify(ws.cell(min_r, min_c).value)
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    # 越界保护：合并区域理论上不会超出 max_row/max_col，此处防御性裁剪
                    if 1 <= r <= max_row and 1 <= c <= max_col:
                        values[r - 1][c - 1] = top_value

        header_labels = [get_column_letter(i) for i in range(1, max_col + 1)]

        return SheetData(
            worksheet=ws,
            max_row=max_row,
            max_col=max_col,
            values=values,
            merged_ranges=merged_ranges,
            header_labels=header_labels,
        )

    @staticmethod
    def _stringify(value) -> str:
        """将单元格值统一转为字符串。

        - None -> ""
        - bool -> "TRUE"/"FALSE"（bool 是 int 子类，必须先于 int 判断）
        - 公式（以 "=" 开头的字符串）原样保留
        - 其它类型直接 str()
        """
        if value is None:
            return ""
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        return str(value)

    @staticmethod
    def save_workbook(wb: Workbook, path: str) -> None:
        """保存工作簿到指定路径。"""
        wb.save(path)

    @staticmethod
    def copy_cell_style(src_cell, tgt_cell) -> None:
        """深拷贝源单元格的样式到目标单元格。

        复制项：font、fill、border、alignment、number_format。
        使用 copy.copy 生成独立样式对象，避免目标修改连锁影响源单元格。
        """
        tgt_cell.font = copy(src_cell.font)
        tgt_cell.fill = copy(src_cell.fill)
        tgt_cell.border = copy(src_cell.border)
        tgt_cell.alignment = copy(src_cell.alignment)
        tgt_cell.number_format = src_cell.number_format

    @staticmethod
    def copy_cell_value(src_cell, tgt_cell) -> None:
        """复制单元格值。

        公式（以 "=" 开头的字符串）原样写入，否则写入原始值（含 None）。
        """
        value = src_cell.value
        if isinstance(value, str) and value.startswith("="):
            tgt_cell.value = value
        else:
            tgt_cell.value = value
