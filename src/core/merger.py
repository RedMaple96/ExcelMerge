"""合并策略执行器 — ExcelMerger。

基于 DiffResult 在 openpyxl Worksheet 上就地执行各类合并策略：
右覆盖、左覆盖、追加差异行、逐项手动合并。
对应需求：FR-04。

行号映射约定：
- SheetData.values 为 0 索引；openpyxl cell 的 row/column 为 1 索引。
- 因此 0-based 行索引 -> worksheet 行号 = idx + 1；0-based 列索引 -> 列号 = idx + 1。
"""

from __future__ import annotations

from typing import List, Optional, Tuple

from src.core.comparator import ColPair, DiffResult, RowPair
from src.core.excel_loader import ExcelLoader, SheetData


class ExcelMerger:
    """Excel 合并执行器，全部为静态方法。

    所有方法均直接修改传入 SheetData 中的 worksheet 对象（就地修改），
    调用方负责后续保存工作簿。
    """

    @staticmethod
    def merge_right_to_left(
        diff_result: DiffResult, left: SheetData, right: SheetData
    ) -> None:
        """右覆盖：以右侧为准，将右侧与左侧不同的单元格值及样式复制到左侧。

        - "different" 行：按 diff_cells（aligned_col 索引）逐列把右侧单元格复制到左侧。
          使用 aligned_cols 映射到实际列号。
        - "right_only" 行：在左侧对应位置插入新行并复制右侧整行（保持对齐位置，
          而非追加到末尾）。多个插入按从后往前执行，避免位置偏移。
        """
        aligned_cols = diff_result.aligned_cols
        insert_max_col = max(left.max_col, right.max_col)

        # 收集需要插入的 (插入位置0-based, 源行0-based)，从后往前执行
        inserts: List[Tuple[int, int]] = []
        for i, pair in enumerate(diff_result.aligned_rows):
            if pair.status == "different":
                left_ws_row = pair.left_row + 1
                right_ws_row = pair.right_row + 1
                for aligned_col in pair.diff_cells:
                    if aligned_col >= len(aligned_cols):
                        continue
                    cp = aligned_cols[aligned_col]
                    if cp.status != "same":
                        continue
                    src_cell = right.worksheet.cell(
                        row=right_ws_row, column=cp.right_col + 1
                    )
                    tgt_cell = left.worksheet.cell(
                        row=left_ws_row, column=cp.left_col + 1
                    )
                    ExcelLoader.copy_cell_value(src_cell, tgt_cell)
                    ExcelLoader.copy_cell_style(src_cell, tgt_cell)
            elif pair.status == "right_only":
                insert_idx = sum(
                    1 for p in diff_result.aligned_rows[:i] if p.left_row is not None
                )
                inserts.append((insert_idx, pair.right_row))

        for insert_idx, src_row in sorted(inserts, key=lambda x: x[0], reverse=True):
            ExcelMerger.insert_row_to_other(
                right, src_row, left, insert_idx, insert_max_col
            )

    @staticmethod
    def merge_left_to_right(
        diff_result: DiffResult, left: SheetData, right: SheetData
    ) -> None:
        """左覆盖：以左侧为准，镜像右覆盖操作。

        - "different" 行：按 diff_cells（aligned_col 索引）逐列把左侧单元格复制到右侧。
          使用 aligned_cols 映射到实际列号。
        - "left_only" 行：在右侧对应位置插入新行并复制左侧整行（保持对齐位置，
          而非追加到末尾）。多个插入按从后往前执行，避免位置偏移。
        """
        aligned_cols = diff_result.aligned_cols
        insert_max_col = max(left.max_col, right.max_col)

        inserts: List[Tuple[int, int]] = []
        for i, pair in enumerate(diff_result.aligned_rows):
            if pair.status == "different":
                left_ws_row = pair.left_row + 1
                right_ws_row = pair.right_row + 1
                for aligned_col in pair.diff_cells:
                    if aligned_col >= len(aligned_cols):
                        continue
                    cp = aligned_cols[aligned_col]
                    if cp.status != "same":
                        continue
                    src_cell = left.worksheet.cell(
                        row=left_ws_row, column=cp.left_col + 1
                    )
                    tgt_cell = right.worksheet.cell(
                        row=right_ws_row, column=cp.right_col + 1
                    )
                    ExcelLoader.copy_cell_value(src_cell, tgt_cell)
                    ExcelLoader.copy_cell_style(src_cell, tgt_cell)
            elif pair.status == "left_only":
                insert_idx = sum(
                    1 for p in diff_result.aligned_rows[:i] if p.right_row is not None
                )
                inserts.append((insert_idx, pair.left_row))

        for insert_idx, src_row in sorted(inserts, key=lambda x: x[0], reverse=True):
            ExcelMerger.insert_row_to_other(
                left, src_row, right, insert_idx, insert_max_col
            )

    @staticmethod
    def append_rows(
        diff_result: DiffResult,
        left: SheetData,
        right: SheetData,
        key_cols: Optional[List[int]] = None,
    ) -> None:
        """追加差异行：仅把 status=="right_only" 的行整行追加到左侧末尾。

        - 不处理 "different" 行，保持左侧原值不变。
        - key_cols 参数保留以匹配调用方签名，本方法不依赖它。
        """
        append_max_col = max(left.max_col, right.max_col)
        for pair in diff_result.aligned_rows:
            if pair.status == "right_only":
                ExcelMerger._append_row(right, pair.right_row, left, append_max_col)

    @staticmethod
    def copy_row_to_other(
        source: SheetData,
        source_row_idx: int,
        target: SheetData,
        target_row_idx: int,
        max_col: int,
    ) -> None:
        """逐项手动合并：把 source 指定行(0-based)整行(列 0..max_col-1)值与样式复制到 target 指定行(0-based)。

        target_row_idx 必须是已存在的行；如需追加到末尾，调用方应先扩展 target worksheet。
        """
        src_ws_row = source_row_idx + 1
        tgt_ws_row = target_row_idx + 1
        for col in range(max_col):
            ws_col = col + 1
            src_cell = source.worksheet.cell(row=src_ws_row, column=ws_col)
            tgt_cell = target.worksheet.cell(row=tgt_ws_row, column=ws_col)
            ExcelLoader.copy_cell_value(src_cell, tgt_cell)
            ExcelLoader.copy_cell_style(src_cell, tgt_cell)

    @staticmethod
    def insert_row_to_other(
        source: SheetData,
        source_row_idx: int,
        target: SheetData,
        target_insert_row_idx: int,
        max_col: int,
    ) -> int:
        """在 target 指定位置插入新行，并从 source 复制整行值与样式。

        用于把 left_only/right_only 行复制到对侧对应位置（在虚拟空行处新增行），
        区别于 copy_row_to_other（覆盖已存在行）与 _append_row（追加到末尾）。

        - target_insert_row_idx: 0-based，新行插入后位于该索引位置；
          原该位置及之后的行整体后移（openpyxl insert_rows）。
        - 逐列(0..max_col-1)复制值与样式。
        - 合并单元格重建：仅处理以源行为左上角的合并区域，在 target 中以插入行
          为新 min_row、保持列范围与行跨度不变重建。
        - 返回插入行号(1-based)。
        """
        insert_ws_row = target_insert_row_idx + 1  # 0-based -> 1-based
        target.worksheet.insert_rows(insert_ws_row)

        src_ws_row = source_row_idx + 1
        for col in range(max_col):
            ws_col = col + 1
            src_cell = source.worksheet.cell(row=src_ws_row, column=ws_col)
            tgt_cell = target.worksheet.cell(row=insert_ws_row, column=ws_col)
            ExcelLoader.copy_cell_value(src_cell, tgt_cell)
            ExcelLoader.copy_cell_style(src_cell, tgt_cell)

        # 重建以源行为左上角的合并区域
        for min_r, min_c, max_r, max_c in ExcelMerger._find_merged_ranges_for_row(
            source, src_ws_row
        ):
            new_min_row = insert_ws_row
            new_max_row = insert_ws_row + (max_r - min_r)
            try:
                target.worksheet.merge_cells(
                    start_row=new_min_row,
                    start_column=min_c,
                    end_row=new_max_row,
                    end_column=max_c,
                )
            except Exception:
                # target 已有重叠合并区域时跳过，避免中断整体流程
                continue

        return insert_ws_row

    @staticmethod
    def copy_single_cell(
        source: SheetData,
        source_row_idx: int,
        col_idx: int,
        target: SheetData,
        target_row_idx: int,
    ) -> None:
        """逐项手动合并：把 source 单个单元格(0-based row, 0-based col)值与样式复制到 target 对应位置。"""
        src_cell = source.worksheet.cell(row=source_row_idx + 1, column=col_idx + 1)
        tgt_cell = target.worksheet.cell(row=target_row_idx + 1, column=col_idx + 1)
        ExcelLoader.copy_cell_value(src_cell, tgt_cell)
        ExcelLoader.copy_cell_style(src_cell, tgt_cell)

    @staticmethod
    def _append_row(
        source: SheetData, source_row_idx: int, target: SheetData, max_col: int
    ) -> int:
        """内部：把 source 指定行(0-based)追加到 target worksheet 末尾。

        - 追加位置 = target.worksheet.max_row + 1。
        - 逐列(0..max_col-1)复制值与样式。
        - 合并单元格重建：仅处理以源行为左上角(min_row==source_row_idx+1)的合并区域，
          在 target 中以 append_row 为新 min_row、保持 min_col/max_col 与行跨度不变重建。
        - 返回追加行号(1-based)。
        """
        append_row = target.worksheet.max_row + 1
        src_ws_row = source_row_idx + 1

        # 逐列复制值与样式
        for col in range(max_col):
            ws_col = col + 1
            src_cell = source.worksheet.cell(row=src_ws_row, column=ws_col)
            tgt_cell = target.worksheet.cell(row=append_row, column=ws_col)
            ExcelLoader.copy_cell_value(src_cell, tgt_cell)
            ExcelLoader.copy_cell_style(src_cell, tgt_cell)

        # 重建以源行为左上角的合并区域
        for min_r, min_c, max_r, max_c in ExcelMerger._find_merged_ranges_for_row(
            source, src_ws_row
        ):
            new_min_row = append_row
            # 保持行跨度(max_r - min_r)不变
            new_max_row = append_row + (max_r - min_r)
            try:
                target.worksheet.merge_cells(
                    start_row=new_min_row,
                    start_column=min_c,
                    end_row=new_max_row,
                    end_column=max_c,
                )
            except Exception:
                # target 已有重叠合并区域时跳过，避免中断整体合并流程
                continue

        return append_row

    @staticmethod
    def _find_merged_ranges_for_row(
        sheet_data: SheetData, row_1based: int
    ) -> List[Tuple[int, int, int, int]]:
        """返回 sheet_data.merged_ranges 中 min_row == row_1based 的所有 range。

        即以该行为左上角的合并区域。
        """
        return [
            (min_r, min_c, max_r, max_c)
            for (min_r, min_c, max_r, max_c) in sheet_data.merged_ranges
            if min_r == row_1based
        ]
