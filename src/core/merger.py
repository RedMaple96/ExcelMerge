"""合并策略执行器 — ExcelMerger。

基于 DiffResult 在 openpyxl Worksheet 上就地执行各类合并策略：
右覆盖、左覆盖、追加差异行、逐项手动合并。
对应需求：FR-04。

行号映射约定：
- SheetData.values 为 0 索引；openpyxl cell 的 row/column 为 1 索引。
- 因此 0-based 行索引 -> worksheet 行号 = idx + 1；0-based 列索引 -> 列号 = idx + 1。
"""

from __future__ import annotations

from copy import copy
from typing import List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from src.core.comparator import ColPair, DiffResult, RowPair
from src.core.excel_loader import ExcelLoader, SheetData


class ExcelMerger:
    """Excel 合并执行器，全部为静态方法。

    所有方法均直接修改传入 SheetData 中的 worksheet 对象（就地修改），
    调用方负责后续保存工作簿。
    """

    @staticmethod
    def copy_sheet_to_workbook(
        source: SheetData, target_wb: Workbook, sheet_name: Optional[str] = None
    ) -> Worksheet:
        """跨工作簿整表复制一个工作表。"""
        title = sheet_name or source.worksheet.title
        if title in target_wb.sheetnames:
            raise ValueError(f"目标工作簿已存在工作表: {title}")

        source_ws = source.worksheet
        target_ws = target_wb.create_sheet(title)
        cached_values = source.cached_values or []

        for row, col, src_cell in ExcelLoader._iter_cells(source_ws):
            tgt_cell = target_ws.cell(row=row, column=col)
            cached_value = None
            if row <= len(cached_values):
                source_row = cached_values[row - 1]
                if col <= len(source_row):
                    cached_value = source_row[col - 1]
            ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached_value)
            ExcelLoader.copy_cell_style(src_cell, tgt_cell)

        for merged_range in source_ws.merged_cells.ranges:
            target_ws.merge_cells(str(merged_range))

        for row_key, row_dim in source_ws.row_dimensions.items():
            target_dim = target_ws.row_dimensions[row_key]
            target_dim.height = row_dim.height
            target_dim.hidden = row_dim.hidden
            target_dim.outlineLevel = row_dim.outlineLevel

        for col_key, col_dim in source_ws.column_dimensions.items():
            target_dim = target_ws.column_dimensions[col_key]
            target_dim.width = col_dim.width
            target_dim.hidden = col_dim.hidden
            target_dim.outlineLevel = col_dim.outlineLevel

        target_ws.freeze_panes = source_ws.freeze_panes
        target_ws.sheet_format = copy(source_ws.sheet_format)
        target_ws.sheet_properties = copy(source_ws.sheet_properties)
        target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines
        if source_ws.auto_filter.ref:
            target_ws.auto_filter.ref = source_ws.auto_filter.ref
        return target_ws

    # ------------------------------------------------------------------ #
    # 合并单元格维护（openpyxl 的 insert/delete_rows/cols 不会同步移动
    # merged ranges，必须按 Excel 语义手动调整，否则会错位甚至丢数据）
    # ------------------------------------------------------------------ #
    @staticmethod
    def _rebuild_merged_ranges(
        ws: Worksheet, specs: List[Tuple[int, int, int, int]]
    ) -> None:
        """清空并按给定 (min_row, min_col, max_row, max_col) 列表重建合并区域。"""
        ws.merged_cells.ranges.clear()
        for min_r, min_c, max_r, max_c in specs:
            if min_r > max_r or min_c > max_c:
                continue
            try:
                ws.merge_cells(
                    start_row=min_r,
                    start_column=min_c,
                    end_row=max_r,
                    end_column=max_c,
                )
            except Exception:
                # 重叠区域跳过，避免中断整体流程
                continue

    @staticmethod
    def _adjust_merged_ranges_for_row_insert(
        ws: Worksheet, insert_row: int
    ) -> None:
        """行插入前调整合并区域（Excel 语义）。

        - 完全在插入行上方：不变；
        - 完全在插入行及其下方：整体下移一行；
        - 跨越插入行：向下扩展一行。
        """
        new_specs: List[Tuple[int, int, int, int]] = []
        for mr in ws.merged_cells.ranges:
            if mr.max_row < insert_row:
                new_specs.append(
                    (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
                )
            elif mr.min_row >= insert_row:
                new_specs.append(
                    (mr.min_row + 1, mr.min_col, mr.max_row + 1, mr.max_col)
                )
            else:
                new_specs.append(
                    (mr.min_row, mr.min_col, mr.max_row + 1, mr.max_col)
                )
        ExcelMerger._rebuild_merged_ranges(ws, new_specs)

    @staticmethod
    def _adjust_merged_ranges_for_row_delete(
        ws: Worksheet, delete_row: int
    ) -> None:
        """行删除前调整合并区域（Excel 语义）。

        - 完全在删除行上方：不变；
        - 完全在删除行下方：整体上移一行；
        - 包含删除行：向下收缩一行，收缩后为空则移除。
        """
        new_specs: List[Tuple[int, int, int, int]] = []
        for mr in ws.merged_cells.ranges:
            if mr.max_row < delete_row:
                new_specs.append(
                    (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
                )
            elif mr.min_row > delete_row:
                new_specs.append(
                    (mr.min_row - 1, mr.min_col, mr.max_row - 1, mr.max_col)
                )
            else:
                new_max = mr.max_row - 1
                if new_max >= mr.min_row:
                    new_specs.append(
                        (mr.min_row, mr.min_col, new_max, mr.max_col)
                    )
        ExcelMerger._rebuild_merged_ranges(ws, new_specs)

    @staticmethod
    def _adjust_merged_ranges_for_col_insert(
        ws: Worksheet, insert_col: int
    ) -> None:
        """列插入前调整合并区域（镜像行插入语义）。"""
        new_specs: List[Tuple[int, int, int, int]] = []
        for mr in ws.merged_cells.ranges:
            if mr.max_col < insert_col:
                new_specs.append(
                    (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
                )
            elif mr.min_col >= insert_col:
                new_specs.append(
                    (mr.min_row, mr.min_col + 1, mr.max_row, mr.max_col + 1)
                )
            else:
                new_specs.append(
                    (mr.min_row, mr.min_col, mr.max_row, mr.max_col + 1)
                )
        ExcelMerger._rebuild_merged_ranges(ws, new_specs)

    @staticmethod
    def _adjust_merged_ranges_for_col_delete(
        ws: Worksheet, delete_col: int
    ) -> None:
        """列删除前调整合并区域（镜像行删除语义）。"""
        new_specs: List[Tuple[int, int, int, int]] = []
        for mr in ws.merged_cells.ranges:
            if mr.max_col < delete_col:
                new_specs.append(
                    (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
                )
            elif mr.min_col > delete_col:
                new_specs.append(
                    (mr.min_row, mr.min_col - 1, mr.max_row, mr.max_col - 1)
                )
            else:
                new_max = mr.max_col - 1
                if new_max >= mr.min_col:
                    new_specs.append(
                        (mr.min_row, mr.min_col, mr.max_row, new_max)
                    )
        ExcelMerger._rebuild_merged_ranges(ws, new_specs)

    @staticmethod
    def _ensure_writable_cell(ws: Worksheet, row: int, col: int):
        """返回可写的目标单元格；若该位置属于某个合并区域，先解除该合并。

        覆盖写入合并单元格的非左上角位置时，openpyxl 返回只读的
        MergedCell（赋值会抛 AttributeError）。合并工具在这里的语义是
        "以复制值覆盖"，因此需要先解除阻碍写入的合并区域。
        """
        cell = ws.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            for mr in list(ws.merged_cells.ranges):
                if (
                    mr.min_row <= row <= mr.max_row
                    and mr.min_col <= col <= mr.max_col
                ):
                    ws.unmerge_cells(str(mr))
            return ws.cell(row=row, column=col)
        return cell

    @staticmethod
    def _get_cached(
        sheet_data: SheetData, src_row: int, col: int
    ) -> Optional[str]:
        """安全获取 sheet_data.cached_values 中指定单元格的缓存值。

        无 cached_values 或越界时返回 None。
        """
        if (
            sheet_data.cached_values is None
            or src_row is None
            or col is None
            or src_row >= len(sheet_data.cached_values)
        ):
            return None
        row = sheet_data.cached_values[src_row]
        if col >= len(row):
            return None
        v = row[col]
        return v if v != "" else None

    @staticmethod
    def _sync_exclusive_columns(
        diff_result: DiffResult,
        source: SheetData,
        target: SheetData,
        src_is_right: bool,
    ) -> None:
        """把 source 侧独占列插入 target 对应位置（从右往左，避免位置偏移）。

        列值按对齐行复制：same/different 行两侧都有数据，复制 source 列值；
        target 独有行留空。合并前先同步列结构，保证随后插入的行不会因
        两侧列结构不同而错位或丢数据。
        """
        aligned_cols = diff_result.aligned_cols
        row_mapping: List[Tuple[int, int]] = [
            (
                (p.right_row if src_is_right else p.left_row),
                (p.left_row if src_is_right else p.right_row),
            )
            for p in diff_result.aligned_rows
            if p.status in ("same", "different")
            and p.left_row is not None
            and p.right_row is not None
        ]
        for aligned_idx in range(len(aligned_cols) - 1, -1, -1):
            cp = aligned_cols[aligned_idx]
            src_col = cp.right_col if src_is_right else cp.left_col
            tgt_col = cp.left_col if src_is_right else cp.right_col
            if src_col is None or tgt_col is not None:
                continue
            insert_idx = sum(
                1
                for c in aligned_cols[:aligned_idx]
                if (c.left_col if src_is_right else c.right_col) is not None
            )
            ExcelMerger.insert_column_to_other(
                source, src_col, target, insert_idx, 0, row_mapping
            )

    @staticmethod
    def merge_right_to_left(
        diff_result: DiffResult, left: SheetData, right: SheetData
    ) -> None:
        """右覆盖：以右侧为准，将右侧与左侧不同的单元格值及样式复制到左侧。

        - "different" 行：按 diff_cells（aligned_col 索引）逐列把右侧单元格复制到左侧。
          使用 aligned_cols 映射到实际列号。
        - "right_only" 行：在左侧对应位置插入新行并复制右侧整行（保持对齐位置，
          而非追加到末尾）。多个插入按从后往前执行，避免位置偏移。
        - "right_only" 列：先插入左侧对应位置（含表头与匹配行数据），再处理
          行插入，确保两侧列结构不同时不错位、不丢数据。
        """
        aligned_cols = diff_result.aligned_cols

        # 1. 覆盖 different 单元格（必须在插列之前，使用原始列索引）
        for pair in diff_result.aligned_rows:
            if pair.status != "different":
                continue
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
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    left.worksheet, left_ws_row, cp.left_col + 1
                )
                cached = ExcelMerger._get_cached(
                    right, pair.right_row, cp.right_col
                )
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)

        # 2. 右侧独占列插入左侧（右覆盖：以右侧列结构为准）
        ExcelMerger._sync_exclusive_columns(
            diff_result, right, left, src_is_right=True
        )

        # 3. right_only 行按位置插入（此时左侧每个对齐列都有实际列，
        #    物理列号 = aligned_idx）
        inserts: List[Tuple[int, int]] = [
            (
                sum(
                    1
                    for p in diff_result.aligned_rows[:i]
                    if p.left_row is not None
                ),
                pair.right_row,
            )
            for i, pair in enumerate(diff_result.aligned_rows)
            if pair.status == "right_only"
        ]
        col_mapping = [
            (cp.right_col, idx)
            for idx, cp in enumerate(aligned_cols)
            if cp.right_col is not None
        ]
        for insert_idx, src_row in sorted(
            inserts, key=lambda x: x[0], reverse=True
        ):
            ExcelMerger.insert_row_to_other(
                right, src_row, left, insert_idx, 0, col_mapping
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
        - "left_only" 列：先插入右侧对应位置（含表头与匹配行数据），再处理
          行插入，确保两侧列结构不同时不错位、不丢数据。
        """
        aligned_cols = diff_result.aligned_cols

        # 1. 覆盖 different 单元格（必须在插列之前，使用原始列索引）
        for pair in diff_result.aligned_rows:
            if pair.status != "different":
                continue
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
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    right.worksheet, right_ws_row, cp.right_col + 1
                )
                cached = ExcelMerger._get_cached(
                    left, pair.left_row, cp.left_col
                )
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)

        # 2. 左侧独占列插入右侧（左覆盖：以左侧列结构为准）
        ExcelMerger._sync_exclusive_columns(
            diff_result, left, right, src_is_right=False
        )

        # 3. left_only 行按位置插入
        inserts: List[Tuple[int, int]] = [
            (
                sum(
                    1
                    for p in diff_result.aligned_rows[:i]
                    if p.right_row is not None
                ),
                pair.left_row,
            )
            for i, pair in enumerate(diff_result.aligned_rows)
            if pair.status == "left_only"
        ]
        col_mapping = [
            (cp.left_col, idx)
            for idx, cp in enumerate(aligned_cols)
            if cp.left_col is not None
        ]
        for insert_idx, src_row in sorted(
            inserts, key=lambda x: x[0], reverse=True
        ):
            ExcelMerger.insert_row_to_other(
                left, src_row, right, insert_idx, 0, col_mapping
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
        - 右侧独占列先同步到左侧（含表头与匹配行数据），追加行不会丢列。
        - key_cols 参数保留以匹配调用方签名，本方法不依赖它。
        """
        aligned_cols = diff_result.aligned_cols

        # 先同步右侧独占列，再按对齐索引追加整行
        ExcelMerger._sync_exclusive_columns(
            diff_result, right, left, src_is_right=True
        )
        col_mapping = [
            (cp.right_col, idx)
            for idx, cp in enumerate(aligned_cols)
            if cp.right_col is not None
        ]
        for pair in diff_result.aligned_rows:
            if pair.status == "right_only":
                ExcelMerger._append_row(
                    right, pair.right_row, left, 0, col_mapping
                )

    @staticmethod
    def copy_row_to_other(
        source: SheetData,
        source_row_idx: int,
        target: SheetData,
        target_row_idx: int,
        max_col: int,
        col_mapping: Optional[List[Tuple[int, int]]] = None,
    ) -> None:
        """把 source 指定行(0-based)整行值与样式复制到 target 指定行(0-based)。

        - col_mapping: 提供 [(src_col_0based, tgt_col_0based), ...] 时按映射复制，
          只复制映射中的列（用于两侧列结构不同时按对齐列复制）。
          不提供时按位置 0..max_col-1 复制（向后兼容）。
        - target_row_idx 必须是已存在的行。
        """
        src_ws_row = source_row_idx + 1
        tgt_ws_row = target_row_idx + 1
        if col_mapping is not None:
            for src_col, tgt_col in col_mapping:
                src_cell = source.worksheet.cell(row=src_ws_row, column=src_col + 1)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, tgt_ws_row, tgt_col + 1
                )
                cached = ExcelMerger._get_cached(source, source_row_idx, src_col)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)
        else:
            for col in range(max_col):
                ws_col = col + 1
                src_cell = source.worksheet.cell(row=src_ws_row, column=ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, tgt_ws_row, ws_col
                )
                cached = ExcelMerger._get_cached(source, source_row_idx, col)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)

    @staticmethod
    def insert_row_to_other(
        source: SheetData,
        source_row_idx: int,
        target: SheetData,
        target_insert_row_idx: int,
        max_col: int,
        col_mapping: Optional[List[Tuple[int, int]]] = None,
    ) -> int:
        """在 target 指定位置插入新行，并从 source 复制整行值与样式。

        用于把 left_only/right_only 行复制到对侧对应位置（在虚拟空行处新增行）。

        - col_mapping: 提供 [(src_col_0based, tgt_col_0based), ...] 时按映射复制，
          只复制映射中的列（用于两侧列结构不同时按对齐列复制）。
          不提供时按位置 0..max_col-1 复制（向后兼容）。
        - target_insert_row_idx: 0-based，新行插入后位于该索引位置；
          原该位置及之后的行整体后移（openpyxl insert_rows）。
        - 合并单元格重建：仅处理以源行为左上角的合并区域，在 target 中以插入行
          为新 min_row、保持列范围与行跨度不变重建。
        - 返回插入行号(1-based)。
        """
        insert_ws_row = target_insert_row_idx + 1  # 0-based -> 1-based
        # openpyxl 不会同步移动合并区域，先按 Excel 语义调整，再插入行
        ExcelMerger._adjust_merged_ranges_for_row_insert(
            target.worksheet, insert_ws_row
        )
        target.worksheet.insert_rows(insert_ws_row)

        src_ws_row = source_row_idx + 1
        if col_mapping is not None:
            for src_col, tgt_col in col_mapping:
                src_cell = source.worksheet.cell(row=src_ws_row, column=src_col + 1)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, insert_ws_row, tgt_col + 1
                )
                cached = ExcelMerger._get_cached(source, source_row_idx, src_col)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)
        else:
            for col in range(max_col):
                ws_col = col + 1
                src_cell = source.worksheet.cell(row=src_ws_row, column=ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, insert_ws_row, ws_col
                )
                cached = ExcelMerger._get_cached(source, source_row_idx, col)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
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
        tgt_cell = ExcelMerger._ensure_writable_cell(
            target.worksheet, target_row_idx + 1, col_idx + 1
        )
        cached = ExcelMerger._get_cached(source, source_row_idx, col_idx)
        ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
        ExcelLoader.copy_cell_style(src_cell, tgt_cell)

    @staticmethod
    def copy_column_to_other(
        source: SheetData,
        source_col_idx: int,
        target: SheetData,
        target_col_idx: int,
        max_row: int,
        row_mapping: Optional[List[Tuple[int, int]]] = None,
    ) -> None:
        """把 source 指定列(0-based)整列值与样式复制到 target 指定列(0-based)。

        - row_mapping: 提供 [(src_row_0based, tgt_row_0based), ...] 时按映射复制，
          只复制映射中的行（用于两侧行结构不同时按对齐行复制）。
          不提供时按位置 0..max_row-1 复制（向后兼容）。
        - target_col_idx 必须是已存在的列。
        """
        src_ws_col = source_col_idx + 1
        tgt_ws_col = target_col_idx + 1
        if row_mapping is not None:
            for src_row, tgt_row in row_mapping:
                src_cell = source.worksheet.cell(row=src_row + 1, column=src_ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, tgt_row + 1, tgt_ws_col
                )
                cached = ExcelMerger._get_cached(source, src_row, source_col_idx)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)
        else:
            for row in range(max_row):
                ws_row = row + 1
                src_cell = source.worksheet.cell(row=ws_row, column=src_ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, ws_row, tgt_ws_col
                )
                cached = ExcelMerger._get_cached(source, row, source_col_idx)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)

    @staticmethod
    def insert_column_to_other(
        source: SheetData,
        source_col_idx: int,
        target: SheetData,
        target_insert_col_idx: int,
        max_row: int,
        row_mapping: Optional[List[Tuple[int, int]]] = None,
    ) -> int:
        """在 target 指定位置插入新列，并从 source 复制整列值与样式。

        用于把 left_only/right_only 列复制到对侧对应位置（在虚拟空列处新增列）。

        - row_mapping: 提供 [(src_row_0based, tgt_row_0based), ...] 时按映射复制，
          只复制映射中的行（用于两侧行结构不同时按对齐行复制）。
          不提供时按位置 0..max_row-1 复制（向后兼容）。
        - target_insert_col_idx: 0-based，新列插入后位于该索引位置；
          原该位置及之后的列整体右移（openpyxl insert_cols）。
        - 合并单元格重建：仅处理以源列为左上角的合并区域，在 target 中以插入列
          为新 min_col、保持行范围与列跨度不变重建。
        - 返回插入列号(1-based)。
        """
        insert_ws_col = target_insert_col_idx + 1  # 0-based -> 1-based
        # openpyxl 不会同步移动合并区域，先按 Excel 语义调整，再插入列
        ExcelMerger._adjust_merged_ranges_for_col_insert(
            target.worksheet, insert_ws_col
        )
        target.worksheet.insert_cols(insert_ws_col)

        src_ws_col = source_col_idx + 1
        if row_mapping is not None:
            for src_row, tgt_row in row_mapping:
                src_cell = source.worksheet.cell(row=src_row + 1, column=src_ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, tgt_row + 1, insert_ws_col
                )
                cached = ExcelMerger._get_cached(source, src_row, source_col_idx)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)
        else:
            for row in range(max_row):
                ws_row = row + 1
                src_cell = source.worksheet.cell(row=ws_row, column=src_ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, ws_row, insert_ws_col
                )
                cached = ExcelMerger._get_cached(source, row, source_col_idx)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)

        # 重建以源列为左上角的合并区域
        for min_r, min_c, max_r, max_c in ExcelMerger._find_merged_ranges_for_col(
            source, src_ws_col
        ):
            new_min_col = insert_ws_col
            new_max_col = insert_ws_col + (max_c - min_c)
            try:
                target.worksheet.merge_cells(
                    start_row=min_r,
                    start_column=new_min_col,
                    end_row=max_r,
                    end_column=new_max_col,
                )
            except Exception:
                continue

        return insert_ws_col

    @staticmethod
    def delete_row(target: SheetData, target_row_idx: int) -> None:
        """删除 target 指定行(0-based)。

        用于从少的一侧（虚拟空行）右键复制到多的一侧时，删除多出的一侧对应的行，
        使两侧行数一致。多个删除按从后往前执行以避免位置偏移。
        """
        del_row = target_row_idx + 1
        ExcelMerger._adjust_merged_ranges_for_row_delete(
            target.worksheet, del_row
        )
        target.worksheet.delete_rows(del_row)

    @staticmethod
    def delete_column(target: SheetData, target_col_idx: int) -> None:
        """删除 target 指定列(0-based)。

        用于从少的一侧（虚拟空列）右键复制到多的一侧时，删除多出的一侧对应的列，
        使两侧列数一致。多个删除按从后往前执行以避免位置偏移。
        """
        del_col = target_col_idx + 1
        ExcelMerger._adjust_merged_ranges_for_col_delete(
            target.worksheet, del_col
        )
        target.worksheet.delete_cols(del_col)

    @staticmethod
    def _append_row(
        source: SheetData,
        source_row_idx: int,
        target: SheetData,
        max_col: int,
        col_mapping: Optional[List[Tuple[int, int]]] = None,
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

        # 逐列复制值与样式（提供列映射时按对齐列复制，避免结构不同时错位）
        if col_mapping is not None:
            for src_col, tgt_col in col_mapping:
                src_cell = source.worksheet.cell(
                    row=src_ws_row, column=src_col + 1
                )
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, append_row, tgt_col + 1
                )
                cached = ExcelMerger._get_cached(
                    source, source_row_idx, src_col
                )
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
                ExcelLoader.copy_cell_style(src_cell, tgt_cell)
        else:
            for col in range(max_col):
                ws_col = col + 1
                src_cell = source.worksheet.cell(row=src_ws_row, column=ws_col)
                tgt_cell = ExcelMerger._ensure_writable_cell(
                    target.worksheet, append_row, ws_col
                )
                cached = ExcelMerger._get_cached(source, source_row_idx, col)
                ExcelLoader.copy_cell_value(src_cell, tgt_cell, cached)
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

    @staticmethod
    def _find_merged_ranges_for_col(
        sheet_data: SheetData, col_1based: int
    ) -> List[Tuple[int, int, int, int]]:
        """返回 sheet_data.merged_ranges 中 min_col == col_1based 的所有 range。

        即以该列为左上角的合并区域。
        """
        return [
            (min_r, min_c, max_r, max_c)
            for (min_r, min_c, max_r, max_c) in sheet_data.merged_ranges
            if min_c == col_1based
        ]
