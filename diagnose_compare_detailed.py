#!/usr/bin/env python3
"""
增强诊断脚本：详细检查单元格值差异和对齐结果。
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.core.excel_loader import ExcelLoader, SheetData
from src.core.comparator import ExcelComparator


def diagnose_compare_detailed(left_path: str, right_path: str, key_col: int = None):
    """详细诊断比较两个 Excel 文件。"""
    
    print(f"正在加载文件...")
    print(f"  左侧: {left_path}")
    print(f"  右侧: {right_path}")
    print(f"  关键列: {key_col}")
    print()
    
    # 加载工作簿
    left_wb = ExcelLoader.load_workbook(left_path)
    right_wb = ExcelLoader.load_workbook(right_path)
    
    # 获取第一个工作表
    left_sheet_name = left_wb.sheetnames[0]
    right_sheet_name = right_wb.sheetnames[0]
    
    print(f"工作表: 左={left_sheet_name}, 右={right_sheet_name}")
    print()
    
    left_ws = ExcelLoader.get_worksheet(left_wb, left_sheet_name)
    right_ws = ExcelLoader.get_worksheet(right_wb, right_sheet_name)
    
    left_sd = ExcelLoader.extract_sheet_data(left_ws)
    right_sd = ExcelLoader.extract_sheet_data(right_ws)
    
    print(f"数据维度: 左={len(left_sd.values)}行, 右={len(right_sd.values)}行")
    print()
    
    # 执行比较
    key_cols = [key_col] if key_col is not None else []
    result = ExcelComparator.compare_sheets(left_sd, right_sd, key_cols=key_cols)
    
    # 打印对齐结果
    print("=" * 100)
    print("对齐结果:")
    print("=" * 100)
    
    for i, pair in enumerate(result.aligned_rows):
        left_val = left_sd.values[pair.left_row] if pair.left_row is not None else None
        right_val = right_sd.values[pair.right_row] if pair.right_row is not None else None
        
        left_str = str(left_val[:3]) if left_val else "<空>"
        right_str = str(right_val[:3]) if right_val else "<空>"
        
        # 安全格式化（处理 None）
        left_idx = pair.left_row if pair.left_row is not None else -1
        right_idx = pair.right_row if pair.right_row is not None else -1
        
        print(f"[{i:3d}] left={left_idx:3}  right={right_idx:3}  "
              f"status={pair.status:12}  L:{left_str}  R:{right_str}")
    
    print()
    print("=" * 100)
    print(f"统计: {result.stats}")
    print(f"对齐行数: {len(result.aligned_rows)}")
    print("=" * 100)
    
    # 详细检查：为什么 identical 的行被标记为 different
    print()
    print("=" * 100)
    print("详细差异分析（前 10 个 different 行）:")
    print("=" * 100)
    
    diff_count = 0
    for i, pair in enumerate(result.aligned_rows):
        if pair.status == 'different' and diff_count < 10:
            left_row = left_sd.values[pair.left_row] if pair.left_row is not None else None
            right_row = right_sd.values[pair.right_row] if pair.right_row is not None else None
            
            print(f"\n对齐行 {i}:")
            print(f"  左侧行 {pair.left_row}:")
            if left_row:
                for c, val in enumerate(left_row):
                    print(f"    列{c}: repr={repr(val)}")
            print(f"  右侧行 {pair.right_row}:")
            if right_row:
                for c, val in enumerate(right_row):
                    print(f"    列{c}: repr={repr(val)}")
            
            print(f"  差异列: {pair.diff_cells}")
            
            # 检查每列的差异
            if left_row and right_row:
                for c in pair.diff_cells:
                    if c < len(left_row) and c < len(right_row):
                        l_val = left_row[c]
                        r_val = right_row[c]
                        print(f"    列{c} 差异:")
                        print(f"      L: repr={repr(l_val)}, type={type(l_val).__name__}")
                        print(f"      R: repr={repr(r_val)}, type={type(r_val).__name__}")
                        print(f"      L == R: {l_val == r_val}")
            
            diff_count += 1
    
    # 检查对齐是否正确
    print()
    print("=" * 100)
    print("对齐正确性检查:")
    print("=" * 100)
    
    # 检查是否有 left_only 或 right_only
    left_only_count = sum(1 for p in result.aligned_rows if p.status == 'left_only')
    right_only_count = sum(1 for p in result.aligned_rows if p.status == 'right_only')
    
    print(f"left_only 行数: {left_only_count}")
    print(f"right_only 行数: {right_only_count}")
    
    if left_only_count == 0 and right_only_count == 0:
        print()
        print("⚠️  警告: 没有检测到 left_only 或 right_only 行！")
        print("这说明 LCS 对齐算法可能没有正确执行。")
        print()
        print("可能的原因:")
        print("  1. 所有行都不相同（different），所以没有相同的锚点行")
        print("  2. LCS 算法实现有 bug")
        print("  3. 行签名计算有误")
        
        # 检查是否有任何 same 行
        same_count = sum(1 for p in result.aligned_rows if p.status == 'same')
        print(f"\n实际 same 行数: {same_count}")
        
        if same_count == 0:
            print("所有行都是 different！这说明比较逻辑有问题。")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("使用方法: python diagnose_compare_detailed.py <左侧文件> <右侧文件> [关键列索引]")
        print("示例: python diagnose_compare_detailed.py left.xlsx right.xlsx 0")
        sys.exit(1)
    
    left_path = sys.argv[1]
    right_path = sys.argv[2]
    key_col = int(sys.argv[3]) if len(sys.argv) > 3 else None
    
    diagnose_compare_detailed(left_path, right_path, key_col)
