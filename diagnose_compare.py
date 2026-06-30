#!/usr/bin/env python3
"""
诊断脚本：检查比较引擎对实际 Excel 文件的输出。
使用方法：
    python diagnose_compare.py <左侧文件> <右侧文件> [关键列索引]
"""

import sys
import os

# 添加 src 目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.core.excel_loader import ExcelLoader, SheetData
from src.core.comparator import ExcelComparator


def diagnose_compare(left_path: str, right_path: str, key_col: int = None):
    """诊断比较两个 Excel 文件。"""
    
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
    print("=" * 80)
    print("对齐结果:")
    print("=" * 80)
    
    for i, pair in enumerate(result.aligned_rows):
        left_val = left_sd.values[pair.left_row] if pair.left_row is not None else None
        right_val = right_sd.values[pair.right_row] if pair.right_row is not None else None
        
        left_str = str(left_val[:3]) if left_val else "<空>"
        right_str = str(right_val[:3]) if right_val else "<空>"
        
        print(f"[{i:3d}] left={pair.left_row:3}  right={pair.right_row:3}  "
              f"status={pair.status:12}  L:{left_str}  R:{right_str}")
    
    print()
    print("=" * 80)
    print(f"统计: {result.stats}")
    print(f"差异行数: {len(result.diff_row_indices)}")
    print("=" * 80)
    
    # 检查是否有错位（different 状态）
    diff_count = result.stats.get('different', 0)
    if diff_count > 0:
        print()
        print(f"⚠️  发现 {diff_count} 行被标记为 'different'（不同）")
        print("如果预期这些行应该相同，请检查：")
        print("  1. 是否有不可见字符（空格、制表符等）")
        print("  2. 大小写是否一致")
        print("  3. 数字格式是否一致（如 '10' vs '10.0'）")
        print()
        
        # 显示前 5 个 different 行的详细差异
        print("前 5 个差异行的详细对比:")
        count = 0
        for i, pair in enumerate(result.aligned_rows):
            if pair.status == 'different' and count < 5:
                left_row = left_sd.values[pair.left_row]
                right_row = right_sd.values[pair.right_row]
                print(f"\n  对齐行 {i}:")
                print(f"    左侧: {left_row}")
                print(f"    右侧: {right_row}")
                print(f"    差异列: {pair.diff_cells}")
                count += 1


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("使用方法: python diagnose_compare.py <左侧文件> <右侧文件> [关键列索引]")
        print("示例: python diagnose_compare.py left.xlsx right.xlsx 0")
        sys.exit(1)
    
    left_path = sys.argv[1]
    right_path = sys.argv[2]
    key_col = int(sys.argv[3]) if len(sys.argv) > 3 else None
    
    diagnose_compare(left_path, right_path, key_col)
