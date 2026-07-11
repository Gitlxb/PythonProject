# -*- coding: utf-8 -*-
"""
工具函数模块
包含：apply_excel_styles、match_month_in_text
"""

import re

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def apply_excel_styles(ws: openpyxl.worksheet.worksheet.Worksheet, df: "pd.DataFrame") -> None:
    """应用统一的Excel样式：表头蓝色背景、数据行居中边框、自动列宽

    如果df包含"状态"列，则对"待核对"行应用黄色背景高亮。
    """
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 黄色高亮（用于"待核对"单元格）
    pending_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 找到"状态"列的索引（如果有）
    status_col_idx = None
    if '状态' in df.columns:
        status_col_idx = list(df.columns).index('状态') + 1  # +1 转为1-based

    for col in range(1, df.shape[1] + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    for row_idx in range(2, df.shape[0] + 2):
        for col_idx in range(1, df.shape[1] + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
            # 只给"状态"列中"待核对"的单元格标黄
            if status_col_idx is not None and col_idx == status_col_idx:
                status_cell = ws.cell(row=row_idx, column=status_col_idx)
                if status_cell.value == '待核对':
                    cell.fill = pending_fill

    for col in range(1, df.shape[1] + 1):
        max_length = 0
        column_letter = get_column_letter(col)
        for row_idx in range(1, df.shape[0] + 2):
            cell = ws.cell(row=row_idx, column=col)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column_letter].width = max(adjusted_width, 10)


def match_month_in_text(month: str, text: str) -> bool:
    """精确匹配月份字符串，避免'1月'匹配到'12月'"""
    import pandas as pd
    if pd.isna(text):
        return False
    pattern = rf'(?<!\d){re.escape(month)}(?!\d)'
    return bool(re.search(pattern, str(text).strip()))


