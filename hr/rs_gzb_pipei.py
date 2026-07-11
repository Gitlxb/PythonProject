"""
核心逻辑模块：工资表匹配
与 GUI 解耦，可单独导入调用
"""

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.cell.cell import MergedCell

# 标红样式：红色字体 + 浅蓝色背景
LIGHT_BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")


def get_sheet_names(filepath):
    """获取 Excel 文件的所有工作簿名称（仅支持 .xlsx / .xlsm）"""
    if not filepath.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("仅支持 .xlsx 或 .xlsm 格式的 Excel 文件")
    wb = openpyxl.load_workbook(filepath, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def read_preview_rows(filepath, sheet_name, max_rows=3, max_cols=30):
    """
    读取指定工作簿的前 max_rows 行数据，用于表头行选择展示
    返回: [(行号, [值1, 值2, ...]), ...]
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"工作簿 '{sheet_name}' 不存在")
        ws = wb[sheet_name]
        result = []
        for row_idx in range(1, min(max_rows + 1, ws.max_row + 1)):
            row_data = []
            for col_idx in range(1, max_cols + 1):
                try:
                    val = ws.cell(row=row_idx, column=col_idx).value
                except Exception:
                    val = None
                row_data.append(val)
            result.append((row_idx, row_data))
        return result
    finally:
        wb.close()


def _find_last_data_row(ws, header_row, header_cols=None):
    """
    从底部向上搜索，返回有数据的最后一行（包含合计行）
    空字符串和纯空白字符串视为空白
    :param header_cols: 表头已识别的列号列表（可选），若提供则仅检测这些列
    """
    cols_to_check = header_cols if header_cols else range(1, ws.max_column + 1)
    for row in range(ws.max_row, header_row, -1):
        for col in cols_to_check:
            val = ws.cell(row=row, column=col).value
            if val is not None:
                if isinstance(val, str) and val.strip() == "":
                    continue
                return row
    return header_row


def _read_headers(ws, header_row):
    """读取指定行的表头，返回 {字段名: 列号}"""
    headers = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is not None:
            headers[str(val).strip()] = col
    return headers


def _compare_values(v1, v2):
    """
    比较两个单元格值是否相同
    空白（None/空字符串）视为相同
    左边空白且右边为0 也视为相同（不标红）
    数值型进行容错比较
    """
    # 标准化空白
    if v1 is None or (isinstance(v1, str) and v1.strip() == ""):
        v1 = None
    if v2 is None or (isinstance(v2, str) and v2.strip() == ""):
        v2 = None

    # 都空白
    if v1 is None and v2 is None:
        return True

    # 需求1：左边空白，右边为0 → 视为相同，不标红
    if v1 is None and v2 is not None:
        try:
            if float(v2) == 0:
                return True
        except (ValueError, TypeError):
            pass
        return False

    # 右边空白（左边非空）
    if v2 is None:
        return False

    # 尝试数值比较
    try:
        n1 = float(v1)
        n2 = float(v2)
        return abs(n1 - n2) < 1e-9
    except (ValueError, TypeError):
        pass

    # 字符串比较（去除首尾空格）
    return str(v1).strip() == str(v2).strip()


def compare_sheets_and_mark(wb1, sheet1, header_row1, wb2, sheet2, header_row2, field_mapping):
    """
    对比两个工作簿的数据，返回需要在 wb2 中标红的单元格坐标集合

    :param wb1: 主表 workbook (data_only=True 加载)
    :param sheet1: 主表工作簿名
    :param header_row1: 主表表头行号
    :param wb2: 被匹配表 workbook (data_only=True 加载)
    :param sheet2: 被匹配表工作簿名
    :param header_row2: 被匹配表表头行号
    :param field_mapping: [(field_wb1, field_wb2), ...] 字段映射列表
    :return: (highlight_cells, compare_rows)
             highlight_cells: set of (row, col) in wb2
             compare_rows: 实际对比的行数
    """
    ws1 = wb1[sheet1]
    ws2 = wb2[sheet2]

    headers1 = _read_headers(ws1, header_row1)
    headers2 = _read_headers(ws2, header_row2)

    last_row1 = _find_last_data_row(ws1, header_row1, list(headers1.values()))
    last_row2 = _find_last_data_row(ws2, header_row2, list(headers2.values()))

    data_rows1 = last_row1 - header_row1
    data_rows2 = last_row2 - header_row2
    compare_rows = min(data_rows1, data_rows2)

    highlight_cells = set()

    for i in range(compare_rows):
        row1 = header_row1 + 1 + i
        row2 = header_row2 + 1 + i

        for f1, f2 in field_mapping:
            if f1 not in headers1 or f2 not in headers2:
                continue

            col1 = headers1[f1]
            col2 = headers2[f2]

            v1 = ws1.cell(row=row1, column=col1).value
            v2 = ws2.cell(row=row2, column=col2).value

            if not _compare_values(v1, v2):
                highlight_cells.add((row2, col2))

    return highlight_cells, compare_rows


def apply_highlight_to_cells(wb, sheet_name, highlight_cells):
    """
    在已加载的 workbook 中（普通模式），对指定单元格应用红色字体+黄色背景
    """
    ws = wb[sheet_name]
    for row, col in highlight_cells:
        cell = ws.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            continue
        old_font = cell.font
        cell.font = Font(
            name=old_font.name,
            size=old_font.size,
            bold=old_font.bold,
            italic=old_font.italic,
            underline=old_font.underline,
            strike=old_font.strike,
            color="FF0000"
        )
        cell.fill = LIGHT_BLUE_FILL


def process_mode_a(jia_file, jia_sheet, jia_header_row,
                   nei_file, nei_sheet, nei_header_row,
                   field_mapping, wb_out=None):
    """
    模式A：甲方工资表（基准）匹配内部工资表
    对内部表标红差异单元格后返回 workbook 对象。
    支持传入已有的 wb_out 以在同一文件上累积标红。

    :param field_mapping: [(甲方字段, 内部字段), ...]
    :param wb_out: 已加载的内部 workbook（可选，用于多次匹配累积标红）
    :return: (wb_out, compare_rows, diff_count)
             wb_out: 处理后的内部 workbook 对象（尚未保存）
    """
    # 1. data_only 加载用于对比
    wb_jia_data = openpyxl.load_workbook(jia_file, data_only=True)
    wb_nei_data = openpyxl.load_workbook(nei_file, data_only=True)

    try:
        highlight_cells, compare_rows = compare_sheets_and_mark(
            wb_jia_data, jia_sheet, jia_header_row,
            wb_nei_data, nei_sheet, nei_header_row,
            field_mapping
        )
    finally:
        wb_jia_data.close()
        wb_nei_data.close()

    # 2. 普通模式加载内部文件用于输出（保留公式和格式）
    if wb_out is None:
        wb_out = openpyxl.load_workbook(nei_file)

    # 3. 对内部表应用标红
    apply_highlight_to_cells(wb_out, nei_sheet, highlight_cells)

    return wb_out, compare_rows, len(highlight_cells)


def process_mode_b_compare(file_path, sheet_a, header_a, sheet_b, header_b, field_mapping):
    """
    模式B：内部工作簿A 匹配 内部工作簿B（同一文件）
    仅执行对比，返回需要标红的单元格坐标

    :return: (highlight_cells, compare_rows, diff_count)
    """
    wb_data = openpyxl.load_workbook(file_path, data_only=True)
    try:
        highlight_cells, compare_rows = compare_sheets_and_mark(
            wb_data, sheet_a, header_a,
            wb_data, sheet_b, header_b,
            field_mapping
        )
        return highlight_cells, compare_rows, len(highlight_cells)
    finally:
        wb_data.close()
