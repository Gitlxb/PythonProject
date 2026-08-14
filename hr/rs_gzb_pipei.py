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


def _is_formula(value):
    """判断单元格值是否为公式（以 '=' 开头的字符串）"""
    return isinstance(value, str) and value.strip().startswith("=")


def _is_total_row(formula_ws, row, header_cols):
    """
    检测指定行是否为合计行
    判定条件：行内任意一个表头对应列的单元格包含 =SUM() 公式
    """
    for col in header_cols:
        formula = formula_ws.cell(row=row, column=col).value
        if isinstance(formula, str) and formula.strip().upper().startswith("=SUM("):
            return True
    return False


def _is_data_total_row(formula_ws, data_ws, row, key_col, header_cols):
    """
    合计行判定双条件：关键列为空（或"合计"）且行内有 =SUM() 公式
    避免误判数据行中使用了 =SUM() 的行间/列间求和行为合计行
    """
    key_val = data_ws.cell(row=row, column=key_col).value
    key_str = str(key_val).strip() if key_val is not None else ""
    return (key_str == "" or key_str == "合计") and _is_total_row(formula_ws, row, header_cols)


def _compare_cell_full(v1, v2, f1, f2):
    """
    比较两个单元格：值 + 公式 双维度
    返回: (value_match, formula_match)
    """
    # 值比较（复用已有逻辑）
    value_match = _compare_values(v1, v2)

    # 公式比较
    f1_is_formula = _is_formula(f1)
    f2_is_formula = _is_formula(f2)

    if f1_is_formula and f2_is_formula:
        # 两边都是公式：去除首尾空格后字面对比
        formula_match = (f1.strip() == f2.strip())
    else:
        # 只有一边是公式或都不是公式：公式维度不参与判定，以值比较为准
        formula_match = True

    return value_match, formula_match


def _find_sections(data_ws, formula_ws, header_row, key_col, key_field_name, header_cols):
    """
    自动发现工作表内的多个数据段（sections）。
    向前扫描，当关键列中出现与 key_field_name 相同的值时，视为新段的表头行。

    返回: [(header_row, last_data_row, has_total_row), ...]
    """
    sections = []
    current_header = header_row

    scan_row = header_row + 1
    while scan_row <= data_ws.max_row:
        cell_val = data_ws.cell(row=scan_row, column=key_col).value
        key_str = str(cell_val).strip() if cell_val is not None else ""

        if key_str == key_field_name:
            # 发现新段表头：上一段结束于 scan_row-1
            end_row = scan_row - 1
            has_total = False
            if end_row >= current_header + 1:
                has_total = _is_data_total_row(formula_ws, data_ws, end_row, key_col, header_cols)
            sections.append((current_header, end_row, has_total))
            current_header = scan_row

        scan_row += 1

    # 最后一段：用 _find_last_data_row 找到尾行
    last_row = _find_last_data_row(data_ws, current_header, header_cols)
    if last_row > current_header:
        has_total = _is_data_total_row(formula_ws, data_ws, last_row, key_col, header_cols)
        sections.append((current_header, last_row, has_total))

    return sections if sections else [(header_row, last_row, False)]


def compare_sheets_and_mark(wb1_data, wb1_formula, sheet1, header_row1,
                            wb2_data, wb2_formula, sheet2, header_row2,
                            key_field_1, key_field_2, field_mapping):
    """
    按关键列进行键值查找匹配，对比两个工作簿的数据（值 + 公式双维度）

    :param wb1_data: 主表 workbook（data_only=True，用于取值）
    :param wb1_formula: 主表 workbook（data_only=False，用于取公式）
    :param sheet1: 主表工作簿名
    :param header_row1: 主表表头行号
    :param wb2_data: 被匹配表 workbook（data_only=True，用于取值）
    :param wb2_formula: 被匹配表 workbook（data_only=False，用于取公式）
    :param sheet2: 被匹配表工作簿名
    :param header_row2: 被匹配表表头行号
    :param key_field_1: 主表关键列字段名
    :param key_field_2: 被匹配表关键列字段名
    :param field_mapping: [(field_wb1, field_wb2), ...] 字段映射列表
    :return: (highlight_cells, matched_count, unmatched_keys, total_info)
             highlight_cells: set of (row, col) in wb2
             matched_count: 实际成功匹配的行数
             unmatched_keys: 基准表中未在目标表找到匹配的关键值列表
             total_info: 合计行参与情况的汇总 dict
    """
    ws1_d = wb1_data[sheet1]
    ws1_f = wb1_formula[sheet1]
    ws2_d = wb2_data[sheet2]
    ws2_f = wb2_formula[sheet2]

    headers1 = _read_headers(ws1_d, header_row1)
    headers2 = _read_headers(ws2_d, header_row2)

    # 校验关键列是否存在
    if key_field_1 not in headers1:
        raise ValueError(f"工作簿「{sheet1}」中找不到关键列「{key_field_1}」")
    if key_field_2 not in headers2:
        raise ValueError(f"工作簿「{sheet2}」中找不到关键列「{key_field_2}」")

    key_col1 = headers1[key_field_1]
    key_col2 = headers2[key_field_2]

    header_cols1 = list(headers1.values())
    header_cols2 = list(headers2.values())

    # ---- 自动发现基准表和目标表的所有数据段 ----
    sections1 = _find_sections(ws1_d, ws1_f, header_row1, key_col1, key_field_1, header_cols1)
    sections2 = _find_sections(ws2_d, ws2_f, header_row2, key_col2, key_field_2, header_cols2)

    # ---- 构建目标表全段关键列索引 ----
    target_index = {}
    total_row_in_target = None
    for h, last, _ in sections2:
        for row2 in range(h + 1, last + 1):
            if _is_data_total_row(ws2_f, ws2_d, row2, key_col2, header_cols2):
                total_row_in_target = row2
                continue
            key_val = ws2_d.cell(row=row2, column=key_col2).value
            if key_val is None:
                continue
            key_str = str(key_val).strip()
            if key_str == "" or key_str == "合计":
                continue
            if key_str not in target_index:
                target_index[key_str] = row2

    # ---- 统计基准表全段数据行数和是否有合计行 ----
    base_has_total = any(s[2] for s in sections1)
    base_data_rows = 0
    for h, last, _ in sections1:
        for row1 in range(h + 1, last + 1):
            if _is_data_total_row(ws1_f, ws1_d, row1, key_col1, header_cols1):
                continue
            key_val = ws1_d.cell(row=row1, column=key_col1).value
            if key_val is None:
                continue
            key_str = str(key_val).strip()
            if key_str and key_str != "合计":
                base_data_rows += 1

    # ---- 合计行参与条件 ----
    target_has_total = total_row_in_target is not None
    target_data_rows = len(target_index)
    total_can_participate = (base_has_total and target_has_total
                             and base_data_rows == target_data_rows)

    # ---- 遍历基准表全段，执行匹配和对比 ----
    highlight_cells = set()
    matched_count = 0
    unmatched_keys = []
    matched_target_keys = set()

    for h, last, _ in sections1:
        for row1 in range(h + 1, last + 1):
            is_total_in_base = _is_data_total_row(ws1_f, ws1_d, row1, key_col1, header_cols1)
            key_val = ws1_d.cell(row=row1, column=key_col1).value
            key_str = str(key_val).strip() if key_val is not None else ""

            if is_total_in_base:
                if total_can_participate:
                    row2 = total_row_in_target
                else:
                    continue
            elif key_str and key_str != "合计":
                row2 = target_index.get(key_str)
                if row2 is None:
                    unmatched_keys.append(key_str)
                    continue
                matched_target_keys.add(key_str)
            else:
                continue

            matched_count += 1

            # 逐字段值 + 公式对比
            for f1, f2 in field_mapping:
                if f1 not in headers1 or f2 not in headers2:
                    continue
                col1 = headers1[f1]
                col2 = headers2[f2]
                v1 = ws1_d.cell(row=row1, column=col1).value
                v2 = ws2_d.cell(row=row2, column=col2).value
                f1_val = ws1_f.cell(row=row1, column=col1).value
                f2_val = ws2_f.cell(row=row2, column=col2).value
                val_ok, formula_ok = _compare_cell_full(v1, v2, f1_val, f2_val)
                if not val_ok or not formula_ok:
                    highlight_cells.add((row2, col2))

    extra_target_keys = [k for k in target_index if k not in matched_target_keys]
    total_info = {
        "participated": total_can_participate,
        "base_has_total": base_has_total,
        "target_has_total": target_has_total,
        "base_data_rows": base_data_rows,
        "target_data_rows": target_data_rows,
        "extra_target_keys": extra_target_keys,
    }

    return highlight_cells, matched_count, unmatched_keys, total_info


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
                   key_field_1, key_field_2,
                   field_mapping, wb_out=None):
    """
    模式A：甲方工资表（基准）匹配内部工资表
    按关键列键值查找，值+公式双维度对比，对内部表标红差异单元格。

    :param key_field_1: 甲方关键列字段名
    :param key_field_2: 内部关键列字段名
    :param field_mapping: [(甲方字段, 内部字段), ...]
    :param wb_out: 已加载的内部 workbook（可选，用于多次匹配累积标红）
    :return: (wb_out, matched_count, diff_count, unmatched_keys, total_info)
             wb_out: 处理后的内部 workbook 对象（尚未保存）
    """
    # 1. 双加载：data_only=True 用于取值，data_only=False 用于取公式
    wb_jia_data = openpyxl.load_workbook(jia_file, data_only=True)
    wb_jia_formula = openpyxl.load_workbook(jia_file, data_only=False)
    wb_nei_data = openpyxl.load_workbook(nei_file, data_only=True)
    wb_nei_formula = openpyxl.load_workbook(nei_file, data_only=False)

    try:
        highlight_cells, matched_count, unmatched_keys, total_info = compare_sheets_and_mark(
            wb_jia_data, wb_jia_formula, jia_sheet, jia_header_row,
            wb_nei_data, wb_nei_formula, nei_sheet, nei_header_row,
            key_field_1, key_field_2,
            field_mapping
        )
    finally:
        wb_jia_data.close()
        wb_jia_formula.close()
        wb_nei_data.close()
        wb_nei_formula.close()

    # 2. 普通模式加载内部文件用于输出（保留公式和格式）
    if wb_out is None:
        wb_out = openpyxl.load_workbook(nei_file)

    # 3. 对内部表应用标红
    apply_highlight_to_cells(wb_out, nei_sheet, highlight_cells)

    return wb_out, matched_count, len(highlight_cells), unmatched_keys, total_info


def process_mode_b_compare(file_path, sheet_a, header_a, sheet_b, header_b,
                           key_field_1, key_field_2, field_mapping):
    """
    模式B：内部工作簿A 匹配 内部工作簿B（同一文件）
    按关键列键值查找，值+公式双维度对比，仅执行对比并返回标红坐标。

    :param key_field_1: 工作簿A关键列字段名
    :param key_field_2: 工作簿B关键列字段名
    :return: (highlight_cells, matched_count, diff_count, unmatched_keys, total_info)
    """
    wb_data = openpyxl.load_workbook(file_path, data_only=True)
    wb_formula = openpyxl.load_workbook(file_path, data_only=False)
    try:
        highlight_cells, matched_count, unmatched_keys, total_info = compare_sheets_and_mark(
            wb_data, wb_formula, sheet_a, header_a,
            wb_data, wb_formula, sheet_b, header_b,
            key_field_1, key_field_2,
            field_mapping
        )
        return highlight_cells, matched_count, len(highlight_cells), unmatched_keys, total_info
    finally:
        wb_data.close()
        wb_formula.close()
