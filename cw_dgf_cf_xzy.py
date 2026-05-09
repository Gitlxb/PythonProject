"""
核心逻辑模块：Excel 数据处理
与 GUI 解耦，可单独导入调用
"""

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from collections import defaultdict

# 标黄填充样式
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# 需要标黄判断的关键字
HIGHLIGHT_KEYWORDS = ["公司", "个体工商户", "刘先锋现金卡"]


def get_sheet_names(filepath):
    """
    获取 Excel 文件的所有工作簿名称（仅支持 .xlsx / .xlsm）
    """
    if not filepath.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("仅支持 .xlsx 或 .xlsm 格式的 Excel 文件，请先另存为 xlsx 格式")
    wb = openpyxl.load_workbook(filepath, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def _scan_headers(sheet, keywords, max_check_rows=3):
    """
    扫描 sheet 的前 max_check_rows 行，寻找包含所有 keywords 的表头行
    :return: (header_row, {keyword: [col1, col2, ...]})
             若未找到返回 (None, {})
    """
    keywords_set = set(keywords)
    limit_row = min(max_check_rows + 1, sheet.max_row + 1)
    for row_idx in range(1, limit_row):
        found = {}
        for col_idx in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            if val in keywords_set:
                found.setdefault(val, []).append(col_idx)
        if all(k in found for k in keywords_set):
            return row_idx, found
    return None, {}


def _find_last_data_row(sheet, header_row, cols_to_check):
    """
    从底部向上搜索，返回 cols_to_check 中任一列有值的最后一行
    """
    for row in range(sheet.max_row, header_row, -1):
        if any(sheet.cell(row=row, column=c).value is not None for c in cols_to_check):
            return row
    return header_row


def process_first_excel(filepath, sheet_name):
    """
    处理第一个 Excel：
      1. 按"介绍人"汇总"合计金额"
      2. 在原数据末尾向下空 5 行写入汇总结果（表头 + 数据）
      3. 保存文件
    :return: [(介绍人, 金额), ...]  供第二阶段直接使用
    """
    if not filepath.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("仅支持 .xlsx 或 .xlsm 格式")

    # ---- 阶段 A：用 data_only=True 读取公式计算结果 ----
    wb_read = openpyxl.load_workbook(filepath, data_only=True)
    if sheet_name not in wb_read.sheetnames:
        wb_read.close()
        raise ValueError(f"工作簿 '{sheet_name}' 不存在")
    sheet_read = wb_read[sheet_name]

    header_row, found = _scan_headers(sheet_read, ["介绍人", "合计金额"])
    if header_row is None:
        wb_read.close()
        raise ValueError("未在首 3 行找到表头：'介绍人' 和 '合计金额'")

    jieshao_col = found["介绍人"][0]
    heji_col = found["合计金额"][0]

    summary = defaultdict(float)
    for row in range(header_row + 1, sheet_read.max_row + 1):
        name = sheet_read.cell(row=row, column=jieshao_col).value
        amount = sheet_read.cell(row=row, column=heji_col).value
        if name is not None and amount is not None:
            try:
                amount = float(amount)
            except (ValueError, TypeError):
                continue
            summary[name] += amount

    wb_read.close()

    if not summary:
        raise ValueError("未找到有效数据（介绍人或合计金额为空）")

    # ---- 阶段 B：用普通模式打开，保留原格式并写入汇总 ----
    wb = openpyxl.load_workbook(filepath)
    sheet = wb[sheet_name]

    last_data_row = _find_last_data_row(sheet, header_row, [jieshao_col, heji_col])
    write_start = last_data_row + 5 + 1  # 空 5 行，再下一行开始写

    # 写汇总表头
    sheet.cell(row=write_start, column=1, value="介绍人")
    sheet.cell(row=write_start, column=2, value="求和项：金额")

    result = []
    for i, (name, amount) in enumerate(summary.items(), start=1):
        sheet.cell(row=write_start + i, column=1, value=name)
        sheet.cell(row=write_start + i, column=2, value=amount)
        result.append((name, amount))

    wb.save(filepath)
    wb.close()
    return result


def _build_lookup_map(wb, lookup_sheet_name, max_rows=5000):
    """
    在 '供应商信息表（定稿）' 中构建查找映射：
      A 列 = 供应商名称（查找键）
      B 列 = 收款姓名
    :return: dict {供应商名称: 收款姓名}
    """
    if lookup_sheet_name not in wb.sheetnames:
        return {}
    lookup_sheet = wb[lookup_sheet_name]
    mapping = {}
    for row in range(2, min(lookup_sheet.max_row + 1, max_rows + 1)):
        key = lookup_sheet.cell(row=row, column=1).value
        val = lookup_sheet.cell(row=row, column=2).value
        if key is not None:
            mapping[key] = val
    return mapping


def _should_highlight(receiver_name):
    """
    判断收款姓名是否包含需要标黄的关键字
    """
    if not receiver_name:
        return False
    receiver_str = str(receiver_name)
    return any(kw in receiver_str for kw in HIGHLIGHT_KEYWORDS)


def _apply_row_highlight(sheet, row, start_col, end_col):
    """
    将指定行的 start_col 到 end_col 单元格背景设为黄色
    """
    for col in range(start_col, end_col + 1):
        cell = sheet.cell(row=row, column=col)
        cell.fill = YELLOW_FILL


def _split_amount(M, P=4995):
    """
    将金额 M 拆分为多份，每份不超过 P
    :return: [第1份, 第2份, ...]  每份 <= P 且 > 0
    """
    M = float(M)
    if M <= 0:
        return []
    # 拆分行数 = ceil(M / P)
    k = int((M + P - 1) // P)
    first = M - (k - 1) * P
    return [first] + [P] * (k - 1)


def _write_vlookup_formulas(sheet, row, supplier_col, bank_card_col, receiver_col,
                            id_card_col, phone_col, b_letter):
    """
    为指定行写入 VLOOKUP 公式
    """
    sheet.cell(row=row, column=bank_card_col,
               value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,4,0)")
    sheet.cell(row=row, column=receiver_col,
               value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,2,0)")
    sheet.cell(row=row, column=id_card_col,
               value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,3,0)")
    sheet.cell(row=row, column=phone_col,
               value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,5,0)")


def process_second_excel(filepath, summary_data, sheet_name):
    """
    处理第二个 Excel：
      1. 将 summary_data 追加到末尾（不覆盖原有数据）
      2. 根据收款姓名判断规则，处理 H 列"打款金额"
         - 含关键字：C列金额复制到H列
         - 不含关键字且 < 5000：C列金额复制到H列
         - 不含关键字且 >= 5000：拆分金额到多行
      3. 银行卡号 / 收款姓名 / 身份证号 / 电话号码 写 VLOOKUP 公式
      4. 若收款姓名含关键字则 B~L 列标黄
      5. 保存文件
    :param summary_data: [(介绍人, 金额), ...]
    """
    if not filepath.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("仅支持 .xlsx 或 .xlsm 格式")

    wb = openpyxl.load_workbook(filepath)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"工作簿 '{sheet_name}' 不存在")
    sheet = wb[sheet_name]

    keywords = [
        "供应商名称", "打款金额", "银行卡号",
        "收款姓名", "身份证号", "电话号码"
    ]
    header_row, found = _scan_headers(sheet, keywords)
    if header_row is None:
        wb.close()
        raise ValueError(f"未在首 3 行找到必要表头：{keywords}")

    supplier_col = found["供应商名称"][0]

    # 处理两列"打款金额"：最左为 C 列（目标），另一列为 H 列
    dk_cols = sorted(found.get("打款金额", []))
    if len(dk_cols) < 2:
        wb.close()
        raise ValueError("需要两列 '打款金额'（C 列和 H 列），但只找到一列")
    dk_c_col = dk_cols[0]   # 粘贴金额
    dk_h_col = dk_cols[1]   # H列打款金额

    bank_card_col = found["银行卡号"][0]
    receiver_col = found["收款姓名"][0]
    id_card_col = found["身份证号"][0]
    phone_col = found["电话号码"][0]

    # 找原有数据最后一行
    last_row = _find_last_data_row(sheet, header_row, [supplier_col, dk_c_col])

    b_letter = get_column_letter(supplier_col)

    # 预加载供应商信息表映射（A列=供应商名称，B列=收款姓名）
    lookup_map = _build_lookup_map(wb, "供应商信息表（定稿）")

    # 标黄范围：B列(2) 到 L列(12)
    HIGHLIGHT_START_COL = 2
    HIGHLIGHT_END_COL = 12

    current_row = last_row

    for name, amount in summary_data:
        # 查询该供应商的收款姓名，判断规则
        receiver_name = lookup_map.get(name)
        need_highlight = _should_highlight(receiver_name)

        if need_highlight:
            # ========== 规则1：含关键字 ==========
            current_row += 1
            sheet.cell(row=current_row, column=supplier_col, value=name)
            sheet.cell(row=current_row, column=dk_c_col, value=amount)
            sheet.cell(row=current_row, column=dk_h_col, value=amount)
            _write_vlookup_formulas(sheet, current_row, supplier_col, bank_card_col,
                                    receiver_col, id_card_col, phone_col, b_letter)
            _apply_row_highlight(sheet, current_row, HIGHLIGHT_START_COL, HIGHLIGHT_END_COL)

        else:
            # ========== 规则2：不含关键字 ==========
            if amount < 5000:
                # ---- 规则2.1：金额 < 5000，直接复制 ----
                current_row += 1
                sheet.cell(row=current_row, column=supplier_col, value=name)
                sheet.cell(row=current_row, column=dk_c_col, value=amount)
                sheet.cell(row=current_row, column=dk_h_col, value=amount)
                _write_vlookup_formulas(sheet, current_row, supplier_col, bank_card_col,
                                        receiver_col, id_card_col, phone_col, b_letter)
            else:
                # ---- 规则2.2：金额 >= 5000，拆分 ----
                parts = _split_amount(amount)
                for idx, part in enumerate(parts):
                    current_row += 1
                    if idx == 0:
                        # 原行：保留供应商名称和C列金额
                        sheet.cell(row=current_row, column=supplier_col, value=name)
                        sheet.cell(row=current_row, column=dk_c_col, value=amount)
                    else:
                        # 拆分行：供应商名称加后缀，C列为空
                        sheet.cell(row=current_row, column=supplier_col, value=f"{name}{idx}")
                        # C列留空
                    sheet.cell(row=current_row, column=dk_h_col, value=part)
                    _write_vlookup_formulas(sheet, current_row, supplier_col, bank_card_col,
                                            receiver_col, id_card_col, phone_col, b_letter)

    wb.save(filepath)
    wb.close()
