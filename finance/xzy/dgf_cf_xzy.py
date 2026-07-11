"""
核心逻辑模块：Excel 数据处理
与 GUI 解耦，可单独导入调用
"""

import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment
from collections import defaultdict

# 标黄填充样式
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# #7cded7 填充样式（刘先锋现金卡支付）
STEEL_BLUE_FILL = PatternFill(start_color="7cded7", end_color="7cded7", fill_type="solid")

# 居中对齐样式
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

# 需要标黄判断的关键字（公司、个体工商户、个体户、服务部）
HIGHLIGHT_KEYWORDS = ["公司", "个体工商户", "个体户", "服务部"]

# 发薪单位/公司全称 → 简写前缀映射
COMPANY_NAME_MAP = {
    "刘先锋个人现金卡": "刘先锋现金卡支付",
    "温州鑫锦途供应链服务有限公司": "温州鑫锦途",
    "浙江锦途人力资源有限公司": "浙江锦途",
    "温州锦途服务外包有限公司": "温州锦途外包",
    "衢州鑫途服务外包有限公司": "衢州鑫途外包",
    "浙江锦途人力资源有限公司乐清分公司": "乐清分公司",
    "芜湖才库人力资源有限公司": "芜湖才库",
    "衢州驰锦企业服务有限公司": "衢州驰锦",
    "安徽锦途企业服务集团有限公司": "安徽锦途服务",
    "安徽锦途企业管理咨询有限公司": "安徽锦途",
    "安庆市同益劳务服务有限公司": "安庆同益",
    "台州众通人才服务有限公司": "台州众通",
    "台州锦途人力资源服务外包有限公司": "台州锦途外包",
    "温州鑫途企业服务有限公司": "温州鑫途",
    "温州驰锦企业服务有限公司": "温州驰锦",
    "浙江盛威安保服务有限公司": "盛威安保",
    "温州市合静人力资源有限公司": "合静",
    "铜陵锦途劳务有限公司": "铜陵锦途",
    "安徽锦途人力资源有限公司": "安徽人力",
}


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


def _normalize_header(val):
    """标准化表头文本：去除首尾空格、内部空格，半角冒号转全角"""
    if val is None:
        return ""
    return str(val).strip().replace(":", "：").replace(" ", "").replace("　", "")


def _scan_headers(sheet, keywords, max_check_rows=3):
    """
    扫描 sheet 的前 max_check_rows 行，寻找包含所有 keywords 的表头行
    :return: (header_row, {keyword: [col1, col2, ...]})
             若未找到返回 (None, {})
    """
    # 标准化关键字（统一为全角冒号、去除空格）
    norm_keywords = [_normalize_header(k) for k in keywords]
    keywords_set = set(norm_keywords)
    # 建立 标准化后 -> 原始关键字 映射，用于返回结果
    norm_to_orig = {_normalize_header(k): k for k in keywords}

    limit_row = min(max_check_rows + 1, sheet.max_row + 1)
    for row_idx in range(1, limit_row):
        found = {}
        for col_idx in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            norm_val = _normalize_header(val)
            if norm_val in keywords_set:
                orig_key = norm_to_orig[norm_val]
                found.setdefault(orig_key, []).append(col_idx)
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
      1. 按"介绍人"汇总"金额"
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

    header_row, found = _scan_headers(sheet_read, ["介绍人", "金额"])
    if header_row is None:
        wb_read.close()
        raise ValueError("未在首 3 行找到表头：'介绍人' 和 '金额'")

    jieshao_col = found["介绍人"][0]
    heji_col = found["金额"][0]

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
        raise ValueError("未找到有效数据（介绍人或金额为空）")

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


def get_company_prefix(filepath, sheet_name="个税统计"):
    """
    从代工费表格的'个税统计'工作簿中读取发薪单位/发薪公司/公司，
    按 COMPANY_NAME_MAP 映射简写后返回前缀。
    若未找到工作簿或匹配关键词的单元格，返回空字符串。
    
    关键词匹配方式：子串包含（非精确表头匹配）
    - 检测到包含"发薪单位"四个字的单元格
    - 检测到包含"发薪公司"四个字的单元格
    - 检测到包含"公司"两个字的单元格
    """
    if not filepath.lower().endswith((".xlsx", ".xlsm")):
        return ""

    wb = openpyxl.load_workbook(filepath)
    try:
        if sheet_name not in wb.sheetnames:
            return ""
        sheet = wb[sheet_name]

        # 扫描第1行，找包含关键词的单元格（子串匹配）
        keywords = ["发薪单位", "发薪公司", "公司"]
        header_row = None
        col = None

        for col_idx in range(1, sheet.max_column + 1):
            val = sheet.cell(row=1, column=col_idx).value
            if val is not None:
                text = str(val).strip()
                for kw in keywords:
                    if kw in text:
                        header_row = 1
                        col = col_idx
                        break
                if header_row is not None:
                    break

        if header_row is None or col is None:
            return ""

        # 读取第一个非空值并映射
        for row in range(header_row + 1, sheet.max_row + 1):
            val = sheet.cell(row=row, column=col).value
            if val is not None:
                full_name = str(val).strip()
                return COMPANY_NAME_MAP.get(full_name, full_name)

        return ""
    finally:
        wb.close()


def init_b_state(filepath, sheet_name_b):
    """
    读取工作簿B到内存字典，返回 {收款姓名: 金额}
    不修改文件，仅读取现有数据。
    """
    b_state = {}
    wb = openpyxl.load_workbook(filepath)
    if sheet_name_b in wb.sheetnames:
        sheet = wb[sheet_name_b]
        header_row, found = _scan_headers(sheet, ["收款姓名", "求和项：打款金额"])
        if header_row is not None:
            receiver_col = found["收款姓名"][0]
            amount_col = found["求和项：打款金额"][0]
            last_data_row = _find_last_data_row(sheet, header_row, [receiver_col, amount_col])
            for row in range(header_row + 1, last_data_row + 1):
                r_name = sheet.cell(row=row, column=receiver_col).value
                r_amount = sheet.cell(row=row, column=amount_col).value
                if r_name is not None:
                    try:
                        b_state[r_name] = float(r_amount) if r_amount is not None else 0.0
                    except (ValueError, TypeError):
                        b_state[r_name] = 0.0
    wb.close()
    return b_state


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


def _should_highlight_stop_split(receiver_name):
    """
    判断是否停止拆分：含"公司"/"个体工商户"/"个体户"/"刘先锋现金卡支付"任意一个
    """
    if not receiver_name:
        return False
    receiver_str = str(receiver_name)
    return any(kw in receiver_str for kw in HIGHLIGHT_KEYWORDS + ["刘先锋现金卡支付"])


def _should_highlight(receiver_name):
    """
    判断收款姓名是否包含需要标黄的关键字（公司/个体户，不含刘先锋）
    """
    if not receiver_name:
        return False
    receiver_str = str(receiver_name)
    return any(kw in receiver_str for kw in HIGHLIGHT_KEYWORDS)


def _is_lxf_cash_card(receiver_name):
    """
    判断收款姓名是否为"刘先锋现金卡支付"
    """
    if not receiver_name:
        return False
    return "刘先锋现金卡支付" in str(receiver_name)


def _apply_row_highlight(sheet, row, start_col, end_col, fill=None):
    """
    将指定行的 start_col 到 end_col 单元格背景设为指定颜色
    """
    if fill is None:
        fill = YELLOW_FILL
    for col in range(start_col, end_col + 1):
        cell = sheet.cell(row=row, column=col)
        cell.fill = fill


def _write_vlookup_formulas(sheet, row, supplier_col, bank_card_col, receiver_col,
                            id_card_col, phone_col, b_letter, align=None):
    """
    为指定行写入 VLOOKUP 公式
    """
    cell = sheet.cell(row=row, column=bank_card_col,
                      value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,4,0)")
    if align:
        cell.alignment = align
    cell = sheet.cell(row=row, column=receiver_col,
                      value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,2,0)")
    if align:
        cell.alignment = align
    cell = sheet.cell(row=row, column=id_card_col,
                      value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,3,0)")
    if align:
        cell.alignment = align
    cell = sheet.cell(row=row, column=phone_col,
                      value=f"=VLOOKUP({b_letter}{row},'供应商信息表（定稿）'!$A:$E,5,0)")
    if align:
        cell.alignment = align


def build_appended_records(wb, summary_data):
    """
    根据 summary_data 和"供应商信息表（定稿）"构建追加记录列表。
    不创建/修改任何 sheet，仅做数据查找。
    :param wb: 已打开的工作簿对象（发放记录表格）
    :param summary_data: [(介绍人, 金额), ...] 来自 process_first_excel
    :return: [(介绍人, 金额, 收款姓名), ...]
    """
    lookup_map = _build_lookup_map(wb, "供应商信息表（定稿）")
    appended_records = []
    for name, amount in summary_data:
        receiver_name = lookup_map.get(name)
        appended_records.append((name, amount, receiver_name))
    return appended_records


def write_split_rows(wb, split_records, sheet_name_a, b_state,
                     appended_names=None):
    """
    在工作簿A中回写拆分行：
      按4995阈值动态拆分：优先尝试原名称（填满到4995），剩余逐块分配给后缀。
      每条记录可能生成多行拆分结果。
      后缀名称优先在"供应商信息表（定稿）"中按顺序查找，
      若在信息表中存在且对应收款姓名在 b_state 中累加后≤4995则采用；
      若含关键词（公司/个体户/刘先锋）则停止查找，剩余金额全部给当前后缀。
      同批次内避免重复名称，且不与追加阶段的原始名称冲突。
      含关键词的拆分记录整行标黄。
    :param split_records: [(介绍人, 金额, 收款姓名), ...]
    :param b_state: dict {收款姓名: 金额}，实时更新
    :param appended_names: set() 本次追加阶段已使用的原始供应商名称集合
    :return: ([(新名称, 金额, 新收款姓名), ...], start_row_in_a)
    """
    if sheet_name_a not in wb.sheetnames:
        raise ValueError(f"工作簿 '{sheet_name_a}' 不存在")

    sheet_a = wb[sheet_name_a]

    # ---- 解析工作簿A表头 ----
    keywords_a = [
        "供应商名称", "打款金额", "银行卡号",
        "收款姓名", "身份证号", "电话号码"
    ]
    header_row_a, found_a = _scan_headers(sheet_a, keywords_a)
    if header_row_a is None:
        raise ValueError(f"未在首 3 行找到必要表头：{keywords_a}")

    supplier_col = found_a["供应商名称"][0]
    dk_cols = sorted(found_a.get("打款金额", []))
    if len(dk_cols) < 2:
        raise ValueError("需要两列 '打款金额'（C 列和 H 列），但只找到一列")
    dk_c_col = dk_cols[0]
    dk_h_col = dk_cols[1]

    bank_card_col = found_a["银行卡号"][0]
    receiver_col_a = found_a["收款姓名"][0]
    id_card_col = found_a["身份证号"][0]
    phone_col = found_a["电话号码"][0]

    last_row = _find_last_data_row(sheet_a, header_row_a, [supplier_col, dk_c_col])
    b_letter = get_column_letter(supplier_col)

    HIGHLIGHT_START_COL = 2
    HIGHLIGHT_END_COL = 12

    # ---- 加载供应商信息表映射 ----
    lookup_map = _build_lookup_map(wb, "供应商信息表（定稿）")
    valid_supplier_names = set(lookup_map.keys())

    # used_names：追加阶段的原始名称 + 本次回写已写入的名称
    used_names = set(appended_names) if appended_names else set()

    current_row = last_row
    new_records = []
    LIMIT = 4995  # 阈值

    for original_name, total_amount, original_receiver in split_records:
        remaining = total_amount  # 剩余待分配金额
        first_split = True

        while remaining > 0:
            candidate = None
            chosen_receiver = None
            allocate_amount = 0

            # 原名称（仅第一次尝试）
            if first_split:
                if _should_highlight(original_receiver):
                    # 原收款人含关键词，直接整条给原名称
                    candidate = original_name
                    chosen_receiver = original_receiver
                    allocate_amount = remaining
                    remaining = 0
                else:
                    # 尝试原名称能塞多少
                    orig_current = b_state.get(original_receiver, 0.0)
                    if orig_current < LIMIT:
                        can_alloc = min(LIMIT - orig_current, remaining)
                        if can_alloc > 0:
                            candidate = original_name
                            chosen_receiver = original_receiver
                            allocate_amount = can_alloc
                            remaining -= can_alloc
                            b_state[original_receiver] = orig_current + can_alloc

                first_split = False
                if candidate is not None:
                    # 写入
                    current_row += 1
                    _write_split_row(sheet_a, current_row, supplier_col, dk_c_col,
                                     dk_h_col, bank_card_col, receiver_col_a,
                                     id_card_col, phone_col, b_letter, candidate,
                                     allocate_amount)
                    _apply_highlight(sheet_a, current_row, chosen_receiver,
                                     HIGHLIGHT_START_COL, HIGHLIGHT_END_COL)
                    new_records.append((candidate, allocate_amount, chosen_receiver))
                    used_names.add(candidate)

                    if remaining <= 0:
                        break
                    # 首轮部分分配完毕，剩余金额由下轮while继续处理
                    continue

            # 后缀1~20
            found_suffix = False
            for suffix in range(1, 21):
                test_name = f"{original_name}{suffix}"
                if test_name in used_names:
                    continue
                if test_name not in valid_supplier_names:
                    continue

                new_receiver = lookup_map[test_name]

                # 含关键词：停止，剩余全部给当前后缀
                if _should_highlight_stop_split(new_receiver):
                    candidate = test_name
                    chosen_receiver = new_receiver
                    allocate_amount = remaining
                    remaining = 0
                    found_suffix = True
                    break

                # 检查能否分配
                current_total = b_state.get(new_receiver, 0.0)
                if current_total >= LIMIT:
                    continue  # 已满，跳过

                can_alloc = min(LIMIT - current_total, remaining)
                if can_alloc > 0:
                    candidate = test_name
                    chosen_receiver = new_receiver
                    allocate_amount = can_alloc
                    remaining -= can_alloc
                    b_state[new_receiver] = current_total + can_alloc
                    found_suffix = True
                    break

            if candidate is None:
                # 后缀1~20都失败，回退到原名称（累加已分配+剩余）
                candidate = original_name
                chosen_receiver = original_receiver
                allocate_amount = remaining
                remaining = 0

            # 写入
            current_row += 1
            _write_split_row(sheet_a, current_row, supplier_col, dk_c_col,
                             dk_h_col, bank_card_col, receiver_col_a,
                             id_card_col, phone_col, b_letter, candidate,
                             allocate_amount)
            _apply_highlight(sheet_a, current_row, chosen_receiver,
                             HIGHLIGHT_START_COL, HIGHLIGHT_END_COL)
            new_records.append((candidate, allocate_amount, chosen_receiver))
            used_names.add(candidate)

    return new_records, last_row + 1


def _write_split_row(sheet, row, supplier_col, dk_c_col, dk_h_col,
                     bank_card_col, receiver_col, id_card_col, phone_col,
                     b_letter, candidate, amount):
    """写入拆分行数据"""
    cell = sheet.cell(row=row, column=supplier_col, value=candidate)
    cell.alignment = CENTER_ALIGN
    cell = sheet.cell(row=row, column=dk_c_col, value=amount)
    cell.alignment = CENTER_ALIGN
    cell = sheet.cell(row=row, column=dk_h_col, value=amount)
    cell.alignment = CENTER_ALIGN
    _write_vlookup_formulas(sheet, row, supplier_col, bank_card_col,
                            receiver_col, id_card_col, phone_col, b_letter,
                            CENTER_ALIGN)


def _apply_highlight(sheet, row, chosen_receiver, start_col, end_col):
    """根据收款姓名应用行高亮"""
    if _is_lxf_cash_card(chosen_receiver):
        _apply_row_highlight(sheet, row, start_col, end_col, STEEL_BLUE_FILL)
    elif _should_highlight(chosen_receiver):
        _apply_row_highlight(sheet, row, start_col, end_col, YELLOW_FILL)


def parse_remark1_from_filename(filename):
    """
    从代工费表格文件名解析备注1内容。
    例：'26年3月登高代工费（已核对）-温州鑫锦途.xlsx' -> '3月登高服务费'
    """
    name = filename.rsplit('.', 1)[0]
    name = re.sub(r'^\d+年', '', name)
    name = re.split(r'[（\(-]', name)[0].strip()
    name = name.replace('代工费', '服务费')
    return name


def parse_company_from_filename(filename):
    """
    从代工费表格文件名提取公司名（最后一个'-'之后的内容）。
    例：'26年3月登高代工费（已核对）-温州鑫锦途.xlsx' -> '温州鑫锦途'
    """
    name = filename.rsplit('.', 1)[0]
    if '-' in name:
        return name.split('-')[-1].strip()
    return ""


def fill_split_rows(wb, sheet_name, split_start_row, new_records,
                    remark1, company_prefix):
    """
    为原sheet中的拆分记录填充I~L列备注信息。
    :param split_start_row: 拆分记录在sheet中的起始行号
    :param new_records: [(新名称, 金额, 新收款姓名), ...]
    :param company_prefix: 从表1'个税统计'工作簿读取并简写后的公司前缀
    """
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"工作簿 '{sheet_name}' 不存在")
    sheet = wb[sheet_name]

    keywords = ["供应商名称", "打款金额", "银行卡号",
                "收款姓名", "身份证号", "电话号码",
                "备注1", "备注2", "是否申报个税", "申报个税月份"]
    header_row, found = _scan_headers(sheet, keywords)
    if header_row is None:
        raise ValueError(f"未在首 3 行找到必要表头：{keywords}")

    remark1_col = found["备注1"][0]
    remark2_col = found["备注2"][0]
    tax_declare_col = found["是否申报个税"][0]
    tax_month_col = found["申报个税月份"][0]

    for i, (candidate, amount, chosen_receiver) in enumerate(new_records):
        row = split_start_row + i

        # 判断收款姓名是否含关键词（对公/代发）
        has_kw = _should_highlight(chosen_receiver)
        # 刘先锋现金卡支付特殊处理（基于收款姓名内容）
        is_lxf = chosen_receiver is not None and "刘先锋现金卡支付" in str(chosen_receiver)

        if is_lxf:
            r2 = "刘先锋现金卡支付"
            tax = "刘先锋现金卡支付"
            month = "/"
        else:
            # 备注2生成逻辑
            if has_kw:
                r2 = f"{company_prefix}对公打款" if company_prefix else "对公打款"
            else:
                r2 = f"{company_prefix}代发打款" if company_prefix else "代发打款"

            # 是否申报个税、申报个税月份
            if has_kw:
                tax = "否，对公打款，后续对接开票"
                month = "/"
            else:
                tax = "是"
                month = None

        cell = sheet.cell(row=row, column=remark1_col, value=remark1)
        cell.alignment = CENTER_ALIGN
        cell = sheet.cell(row=row, column=remark2_col, value=r2)
        cell.alignment = CENTER_ALIGN
        cell = sheet.cell(row=row, column=tax_declare_col, value=tax)
        cell.alignment = CENTER_ALIGN
        cell = sheet.cell(row=row, column=tax_month_col, value=month)
        cell.alignment = CENTER_ALIGN
