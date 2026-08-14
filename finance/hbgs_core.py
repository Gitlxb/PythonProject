# -*- coding: utf-8 -*-
"""hbgs_core.py - 合并个税核心逻辑 (VBA "个税汇总工具_项目提取_改版" Python 重写)

功能对应:
  VBA ListSubFolders/ListFileNames/汇总  ->  scan_excel_files + run_merge_collect
  VBA LoadFeeRules                       ->  load_fee_rules
  VBA ComputeFee_V2                      ->  compute_fee_v2
  VBA ClassifyRuleExact                  ->  classify_rule_exact
  VBA IsBankExempt                       ->  is_bank_exempt
  VBA ResolveAmount                      ->  resolve_amount
  VBA HandleDualUnitRule                 ->  handle_dual_unit_rule
  VBA CleanProjectName/RemoveSpaces/ExtractMonthFromFileName/NormalizeName -> 同名函数

依赖: openpyxl (必须)
"""

import os
import re
import threading

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


# ============================================================
# 文本工具函数
# ============================================================

def remove_spaces(s):
    """去除半角空格与全角空格(与 VBA RemoveSpaces 等价, 仅去空格不含其他字符)"""
    if s is None:
        return ""
    return re.sub(r"[\u3000 ]", "", str(s))


def normalize_name(s):
    """姓名归一化: 去空格, 全角数字/字母转半角 (与 VBA NormalizeName 等价)"""
    if s is None:
        return ""
    s = str(s).strip()
    out = []
    for ch in s:
        if ch == " " or ch == "\u3000":
            continue
        o = ord(ch)
        if 0xFF10 <= o <= 0xFF19:          # 全角数字
            out.append(chr(o - 0xFF10 + 0x30))
        elif 0xFF21 <= o <= 0xFF3A:        # 全角大写字母
            out.append(chr(o - 0xFF21 + 0x41))
        elif 0xFF41 <= o <= 0xFF5A:        # 全角小写字母
            out.append(chr(o - 0xFF41 + 0x61))
        else:
            out.append(ch)
    return "".join(out)


def clean_project_name(raw):
    """清洗"项目"长文本: 去年月/工资表/单位元等 (与 VBA CleanProjectName 等价)"""
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""
    s = re.sub(r"^\d{2,4}年\d{1,2}月份?", "", s)
    s = re.sub(r"工资表", "", s)
    s = re.sub(r"单位：元", "", s)
    s = re.sub(r"\s*（", "（", s)
    s = re.sub(r"\s*\(", "(", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def extract_month_from_filename(fname):
    """从文件名提取月份数字 (与 VBA ExtractMonthFromFileName 等价)"""
    m = re.search(r"(\d+)月", fname)
    if not m:
        return ""
    return m.group(1)


# ============================================================
# 手续费规则分类与计算
# ============================================================

# A~M 各类精确匹配串 (必须与 VBA 原值逐字一致, 含全角标点)
_RULE_SETS = {
    "A": {
        "只有代工费", "只有管理费，没有工资", "只有代工费，没有工资", "不扣手续费",
        "日结", "联动代发不收手续费", "不发工资，没有手续费", "不扣手续费，锦途发放",
        "不收手续费", "都不扣",
    },
    "B": {"除招商都扣", "招商", "招商不收", "除招商外都扣"},
    "C": {
        "除招商、建设外都扣，离职员工不扣，前15个员工不扣", "除招商和建设外都扣",
        "招商、建设不收", "除招商、建设外都扣", "招商、建设不扣，其他银行都扣",
    },
    "D": {"除平安外都扣"},
    "E": {"除宁波外都扣"},
    "F": {"除工商都扣"},
    "G": {"除温州银行外都扣手续费", "除温州银行外都收手续费"},
    "H": {"邮政卡不收手续费（除邮政都收）", "除了邮政之外的手续费员工承担"},
    "I": {"除邮政银行，其他银行都扣"},
    "J": {"招商和邮政不扣手续费", "除招商、邮政外都扣", "招商、邮储不扣，其他银行都扣"},
    "K": {"除邮政，中信外都扣"},
    "L": {"中信不扣，其他都扣"},
    "M": {"签浙江锦途：招商、建设不扣；签温州驰锦：民生不扣"},
}


def classify_rule_exact(rule_text):
    """将规则文本精确匹配到类别 A~M; 未命中返回空串 (与 VBA ClassifyRuleExact 等价)"""
    s = remove_spaces(rule_text)
    for cat, sset in _RULE_SETS.items():
        if s in sset:
            return cat
    return ""


def is_bank_exempt(category, bank_name):
    """按类别 + 开户行(包含匹配)判断是否免扣 (与 VBA IsBankExempt 等价)"""
    b = remove_spaces(bank_name)
    if category == "A":
        return True
    if category == "B":
        return "招商银行" in b
    if category == "C":
        return ("招商银行" in b) or ("建设银行" in b)
    if category == "D":
        return "平安银行" in b
    if category == "E":
        return "宁波银行" in b
    if category == "F":
        return "工商" in b
    if category == "G":
        return "温州银行" in b
    if category in ("H", "I"):
        return ("邮政" in b) or ("邮储" in b)
    if category == "J":
        return ("招商银行" in b) or ("邮政" in b) or ("邮储" in b)
    if category == "K":
        return ("邮政" in b) or ("邮储" in b) or ("中信银行" in b)
    if category == "L":
        return "中信银行" in b
    return False


def resolve_amount(table_amount, category):
    """解析最终金额: 表格有效正数优先, 否则按类别默认 (A=0, I=3, 其余=5)"""
    if table_amount is not None and table_amount != "":
        try:
            num = float(table_amount)
            if num > 0:
                return int(num)
        except (ValueError, TypeError):
            pass
    if category == "A":
        return 0
    if category == "I":
        return 3
    return 5


def handle_dual_unit_rule(pay_unit, bank, table_amount):
    """M 类双单位特殊规则 (与 VBA HandleDualUnitRule 等价)"""
    pu = remove_spaces(pay_unit)
    b = remove_spaces(bank)
    amt = 5
    if table_amount is not None and table_amount != "":
        try:
            tn = float(table_amount)
            if tn > 0:
                amt = int(tn)
        except (ValueError, TypeError):
            pass
    if pu == "浙江锦途人力资源有限公司":
        return 0 if ("招商银行" in b or "建设银行" in b) else amt
    if pu == "温州驰锦企业服务有限公司":
        return 0 if "民生" in b else amt
    # 兜底: 招商/建行不扣
    return 0 if ("招商银行" in b or "建设银行" in b) else amt


def _is_invalid_name_id(v):
    """姓名/身份证无效判定: 空 / "0" / 错误值"""
    if v is None:
        return True
    s = str(v).strip()
    if s == "" or s == "0":
        return True
    if s.startswith("#"):
        return True
    return False


def compute_fee_v2(name_val, id_val, pay_unit_val, info_lookup, fee_rules, missing_ops=None):
    """逐行计算手续费 (与 VBA ComputeFee_V2 等价)

    info_lookup: dict{ normalized_name: [(ops, bank), ...] }
    fee_rules:   dict{ 运营项目: [扣费银行规则文本, 扣费金额] }
    missing_ops: 可选 set, 用于收集"源文件有但规则表缺"的运营项目(对应 VBA gMissingOps)
    返回: 0 / 3 / 5 / 其他金额, 或 None(留空)
    """
    if _is_invalid_name_id(name_val) or _is_invalid_name_id(id_val):
        return None
    if not fee_rules:
        return None
    if not info_lookup:
        return None
    key = normalize_name(name_val)
    entries = info_lookup.get(key)
    if not entries:
        return None
    for ops, bank in entries:
        if ops is None:
            continue
        base_ops = str(ops).strip()
        if base_ops == "":
            continue
        # "代发Xxx" 去掉前缀后按基础运营项目查规则
        if base_ops.startswith("代发"):
            base_ops = base_ops[2:]
        if base_ops not in fee_rules:
            # 对应 VBA gMissingOps: 源文件有此运营项目但规则表无该行
            if missing_ops is not None:
                missing_ops.add(base_ops)
            continue
        rule_text, table_amount = fee_rules[base_ops]
        category = classify_rule_exact(rule_text)
        if category == "":
            # 与最新 VBA 一致: 规则文本无法识别时直接留空, 不静默扣5元, 也不再收集
            continue
        if category == "M":
            return handle_dual_unit_rule(pay_unit_val, bank, table_amount)
        if is_bank_exempt(category, bank):
            return 0
        return resolve_amount(table_amount, category)
    return None


# ============================================================
# 规则表加载
# ============================================================

def load_fee_rules(path, on_log=None):
    """从规则 Excel 加载规则字典 (与 VBA LoadFeeRules 等价)

    自动选择名字含"手续费"的工作表, 否则取第一个;
    动态定位 客户状态/运营项目/扣费银行（多填）/扣费金额 列;
    跳过"客户状态=终止合作"; 重复运营项目以最后一条为准。
    返回: dict{ 运营项目: [扣费银行规则文本, 扣费金额] }
    """
    if on_log is None:
        on_log = lambda m: None
    wb = load_workbook(path, data_only=True, read_only=True)
    sheet_name = None
    for sn in wb.sheetnames:
        if "手续费" in sn:
            sheet_name = sn
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
    on_log(f"规则表使用工作表: {sheet_name}")
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return {}
    header = rows[0]
    col_status = col_ops = col_bank = col_amount = 0
    for i, h in enumerate(header, start=1):
        ht = remove_spaces(h) if h is not None else ""
        if ht == "客户状态":
            col_status = i
        elif ht == "运营项目":
            col_ops = i
        elif ht in ("扣费银行（多填）", "扣费银行(多填)"):
            col_bank = i
        elif ht == "扣费金额":
            col_amount = i
    if col_ops == 0 or col_bank == 0:
        on_log("⚠ 规则表缺少「运营项目」或「扣费银行（多填）」列, 无法加载")
        return {}
    rules = {}
    for r in rows[1:]:
        if col_status > 0:
            st = r[col_status - 1] if len(r) >= col_status else None
            if st is not None and remove_spaces(str(st)) == "终止合作":
                continue
        ops_val = r[col_ops - 1] if len(r) >= col_ops else None
        if ops_val is None or str(ops_val).strip() == "":
            continue
        ops_key = str(ops_val).strip()
        bank_val = r[col_bank - 1] if len(r) >= col_bank else None
        amount_val = r[col_amount - 1] if (col_amount > 0 and len(r) >= col_amount) else None
        rules[ops_key] = [bank_val, amount_val]
    on_log(f"已加载手续费规则 {len(rules)} 条")
    return rules


# ============================================================
# 文件扫描
# ============================================================

def scan_excel_files(root, recursive=True, exclude_paths=None):
    """扫描 Excel 文件列表

    recursive=True 递归子目录 (默认); 否则仅根目录一层。
    排除: ~$ 临时文件, 非 .xlsx/.xls, exclude_paths 中的路径。
    """
    exclude = set()
    if exclude_paths:
        exclude.update(os.path.normcase(os.path.abspath(p)) for p in exclude_paths)

    result = []

    def _should_skip(full):
        fn = os.path.basename(full)
        if fn.startswith("~$"):
            return True
        if not fn.lower().endswith((".xlsx", ".xls")):
            return True
        if os.path.normcase(os.path.abspath(full)) in exclude:
            return True
        return False

    if recursive:
        for dirpath, _dirnames, filenames in os.walk(root):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                if not _should_skip(full):
                    result.append(full)
    else:
        for fn in os.listdir(root):
            full = os.path.join(root, fn)
            if not os.path.isfile(full):
                continue
            if not _should_skip(full):
                result.append(full)
    return result


# ============================================================
# 读取单个个税工作簿
# ============================================================

def _read_payroll_file(fpath):
    """读取单个工作簿的个税统计表 + 员工信息表 (纯 openpyxl, data_only)

    返回: (src_headers, src_rows, col_name, col_id, col_gjj, col_pay,
           info_lookup, info_found, warns)
    若无"个税统计"表返回 None; 若缺关键列抛 ValueError。
    """
    wb = None
    try:
        wb = load_workbook(fpath, data_only=True, read_only=True)

        # ---- 找"个税统计"表 ----
        target = None
        for sn in wb.sheetnames:
            if remove_spaces(sn) == "个税统计":
                target = sn
                break
        if target is None:
            wb.close()
            return None
        ws = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # ---- 找"员工信息-X月"表 ----
        info_sheet = None
        month = extract_month_from_filename(os.path.basename(fpath))
        for sn in wb.sheetnames:
            if month and remove_spaces(sn) == ("员工信息-" + month + "月"):
                info_sheet = sn
                break
        if info_sheet is None:
            for sn in wb.sheetnames:
                if remove_spaces(sn).startswith("员工信息-"):
                    info_sheet = sn
                    break

        info_lookup = {}
        info_found = info_sheet is not None
        warns = []
        if info_sheet is not None:
            iws = wb[info_sheet]
            irows = list(iws.iter_rows(values_only=True))
            if irows:
                iheader = irows[0]
                ci = co = cb = 0
                for i, h in enumerate(iheader, start=1):
                    ht = remove_spaces(h) if h is not None else ""
                    if ht == "姓名":
                        ci = i
                    elif ht == "运营项目":
                        co = i
                    elif ht == "开户行":
                        cb = i
                if ci > 0 and co > 0 and cb > 0:
                    for r in irows[1:]:
                        nm = r[ci - 1] if len(r) >= ci else None
                        if nm is None or str(nm).strip() == "":
                            continue
                        ops = r[co - 1] if len(r) >= co else None
                        bank = r[cb - 1] if len(r) >= cb else None
                        info_lookup.setdefault(normalize_name(nm), []).append((ops, bank))
                else:
                    warns.append("⚠ 员工信息表缺少 姓名/运营项目/开户行 列")

        if not rows:
            wb.close()
            return None
        header = list(rows[0])

        # ---- 实际列数(最后有数据的列) ----
        src_cols = 0
        for r in rows:
            for j in range(len(r) - 1, -1, -1):
                if r[j] is not None and str(r[j]).strip() != "":
                    if j + 1 > src_cols:
                        src_cols = j + 1
                    break
        if src_cols == 0:
            src_cols = len(header)

        # ---- 定位关键列 ----
        col_name = col_id = col_gjj = col_pay = 0
        for i, h in enumerate(header, start=1):
            if i > src_cols:
                break
            ht = remove_spaces(h) if h is not None else ""
            if ht == "姓名":
                col_name = i
            elif ht == "身份证":
                col_id = i
            elif ht == "公积金":
                col_gjj = i
            elif ht == "发薪单位":
                col_pay = i

        if col_gjj == 0:
            wb.close()
            raise ValueError("表头缺少「公积金」列")
        if col_name == 0 or col_id == 0:
            wb.close()
            raise ValueError("表头缺少「姓名」或「身份证」列")

        src_headers = [(header[i - 1] if i - 1 < len(header) else None)
                       for i in range(1, src_cols + 1)]

        src_rows = []
        for r in rows[1:]:
            rr = list(r)
            if len(rr) < src_cols:
                rr = rr + [None] * (src_cols - len(rr))
            else:
                rr = rr[:src_cols]
            src_rows.append(rr)

        wb.close()
        return (src_headers, src_rows, col_name, col_id, col_gjj, col_pay,
                info_lookup, info_found, warns)
    except Exception:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        raise


def write_merge_output(output_path, header, all_rows, target_id_col):
    """写出汇总表, 身份证列文本化"""
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for out in all_rows:
        row_copy = list(out)
        if target_id_col - 1 < len(row_copy):
            v = row_copy[target_id_col - 1]
            row_copy[target_id_col - 1] = "" if v is None else str(v)
        ws.append(row_copy)
    if 0 < target_id_col <= len(header):
        col_letter = get_column_letter(target_id_col)
        for cell in ws[col_letter]:
            cell.number_format = "@"
    wb.save(output_path)


# ============================================================
# 主流程
# ============================================================

def run_merge_collect(scan_root, rules_path, cancel_event=None,
                      on_log=None, on_progress=None, on_done=None):
    """合并个税 - 阶段一: 扫描 + 读取 + 内存汇总(不写文件)

    参数:
      scan_root    扫描根目录(始终递归子目录)
      rules_path   手续费规则表路径
      cancel_event threading.Event, 置位则中止
      on_log/on_progress/on_done 回调
    结果经 on_done 返回 dict:
      header        输出表头(list)
      rows          汇总数据(list[list])
      target_id_col 身份证列在 header 中的 1-based 位置
      total_files / total_rows / skipped / errors / empty_keyword_ops / missing_ops
      output_path   固定为 None (本阶段不写文件, 由 GUI 弹出保存对话框后再调 write_merge_output)
    """
    if on_log is None:
        on_log = lambda m: None
    if on_progress is None:
        on_progress = lambda c, t, m: None
    if on_done is None:
        on_done = lambda r: None
    if cancel_event is None:
        cancel_event = threading.Event()

    result = {"header": None, "rows": [], "target_id_col": None,
              "total_files": 0, "total_rows": 0, "skipped": 0,
              "errors": [], "empty_keyword_ops": set(), "missing_ops": set(),
              "output_path": None}

    # ---- 加载规则 ----
    try:
        fee_rules = load_fee_rules(rules_path, on_log)
    except Exception as e:
        on_log(f"⚠ 读取规则表失败: {e}")
        on_done(result)
        return
    if not fee_rules:
        on_log("⚠ 未加载到任何手续费规则, 手续费列将全部留空")

    # ---- 校验：规则表中"扣费银行（多填）"列是否为空 (对应 VBA gEmptyKeywordOps) ----
    empty_keyword_ops = result["empty_keyword_ops"]
    for ops_key, entry in fee_rules.items():
        rule_text = entry[0]
        if rule_text is None or remove_spaces(str(rule_text)) == "":
            empty_keyword_ops.add(ops_key)

    # ---- 扫描文件(始终递归) ----
    file_list = scan_excel_files(scan_root, recursive=True,
                                  exclude_paths=[rules_path])
    if not file_list:
        on_log("未找到任何 Excel 文件")
        on_done(result)
        return

    result["total_files"] = len(file_list)
    missing_ops = result["missing_ops"]

    # ---- 首文件模板 + 按列名映射(对齐 VBA: 表头以首文件为准, 后续文件按列名映射, 多余列丢弃) ----
    # VBA 原始行为: 输出表头完全由首个有效文件决定, 后续文件多出的列不加入输出。
    # 改进点: 不再要求后续文件列顺序/数量与首文件一致(否则大面积跳过),
    #         而是按"归一化列名"做 name-based 映射到首文件模板。
    template_headers = None   # 首文件的 src_headers (含第0列项目)
    template_gjj = 0          # 首文件的公积金列号(1-based)
    file_data = []            # 收集每个文件的数据

    for idx, fpath in enumerate(file_list, start=1):
        if cancel_event.is_set():
            on_log("⚠ 用户取消, 停止处理")
            break
        on_progress(idx, len(file_list), f"读取中: {os.path.basename(fpath)}")
        try:
            data = _read_payroll_file(fpath)
        except Exception as e:
            msg = f"✗ 无法读取 {os.path.basename(fpath)}: {e}"
            on_log(msg)
            result["errors"].append(msg)
            result["skipped"] += 1
            continue
        if data is None:
            msg = f"✗ {os.path.basename(fpath)} 无「个税统计」表, 已跳过"
            on_log(msg)
            result["errors"].append(msg)
            result["skipped"] += 1
            continue

        (src_headers, src_rows, col_name, col_id, col_gjj, col_pay,
         info_lookup, info_found, warns) = data
        for w in warns:
            on_log(w)
        if not info_found:
            on_log(f"⚠ {os.path.basename(fpath)} 无员工信息表, 手续费将留空")

        # 首个有效文件 → 锁定模板
        if template_headers is None:
            template_headers = src_headers
            template_gjj = col_gjj

        file_data.append({
            "fname": os.path.splitext(os.path.basename(fpath))[0],
            "src_headers": src_headers,
            "src_rows": src_rows,
            "cols": (col_name, col_id, col_gjj, col_pay),
            "info_lookup": info_lookup,
        })

    if not file_data or template_headers is None:
        on_log("没有有效数据可汇总！")
        on_done(result)
        return

    # ---- 构建输出表头(与 VBA 一致): 文件名/项目 + 首文件除第1列外的列(公积金后插手续费) + 采用/不采用 ----
    # 关键列 -> 稳定 key: 有列名用归一名, 无列名(空白表头)用位置占位 __pos_N__ (N=模板中的列序号)
    # 这样空白表头列也能按位置正确取数, 不会丢数据。
    def _col_key(header_text, pos):
        norm = remove_spaces(header_text)
        return norm if norm != "" else f"__pos_{pos}__"

    out_header = ["文件名", "项目"]
    out_keys = ["_fname", "_proj"]
    for i in range(1, len(template_headers)):
        out_header.append(template_headers[i])
        out_keys.append(_col_key(template_headers[i], i))
        if remove_spaces(template_headers[i]) == "公积金":
            out_header.append("手续费")
            out_keys.append("_fee")
    out_header.append("采用/不采用")
    out_keys.append("_tag")

    # 用归一名查找(兼容表头中的空格/全角空格), 与 _read_payroll_file 保持一致
    name_out_idx = next((i for i, h in enumerate(out_header) if remove_spaces(h) == "姓名"), None)
    id_out_idx = next((i for i, h in enumerate(out_header) if remove_spaces(h) == "身份证"), None)
    target_id_col = (id_out_idx + 1) if id_out_idx is not None else None

    # ---- 逐文件映射数据 ----
    all_rows = []
    for fd in file_data:
        fname = fd["fname"]
        src_headers = fd["src_headers"]
        src_rows = fd["src_rows"]
        col_name, col_id, col_gjj, col_pay = fd["cols"]
        info_lookup = fd["info_lookup"]

        # 本文件: 稳定 key -> 源行 0-based 索引 (空白列用 __pos_N__ 位置占位)
        src_key_to_idx = {}
        for i in range(len(src_headers)):
            src_key_to_idx[_col_key(src_headers[i], i)] = i

        for row in src_rows:
            proj = clean_project_name(row[0] if len(row) > 0 else None)
            out = [fname, proj]
            fee_computed = False
            # 遍历除 文件名/项目(前2) 和 采用/不采用(末列) 外的输出列
            for h, key in zip(out_header[2:-1], out_keys[2:-1]):
                if key == "_fee":
                    if not fee_computed:
                        name_v = row[col_name - 1] if (col_name > 0 and col_name - 1 < len(row)) else None
                        id_v = row[col_id - 1] if (col_id > 0 and col_id - 1 < len(row)) else None
                        pay_v = row[col_pay - 1] if (col_pay > 0 and col_pay - 1 < len(row)) else None
                        fee = compute_fee_v2(name_v, id_v, pay_v, info_lookup, fee_rules, missing_ops)
                        out.append(fee)
                        fee_computed = True
                    continue
                src_idx = src_key_to_idx.get(key)
                val = row[src_idx] if (src_idx is not None and src_idx < len(row)) else None
                out.append(val)

            nv = out[name_out_idx] if (name_out_idx is not None and name_out_idx < len(out)) else None
            iv = out[id_out_idx] if (id_out_idx is not None and id_out_idx < len(out)) else None
            tag = "不采用" if (_is_invalid_name_id(nv) or _is_invalid_name_id(iv)) else "采用"
            out.append(tag)
            all_rows.append(out)

    result["header"] = out_header
    result["rows"] = all_rows
    result["target_id_col"] = target_id_col
    result["total_rows"] = len(all_rows)
    on_log(f"✔ 内存汇总完成: 共 {len(all_rows)} 条数据, 请在弹出的对话框中选择保存位置")

    # ---- 提示汇总①：规则表"扣费银行（多填）"列为空 (对应 VBA gEmptyKeywordOps) ----
    if empty_keyword_ops:
        e_list = " ".join("[" + str(o) + "]" for o in sorted(empty_keyword_ops))
        on_log('【规则表提示】以下运营项目的"扣费银行（多填）"列为空，需要填写关键词：' + e_list)

    # ---- 提示汇总②：源文件有但规则表缺的运营项目 (对应 VBA gMissingOps) ----
    if missing_ops:
        m_list = " ".join("[" + str(o) + "]" for o in sorted(missing_ops))
        on_log('【项目提示】以下运营项目在规则表中找不到，请在"收取银行手续费信息"工作表中填写运营项目及文本关键词：' + m_list)

    on_done(result)
