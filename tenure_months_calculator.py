#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================
员工在职月份计算工具（形态 C：按人分组，主在职扣除短离职）
=====================================================================

【业务规则】
1. 表头精确识别：只识别 5 个固定列名——状态、姓名、身份证号码、入职日期、离职日期。
2. 人员分组键：姓名 + 身份证号码。同名不同身份证视为不同人员，各自独立计算。
3. 行类型：
   - 在职行：状态 = "在职"（离职日期列应为空，程序忽略其离职日期）。
   - 离职行：状态 = "离职"，有入职日期和离职日期。
   - 其他状态：保留原行，在职月份为空、在职天数为 0。
4. 主在职区间：
   - 对每个人（姓名+身份证），取所有在职行的最早入职日。
   - 主在职起点 = max(在职行最早入职日, 该人所有离职行最早入职日) + 起算偏移。
     起算偏移：0 = 含入职日当天；1 = 从入职次日算起（默认）。
   - 主在职终点 = min(统计当月月末, 在职天数截止日 as_of)。
   - 主在职区间按闭区间 [起点, 终点] 计算天数。
5. 短离职区间：对每个人的每条离职行，区间 = [离职行入职起算日, 离职日]。
   起算日与主在职规则一致（默认入职次日）。
6. 主在职扣除短离职：从主在职区间中减去所有短离职区间的并集，得到剩余区间。
   剩余区间按年月聚合，归属到该人的在职行。
7. 入职 = 离职：短离职区间为空，但离职行仍需输出该月份、0 天。
8. 主在职被完全扣除：剩余区间为空，在职行输出在职月份为空、在职天数为 0。

【输出结构 —— 单张工作表】
列 = 原表所有列 + 两个新增列：
  · 在职月份：单值年月标签（如 "2025-01"），表示当事人该月有在职/离职记录。
  · 在职天数：该年-月内实际天数（整数）。
在职行按月展开主在职剩余区间；离职行按月展开自己的短离职区间；其他状态行保留原行。

【参数】
- stat_month：统计截止月份 "YYYY-MM"，决定主在职终点上限和月份标记范围。
- as_of_date：在职天数截止日；None 表示取今天。主在职终点 = min(月末, as_of)。
- include_hire_day：True 表示入职当月含入职日；False（默认）从入职次日算起。
=====================================================================
"""

import calendar
from datetime import datetime, date, timedelta
from typing import List, Tuple, Dict, Optional

import pandas as pd


# ============================ 默认配置区 ============================
CONFIG = {
    "sheet_name": 0,                       # 0 = 第一个工作表；也可填表名

    # 表头精确匹配（不再使用同义词）
    "col_employee": "姓名",
    "col_hire": "入职日期",
    "col_depart": "离职日期",
    "col_id": "身份证号码",
    "col_status": "状态",

    # 统计截止月份：格式 "YYYY-MM"
    "stat_month": "2026-07",

    # 在职天数截止日：None / "today" / 空字符串 都表示今天
    "as_of_date": None,

    # 入职当月是否包含入职当天：False = 从入职次日算起（默认）
    "include_hire_day": False,
}
# ==================================================================


# ------------------------- 基础日期工具 -------------------------
def month_end(y: int, m: int) -> date:
    """返回某年某月的最后一天。"""
    return date(y, m, calendar.monthrange(y, m)[1])


def days_by_month(start: date, end: date) -> Dict[Tuple[int, int], int]:
    """闭区间 [start, end] 按 (year, month) 聚合的天数。start > end 返回空 dict。"""
    result = {}
    if start > end:
        return result
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        dim = calendar.monthrange(y, m)[1]
        seg_start = date(y, m, 1)
        seg_end = date(y, m, dim)
        s = max(seg_start, start)
        e = min(seg_end, end)
        result[(y, m)] = (e - s).days + 1
        if m == 12:
            y, m = y + 1, 1
        else:
            m += 1
    return result


def subtract_interval(
    start: date, end: date, sub_start: date, sub_end: date
) -> List[Tuple[date, date]]:
    """从 [start, end] 中减去 [sub_start, sub_end]，返回剩余闭区间列表。"""
    if sub_end < start or sub_start > end:
        return [(start, end)]
    result = []
    if sub_start > start:
        result.append((start, sub_start - timedelta(days=1)))
    if sub_end < end:
        result.append((sub_end + timedelta(days=1), end))
    return result


def merge_intervals(intervals: List[Tuple[date, date]]) -> List[Tuple[date, date]]:
    """合并重叠或相邻的闭区间。"""
    if not intervals:
        return []
    intervals = sorted(intervals)
    merged = [intervals[0]]
    for s, e in intervals[1:]:
        last_s, last_e = merged[-1]
        if s <= last_e + timedelta(days=1):
            merged[-1] = (last_s, max(last_e, e))
        else:
            merged.append((s, e))
    return merged


# ------------------------- 核心算法 -------------------------
def hire_start_date(hire: date, include_hire_day: bool) -> date:
    """根据 include_hire_day 返回区间起算日。"""
    return hire if include_hire_day else hire + timedelta(days=1)


def compute_main_employment(
    employed_hires: List[date],
    leave_intervals: List[Tuple[date, date]],
    stat_end: date,
    as_of: date,
    include_hire_day: bool,
) -> Dict[Tuple[int, int], int]:
    """
    计算某个人主在职区间扣除短离职后的年月天数。

    参数
    ----
    employed_hires : 所有在职行的入职日期列表
    leave_intervals : 所有离职行的 (入职起算日, 离职日) 列表
    stat_end        : 统计截止月份月末
    as_of           : 在职天数截止日
    include_hire_day: 是否含入职当天
    """
    if not employed_hires:
        return {}

    # 主在职起点：max(在职最早入职, 离职最早入职) + 起算偏移
    earliest_hire = min(employed_hires)
    if leave_intervals:
        earliest_leave_hire = min(ls for ls, _ in leave_intervals)
        base_start = max(earliest_hire, earliest_leave_hire)
    else:
        base_start = earliest_hire

    main_start = hire_start_date(base_start, include_hire_day)
    main_end = min(stat_end, as_of)

    if main_start > main_end:
        return {}

    # 主在职区间
    remaining = [(main_start, main_end)]

    # 从主在职中逐个扣除短离职区间（先合并，再扣除）
    merged_leaves = merge_intervals(leave_intervals)
    for ls, le in merged_leaves:
        if le < main_start or ls > main_end:
            continue
        new_remaining = []
        for s, e in remaining:
            new_remaining.extend(subtract_interval(s, e, ls, le))
        remaining = new_remaining
        if not remaining:
            break

    # 按年月聚合
    result = {}
    for s, e in remaining:
        for k, v in days_by_month(s, e).items():
            result[k] = result.get(k, 0) + v
    return result


def compute_leave_interval(
    hire: date, depart: date, include_hire_day: bool
) -> Tuple[date, date]:
    """返回单个离职行的短离职闭区间 [start, end]；start > end 表示区间为空。"""
    start = hire_start_date(hire, include_hire_day)
    end = depart
    return start, end


# ------------------------- 表头精确识别 -------------------------
REQUIRED_HEADERS = {
    "status": "状态",
    "employee": "姓名",
    "id": "身份证号码",
    "hire": "入职日期",
    "depart": "离职日期",
}


def detect_columns(df: pd.DataFrame) -> dict:
    """精确匹配 5 个固定表头；返回键 -> 列名 或 None。"""
    return {key: name if name in df.columns else None for key, name in REQUIRED_HEADERS.items()}


def to_date(ts):
    """pandas Timestamp/NaT -> date 或 None。"""
    return ts.date() if pd.notna(ts) else None


def person_key(r, col_emp, col_id) -> Tuple[str, str]:
    """构造人员分组键。"""
    name = r[col_emp] if col_emp and col_emp in r.index else None
    id_ = r[col_id] if col_id and col_id in r.index else None
    return (
        str(name) if pd.notna(name) else "",
        str(id_) if pd.notna(id_) else "",
    )


# ------------------------- 主流程 -------------------------
def process(input_file: str, output_file: str, **options):
    """
    计算员工在职月份/天数，并将结果写入 output_file。

    参数
    ----
    input_file  : 输入 Excel 文件路径（必填）
    output_file : 输出 Excel 文件路径（必填）
    **options   : 覆盖 CONFIG 默认值
    """
    # 合并配置
    cfg = dict(CONFIG)
    cfg.update(options)

    # 读 Excel
    df = pd.read_excel(input_file, sheet_name=cfg["sheet_name"])

    # 精确识别表头
    detected = detect_columns(df)
    col_emp = cfg["col_employee"] or detected["employee"]
    col_hire = cfg["col_hire"] or detected["hire"]
    col_dep = cfg["col_depart"] or detected["depart"]
    col_id = cfg["col_id"] or detected["id"]
    col_status = cfg["col_status"] or detected["status"]

    print("=" * 64)
    print("表头识别结果：", detected)
    print("最终使用表头：", {
        "状态": col_status, "姓名": col_emp, "身份证": col_id,
        "入职": col_hire, "离职": col_dep,
    })
    print("统计截止月份：", cfg["stat_month"],
          " | 含入职日：", cfg["include_hire_day"])
    print("=" * 64)

    # 必需列校验
    if not col_status:
        print("❌ 未识别到『状态』列，请检查表头是否为『状态』。")
        return
    if not col_hire:
        print("❌ 未识别到『入职日期』列，请检查表头是否为『入职日期』。")
        return
    if not col_emp:
        print("⚠️ 未识别到『姓名』列。")
    if not col_id:
        print("⚠️ 未识别到『身份证号码』列，将仅按姓名分组。")

    # 解析统计月末
    sm = datetime.strptime(cfg["stat_month"], "%Y-%m").date()
    stat_end = month_end(sm.year, sm.month)

    # 解析在职天数截止日
    raw_as_of = cfg.get("as_of_date")
    if raw_as_of in (None, "today", ""):
        as_of = date.today()
    else:
        as_of = datetime.strptime(raw_as_of, "%Y-%m-%d").date()

    print(f"在职天数截止日(as_of)：{as_of}（主在职终点不超过该日）")

    # 解析日期列（保持原列，新增内部列）
    df["_hire"] = pd.to_datetime(df[col_hire], errors="coerce")
    if col_dep and col_dep in df.columns:
        df["_depart"] = pd.to_datetime(df[col_dep], errors="coerce")
    else:
        df["_depart"] = pd.NaT

    include_hire_day = cfg["include_hire_day"]

    # 按人分组
    groups = {}
    for idx, r in df.iterrows():
        key = person_key(r, col_emp, col_id)
        groups.setdefault(key, []).append((idx, r))

    # 汇总统计
    n_employed = 0
    n_left = 0
    n_other = 0
    n_missing_hire = 0
    n_no_months = 0

    rows = []

    for key, members in groups.items():
        employed_hires = []
        leave_intervals = []          # (start, end) 列表
        leave_rows = []               # (idx, r, start, end) 列表
        employed_rows = []
        other_rows = []

        for idx, r in members:
            status = str(r[col_status]).strip() if col_status in r.index else ""
            hire = to_date(r["_hire"])
            depart = to_date(r["_depart"]) if col_dep and col_dep in r.index else None

            if status == "在职":
                n_employed += 1
                employed_rows.append((idx, r))
                if hire is not None:
                    employed_hires.append(hire)
                else:
                    n_missing_hire += 1
            elif status == "离职":
                n_left += 1
                if hire is None:
                    # 离职行缺少入职日期：保留原行，0 天
                    leave_rows.append((idx, r, None, None))
                else:
                    # 截断到统计月末
                    leave_end = min(depart, stat_end) if depart is not None else stat_end
                    start, end = compute_leave_interval(hire, leave_end, include_hire_day)
                    leave_intervals.append((start, end))
                    leave_rows.append((idx, r, start, end))
            else:
                n_other += 1
                other_rows.append((idx, r))

        # 计算该人主在职剩余区间（按年月聚合）
        main_days = compute_main_employment(
            employed_hires, leave_intervals, stat_end, as_of, include_hire_day
        )

        # 输出主在职结果到每个在职行
        if employed_rows:
            if not main_days:
                n_no_months += 1
            for idx, r in employed_rows:
                if not main_days:
                    nr = r.to_dict()
                    nr["在职月份"] = ""
                    nr["在职天数"] = 0
                    rows.append(nr)
                else:
                    for (y, m), d in sorted(main_days.items()):
                        nr = r.to_dict()
                        nr["在职月份"] = pd.Timestamp(y, m, 1)
                        nr["在职天数"] = d
                        rows.append(nr)

        # 输出每个离职行的短离职结果
        for idx, r, start, end in leave_rows:
            nr = r.to_dict()
            if start is None or start > end:
                # 区间为空：入职=离职 或 缺少入职日期 -> 输出 0 天，月份取入职日本月
                if start is None:
                    nr["在职月份"] = ""
                else:
                    # start 是 hire_start_date，用 hire 原始月
                    hire = to_date(r["_hire"])
                    if hire is not None:
                        nr["在职月份"] = pd.Timestamp(hire.year, hire.month, 1)
                    else:
                        nr["在职月份"] = ""
                nr["在职天数"] = 0
                rows.append(nr)
            else:
                month_days = days_by_month(start, end)
                for (y, m), d in sorted(month_days.items()):
                    nr = r.to_dict()
                    nr["在职月份"] = pd.Timestamp(y, m, 1)
                    nr["在职天数"] = d
                    rows.append(nr)

        # 其他状态行原样保留
        for idx, r in other_rows:
            nr = r.to_dict()
            nr["在职月份"] = ""
            nr["在职天数"] = 0
            rows.append(nr)

    # 构建输出 DataFrame
    out = pd.DataFrame(rows)
    orig_cols = [c for c in df.columns if c not in ("_hire", "_depart")]
    front = orig_cols + ["在职月份", "在职天数"]
    # 确保列存在
    front = [c for c in front if c in out.columns]
    out = out[front]

    # 新增“总在职天数”列：按人（姓名 + 身份证）汇总所有展开行的在职天数
    if col_emp and col_id and col_emp in out.columns and col_id in out.columns:
        grp_cols = [col_emp, col_id]
        for c in grp_cols:
            out[c] = out[c].fillna("")
        out["总在职天数"] = out.groupby(grp_cols, dropna=False)["在职天数"].transform("sum")
    else:
        out["总在职天数"] = out["在职天数"]
    front = front + ["总在职天数"]
    out = out[front]

    # 日期列保持 datetime
    for col in [col_hire, col_dep]:
        if col and col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")

    # 写出 Excel 并设置格式
    from openpyxl.styles import Font, Border, Alignment

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="在职月份结果", index=False)
        ws = writer.sheets["在职月份结果"]
        hdr = {cell.value: cell.column_letter for cell in ws[1]}
        for col in [col_hire, col_dep]:
            if col and col in hdr:
                col_letter = hdr[col]
                for row in range(2, ws.max_row + 1):
                    ws[f"{col_letter}{row}"].number_format = "YYYY-MM-DD"
        if "在职月份" in hdr:
            col_letter = hdr["在职月份"]
            for row in range(2, ws.max_row + 1):
                ws[f"{col_letter}{row}"].number_format = "yyyy/m"
        # 清除表头默认加粗/边框
        for cell in ws[1]:
            cell.font = Font(bold=False)
            cell.border = Border()

        # 合并单元格：仅对“总在职天数”列按人连续块合并并垂直居中
        # （数据量大时仅合并单列，避免过多合并操作影响性能）
        merge_cols = []
        if "总在职天数" in hdr:
            merge_cols.append("总在职天数")

        if merge_cols and ws.max_row >= 2:
            # 可用于分组的键列（按姓名 / 身份证；缺失则只用存在的）
            key_cols = []
            if col_emp and col_emp in hdr:
                key_cols.append(hdr[col_emp])
            if col_id and col_id in hdr:
                key_cols.append(hdr[col_id])

            def row_key(r):
                vals = []
                for col_letter in key_cols:
                    v = ws[f"{col_letter}{r}"].value
                    vals.append(str(v) if v is not None else "")
                return tuple(vals)

            # 确定按人的连续块
            blocks = []
            start = 2
            prev_key = None
            for r in range(2, ws.max_row + 1):
                k = row_key(r)
                if prev_key is None:
                    prev_key = k
                    start = r
                elif k != prev_key:
                    blocks.append((start, r - 1))
                    prev_key = k
                    start = r
            blocks.append((start, ws.max_row))

            for col_name in merge_cols:
                col_letter = hdr[col_name]
                for s, e in blocks:
                    if e > s:
                        ws.merge_cells(f"{col_letter}{s}:{col_letter}{e}")
                    for rr in range(s, e + 1):
                        ws[f"{col_letter}{rr}"].alignment = Alignment(vertical="center", horizontal="center")

    # 汇总信息
    valid_rows = pd.to_datetime(out["在职月份"], errors="coerce").notna().sum()
    print(f"✅ 完成。输出文件：{output_file}")
    print(f"   原始记录：{len(df)} 行")
    print(f"   - 在职（参与主在职计算）：{n_employed} 行")
    print(f"   - 离职（按行展开短离职）：{n_left} 行")
    print(f"   - 其他状态（原样保留）：{n_other} 行")
    print(f"   展开后总明细：{len(out)} 行")
    print(f"   有在职月份的有效明细：{valid_rows} 行")
    print(f"   - 在职行中入职日期缺失：{n_missing_hire} 行")
    print(f"   - 在职行中无在职月份：{n_no_months} 行")
    print(f"   工作表『在职月份结果』列：原表列 + 在职月份 + 在职天数 + 总在职天数")
    print(f"   已按人合并：仅“总在职天数”列（垂直居中）")


def main():
    print("本模块为纯计算核心，请通过 GUI 调用 process(input_file, output_file, ...) 或传入参数运行。")
    print("示例：")
    print('  tenure_months_calculator.process("员工信息.xlsx", "结果.xlsx")')


if __name__ == "__main__":
    main()
