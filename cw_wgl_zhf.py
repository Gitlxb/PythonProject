"""
从“员工信息汇总-26年1月”生成：
1) 项目驻场-刷新
2) 项目经理-刷新
3) 稳岗率
4) 驻场&项目经理-奖励汇总
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from tkinter import Tk, filedialog

import numpy as np
import pandas as pd


def choose_excel_file() -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="请选择源 Excel 文件",
        filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")],
    )
    root.destroy()
    if not path:
        return None
    return Path(path)


def normalize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_month_str(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "NaT": np.nan})
    # 已是 YYYY/MM
    ok = s.str.match(r"^\d{4}/\d{2}$", na=False)
    out = s.where(ok, np.nan)

    dt = pd.to_datetime(s, errors="coerce")
    out = out.fillna(dt.dt.strftime("%Y/%m"))
    return out


def to_date(series: pd.Series) -> pd.Series:
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        s = series.astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "NaT": np.nan})
        return pd.to_datetime(s, errors="coerce")


def extract_months(raw: pd.DataFrame) -> list[str]:
    m1 = to_month_str(raw.get("当月入职", pd.Series(dtype=str)))
    m2 = to_month_str(raw.get("当月离职", pd.Series(dtype=str)))
    months = pd.concat([m1, m2], ignore_index=True).dropna().unique().tolist()
    months = sorted(months)
    return months


def read_workbook(input_file: Path, sheet_name: str = None) -> pd.DataFrame:
    """
    读取 Excel 文件
    
    Args:
        input_file: Excel 文件路径
        sheet_name: 工作表名称，如果为 None 则默认读取第一个工作表
    """
    if sheet_name is None:
        # 默认读取第一个工作表
        excel_file = pd.ExcelFile(input_file)
        sheet_name = excel_file.sheet_names[0]
        excel_file.close()
    
    raw = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    raw = normalize_colnames(raw)
    return raw


def prepare_raw(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()

    # 只保留核心列，缺失时补空
    for col in [
        "姓名",
        "所属运营中心",
        "状态",
        "当月入职",
        "当月离职",
        "入职日期",
        "离职日期",
        "项目驻场",
        "项目经理",
        "工厂实际结算",
    ]:
        if col not in df.columns:
            df[col] = np.nan

    df["当月入职_月"] = to_month_str(df["当月入职"])
    df["当月离职_月"] = to_month_str(df["当月离职"])
    df["入职日期_dt"] = to_date(df["入职日期"])
    df["离职日期_dt"] = to_date(df["离职日期"])

    # 优先使用同名列
    df["项目驻场"] = df["项目驻场"].astype(str).str.strip().replace({"": np.nan, "nan": np.nan})
    if "项目经理" in raw.columns:
        df["项目经理"] = raw["项目经理"].astype(str).str.strip().replace({"": np.nan, "nan": np.nan})
    return df


def build_refresh(df: pd.DataFrame, months: list[str], role_col: str, start_month: str, end_month: str) -> pd.DataFrame:
    actors = sorted([x for x in df[role_col].dropna().unique().tolist() if str(x).strip()])

    # 入职透视
    in_tmp = (
        df.dropna(subset=[role_col, "当月入职_月"])
        .groupby(["当月入职_月", role_col], dropna=False)
        .size()
        .unstack(fill_value=0)
        .reindex(index=months, fill_value=0)
    )
    if actors:
        in_tmp = in_tmp.reindex(columns=actors, fill_value=0)
    in_tmp["总计"] = in_tmp.sum(axis=1)

    # 离职透视
    out_tmp = (
        df.dropna(subset=[role_col, "当月离职_月"])
        .groupby(["当月离职_月", role_col], dropna=False)
        .size()
        .unstack(fill_value=0)
        .reindex(index=months, fill_value=0)
    )
    if actors:
        out_tmp = out_tmp.reindex(columns=actors, fill_value=0)
    out_tmp["总计"] = out_tmp.sum(axis=1)

    # 计算稳岗率（针对每个项目驻场）
    # 使用指定的时间范围
    target_months = [m for m in months if m >= start_month and m <= end_month]

    # 计算每个驻场的稳岗率
    stability_rows = []

    # 第一行：时间 + 姓名标题
    # 第一个单元格：结束月份，第二个单元格开始：具体姓名
    name_row = {"指标": end_month, "分类": ""}
    for actor in actors:
        name_row[actor] = actor
    name_row["总计"] = "总计"
    stability_rows.append(name_row)

    # 第二行：月末在职人数
    month_end_row = {"指标": "月末在职人数", "分类": ""}
    for actor in actors:
        # 计算该驻场在目标月份范围内的入职总数
        in_total = in_tmp.loc[target_months, actor].sum() if actor in in_tmp.columns else 0
        # 计算该驻场在目标月份范围内的离职总数
        out_total = out_tmp.loc[target_months, actor].sum() if actor in out_tmp.columns else 0
        month_end_row[actor] = in_total - out_total
    month_end_row["总计"] = sum(month_end_row[actor] for actor in actors)
    stability_rows.append(month_end_row)

    # 第三行：当月总在职人数
    total_on_job_row = {"指标": "当月总在职人数", "分类": ""}
    
    out_months_until_dec = [m for m in target_months if m < end_month]
    for actor in actors:
        # 计算该驻场在目标月份范围内的入职总数
        in_total = in_tmp.loc[target_months, actor].sum() if actor in in_tmp.columns else 0
        # 离职计算：例如开始月份是2025/10，结束月份是2026/02，那么离职计算：当月入职总计减去当月离职的"2025/10-2026/01"总计
        out_total = out_tmp.loc[out_months_until_dec, actor].sum() if actor in out_tmp.columns else 0
        total_on_job_row[actor] = in_total - out_total
    total_on_job_row["总计"] = sum(total_on_job_row[actor] for actor in actors)
    stability_rows.append(total_on_job_row)

    # 第四行：稳岗率（保留小数点）
    stability_rate_row = {"指标": "稳岗率", "分类": ""}
    for actor in actors:
        month_end = month_end_row[actor]
        total_on_job = total_on_job_row[actor]
        if total_on_job > 0:
            rate = month_end / total_on_job
            # 保留两位小数，四舍五入
            stability_rate_row[actor] = round_half_up(rate, 2)
        else:
            stability_rate_row[actor] = 0.0
    # 总计列的稳岗率计算
    total_month_end = month_end_row["总计"]
    total_on_job = total_on_job_row["总计"]
    if total_on_job > 0:
        total_rate = total_month_end / total_on_job
        stability_rate_row["总计"] = round_half_up(total_rate, 2)
    else:
        stability_rate_row["总计"] = 0.0
    stability_rows.append(stability_rate_row)

    # 拼成接近原表的结构
    rows = []
    rows.append({"指标": f"计数项:姓名", "分类": role_col})

    # 当月入职行
    month_in_row = {"指标": "当月入职", "分类": ""}
    for actor in actors:
        month_in_row[actor] = actor
    month_in_row["总计"] = "总计"
    rows.append(month_in_row)

    for m in months:
        rec = {"指标": m, "分类": ""}
        rec.update(in_tmp.loc[m].to_dict())
        rows.append(rec)
    rec_total_in = {"指标": "总计", "分类": ""}
    rec_total_in.update(in_tmp.sum(axis=0).to_dict())
    rows.append(rec_total_in)

    rows.append({"指标": "", "分类": ""})
    rows.append({"指标": f"计数项:姓名", "分类": role_col})

    # 当月离职行
    month_out_row = {"指标": "当月离职", "分类": ""}
    for actor in actors:
        month_out_row[actor] = actor
    month_out_row["总计"] = "总计"
    rows.append(month_out_row)

    for m in months:
        rec = {"指标": m, "分类": ""}
        rec.update(out_tmp.loc[m].to_dict())
        rows.append(rec)
    rec_total_out = {"指标": "总计", "分类": ""}
    rec_total_out.update(out_tmp.sum(axis=0).to_dict())
    rows.append(rec_total_out)

    # 添加稳岗率部分
    rows.append({"指标": "", "分类": ""})
    rows.extend(stability_rows)

    out = pd.DataFrame(rows)
    return out


def build_base_rate_sheet(zc_base_rate: float, pm_base_rate: float) -> pd.DataFrame:
    """
    构建基准数表，包含项目驻场和项目经理的基准数规则
    左侧：项目驻场；右侧：项目经理
    """
    # 项目驻场规则（左侧）
    # 档位从 -8 到 10
    rows = []

    for gear in range(10, -9, -1):  # 10, 9, ..., 0, -1, ..., -8
        # 计算对应的稳岗率
        rate_zc = zc_base_rate + gear * 0.05
        rate_pct_zc = f"{rate_zc*100:.0f}%"

        # 计算单价（项目驻场）
        if gear >= 6:
            unit_zc = 10.0
            note_zc = "≥8元时，实际按照10元结算" if gear == 6 else ""
        elif gear < -5:
            unit_zc = 0.0
            note_zc = "＜2元时，不结算" if gear == -6 else ""
        else:
            unit_zc = gear * 0.5 + 5
            note_zc = ""

        # 如果是基准数所在的行（gear=0），在第一列显示基准数
        col1 = zc_base_rate if gear == 0 else ""

        # 项目经理部分
        rate_pm = pm_base_rate + gear * 0.05
        rate_pct_pm = f"{rate_pm*100:.0f}%"

        # 计算单价（项目经理）
        if gear >= 6:
            unit_pm = 2.0
            note_pm = "≥1.6元时，实际按照2元结算" if gear == 6 else ""
        elif gear < -5:
            unit_pm = 0.0
            note_pm = "＜0.4元时，不结算" if gear == -6 else ""
        else:
            unit_pm = gear * 0.1 + 1
            note_pm = ""

        # 如果是基准数所在的行（gear=0），在第八列显示基准数
        col9 = pm_base_rate if gear == 0 else ""

        row = {
            1: col1,      # 第一列：项目驻场基准数（只在 gear=0 时显示）
            2: gear,      # 第三列：档位
            3: rate_pct_zc,  # 第四列：项目驻场稳岗率百分比
            4: unit_zc,   # 第五列：项目驻场单价
            5: note_zc,   # 第六列：项目驻场备注
            6: "",        # 第七列：空白（用于分隔左右）
            8: col9,      # 第八列：项目经理基准数（只在 gear=0 时显示）
            9: gear,      # 第十列：档位
            10: rate_pct_pm,  # 第十一列：项目经理稳岗率百分比
            11: unit_pm,  # 第十二列：项目经理单价
            12: note_pm,  # 第十三列：项目经理备注
        }
        rows.append(row)

    # 创建 DataFrame
    df = pd.DataFrame(rows)

    # 添加标题行（第一行）
    title_row = {i: "" for i in range(13)}
    title_row[2] = "项目驻场"
    title_row[9] = "项目经理"

    # 将标题行添加到数据中
    df_with_title = pd.concat([
        pd.DataFrame([title_row]),
        df
    ], ignore_index=True)

    return df_with_title


def round_half_up(value, decimals=2):
    if pd.isna(value) or value is None:
        return 0.0
    d = Decimal(str(value))
    rounded = d.quantize(Decimal(10) ** -decimals, rounding=ROUND_HALF_UP)
    return float(rounded)


def build_stability_sheet(df: pd.DataFrame, month: str, zc_refresh_df: pd.DataFrame) -> pd.DataFrame:
    """
    构建稳岗率表
    数据来源：项目驻场表
    """
    # 获取所有项目驻场人员姓名
    actors = sorted([x for x in df["项目驻场"].dropna().unique().tolist() if str(x).strip()])

    # 从刷新表中提取需要的数据
    # 查找"月末在职人数"行
    month_end_row = zc_refresh_df[zc_refresh_df["指标"] == "月末在职人数"]
    # 查找"当月总在职人数"行
    total_on_job_row = zc_refresh_df[zc_refresh_df["指标"] == "当月总在职人数"]
    # 查找"当月入职"部分的数据（使用传入的结束月份）
    month_in_data = zc_refresh_df[zc_refresh_df["指标"] == month].iloc[0] if len(
        zc_refresh_df[zc_refresh_df["指标"] == month]) > 0 else None

    # 构建数据
    records = []
    for actor in actors:
        rec = {"项目驻场": actor}
        # 从月末在职人数行获取数据
        if not month_end_row.empty and actor in month_end_row.iloc[0]:
            rec["月末在职人数"] = int(month_end_row.iloc[0][actor]) if pd.notna(month_end_row.iloc[0][actor]) else 0
        else:
            rec["月末在职人数"] = 0

        # 从当月总在职人数行获取数据
        if not total_on_job_row.empty and actor in total_on_job_row.iloc[0]:
            rec["当月总在职人数"] = int(total_on_job_row.iloc[0][actor]) if pd.notna(
                total_on_job_row.iloc[0][actor]) else 0
        else:
            rec["当月总在职人数"] = 0

        # 计算稳岗率
        if rec["当月总在职人数"] > 0:
            rec["稳岗率"] = round_half_up(rec["月末在职人数"] / rec["当月总在职人数"], 2)
        else:
            rec["稳岗率"] = 0.0

        # 从当月入职数据中获取 2026/01 的数据
        if month_in_data is not None and actor in month_in_data:
            rec["当月入职人数"] = int(month_in_data[actor]) if pd.notna(month_in_data[actor]) else 0
        else:
            rec["当月入职人数"] = 0

        records.append(rec)

    out = pd.DataFrame(records)

    # 固定单价为 5，但如果项目驻场包含“锦途”则设为 0
    def get_unit_price(actor):
        if "锦途" in str(actor):
            return 0.0
        return 5.0

    out["单价"] = out["项目驻场"].apply(get_unit_price)

    # 入职补贴 = 当月入职人数 * 单价
    out["入职补贴（5元/人）"] = out["当月入职人数"] * out["单价"]

    # 当月结算总人数：从原始数据统计每个项目驻场的在职人数
    # 需要与"工厂实际结算"列相对应，只要该列是有效值（非空、非乱码）就计数

    g = df.dropna(subset=["项目驻场"]).copy()
    g["项目驻场"] = g["项目驻场"].astype(str).str.strip()
    g = g[g["项目驻场"] != ""]

    # 动态查找包含"工厂实际结算"字样的列
    settlement_col = None
    for col in g.columns:
        if "工厂实际结算" in str(col):
            settlement_col = col
            break

    # 检查是否找到列，只要是有效值就计数（不需要是数字）
    if settlement_col is not None:
        # 判断是否为有效值：非空且不是乱码（字符串长度合理）
        def is_valid_value(val):
            if pd.isna(val) or val == '':
                return False
            # 转换为字符串，排除明显的乱码（如过长的字符串）
            str_val = str(val).strip()
            if len(str_val) == 0:
                return False
            # 如果字符串过长（超过 50 个字符），可能是乱码
            if len(str_val) > 50:
                return False
            return True

        # 同时满足两个条件：项目驻场有效 + 工厂实际结算有效
        g_valid = g[g[settlement_col].apply(is_valid_value)].copy()
        print(f"原始数据行数：{len(g)}, 过滤后的有效数据行数：{len(g_valid)}")
    else:
        # 如果没有该列，使用所有数据
        print("警告：未找到包含'工厂实际结算'的列，使用所有数据")
        g_valid = g.copy()

    # 统计每个项目驻场的总人数（只要两列都有效就计数，不限制日期，不需要数字转换）
    settlement_counts = g_valid.groupby("项目驻场").size()
    print(f"各驻场人数统计：{settlement_counts.to_dict()}")

    out["当月结算总人数"] = out["项目驻场"].map(settlement_counts).fillna(0).astype(int)

    # 分离"锦途"行和其他行
    jintu_rows = out[out["项目驻场"].str.contains("锦途", na=False)]
    other_rows = out[~out["项目驻场"].str.contains("锦途", na=False)]

    # 对其他行按稳岗率升序排列
    other_rows_sorted = other_rows.sort_values("稳岗率", ascending=True)

    # 计算平均数（只针对其他行，不包含"锦途"行）
    valid_rates_for_avg = other_rows_sorted["稳岗率"].dropna()
    if len(valid_rates_for_avg) > 0:
        avg_rate_calc = valid_rates_for_avg.mean()
    else:
        avg_rate_calc = 0.0

    # 计算中位数（只针对其他行，不包含"锦途"行）
    valid_rates_for_median = other_rows_sorted["稳岗率"].dropna()
    if len(valid_rates_for_median) > 0:
        median_rate_calc = valid_rates_for_median.median()
    else:
        median_rate_calc = 0.0

    # 先添加数据行，然后空两行，再添加总计行和统计行
    # 创建空白行数据（只包含项目驻场列，其他列为 NaN）
    blank_row_data = {"项目驻场": np.nan}

    # 创建总计行（先只设置"项目驻场"列，其他列后面再填充）
    total_row = {"项目驻场": "总计"}

    # 创建平均数行（只设置项目驻场和稳岗率列）
    avg_row_data = {"项目驻场": "平均数", "稳岗率": round_half_up(avg_rate_calc, 2)}

    # 创建中位数行（只设置项目驻场和稳岗率列）
    median_row_data = {"项目驻场": "中位数", "稳岗率": round_half_up(median_rate_calc, 2)}

    # 按顺序拼接：排序后的其他行 + 锦途行 + 总计行 + 空行 + 平均数 + 中位数
    out = pd.concat([other_rows_sorted, jintu_rows], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([blank_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([blank_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([avg_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([median_row_data])], ignore_index=True)

    # 从平均数行和中位数行提取稳岗率值，取较大值作为基准数
    avg_rate = avg_row_data["稳岗率"]
    median_rate = median_row_data["稳岗率"]
    base_rate = max(avg_rate, median_rate)
    print(f"平均数：{avg_rate:.2f}, 中位数：{median_rate:.2f}, 基准数：{base_rate:.2f}")

    # 确保所有需要的列都存在
    # 如果不存在则初始化为 0
    if "单价" not in out.columns:
        out["单价"] = 0.0
    if "入职补贴（5元/人）" not in out.columns:
        out["入职补贴（5元/人）"] = 0.0
    if "当月结算总人数" not in out.columns:
        out["当月结算总人数"] = 0
    if "折算后单价5元/人 (期初)-稳岗" not in out.columns:
        out["折算后单价5元/人 (期初)-稳岗"] = 0.0
    if "稳岗补贴" not in out.columns:
        out["稳岗补贴"] = 0.0

    # 折算后单价5元/人（期初）-稳岗
    # 根据图片和 Excel 公式：IF(INT((E2-$O$5)/0.05)>=6,10,IF(INT((E2-$O$5)/0.05)<-5,0,INT((E2-$O$5)/0.05)*5*0.1+5))
    # E2 是稳岗率，$O$5 是基准数（0.68）
    # 规则说明：
    #   - 43% 以下（不含）：0 元
    #   - 43%（含）~48%（不含）：2.5 元
    #   - 48%（含）~53%（不含）：3 元
    #   - 53%（含）~58%（不含）：3.5 元
    #   - 58%（含）~63%（不含）：4 元
    #   - 63%（含）~68%（不含）：4.5 元
    #   - 68%（含）~73%（不含）：5 元
    #   - 73%（含）~78%（不含）：5.5 元
    #   - 78%（含）~83%（不含）：6 元
    #   - 83%（含）~88%（不含）：6.5 元
    #   - 88%（含）~93%（不含）：7 元
    #   - 93%（含）~98%（不含）：7.5 元
    #   - 98% 及以上：10 元
    def calc_adjusted_unit(rate, actor):
        if pd.isna(rate):
            return 0.0
        # 如果项目驻场包含"锦途"，则折算后单价为 0
        if "锦途" in str(actor):
            return 0.0

        # 计算稳岗率与基准数的差值，每 5% 为一档
        diff = rate - base_rate
        # 计算档位：除以 0.05 后向下取整
        gear = math.floor(diff / 0.05)

        # 根据档位计算单价
        if gear >= 6:
            return 10.0  # 98% 及以上，封顶 10 元
        elif gear < -5:
            return 0.0  # 43% 以下，不结算
        else:
            return gear * 0.5 + 5

    # 只对实际的数据行（非空白行、非总计行、非平均数行、非中位数行）计算折算后单价和正确档位
    data_mask = ~out["项目驻场"].isin(["", "总计", "平均数", "中位数"]) & out["项目驻场"].notna()

    # 计算折算后单价
    out.loc[data_mask, "折算后单价5元/人 (期初)-稳岗"] = out[data_mask].apply(
        lambda row: calc_adjusted_unit(row["稳岗率"], row["项目驻场"]), axis=1
    )

    # 计算稳岗补贴 = 当月结算总人数 * 折算后单价5元/人 (期初)-稳岗
    out.loc[data_mask, "稳岗补贴"] = out.loc[data_mask, "当月结算总人数"] * out.loc[
        data_mask, "折算后单价5元/人 (期初)-稳岗"]

    # 选择需要的列（调整顺序）
    cols = [
        "项目驻场", "月末在职人数", "当月总在职人数", "稳岗率",
        "当月入职人数", "单价", "入职补贴（5元/人）",
        "当月结算总人数", "折算后单价5元/人 (期初)-稳岗", "稳岗补贴"
    ]
    out = out[cols]

    # 现在填充总计行数据（所有列都已计算完成）
    for col in ["月末在职人数", "当月总在职人数", "当月入职人数", "单价", "入职补贴（5元/人）", "当月结算总人数",
                "稳岗补贴"]:
        total_row[col] = out.loc[data_mask, col].sum()
    # 稳岗率列的总计 = 月末在职人数总计 / 当月总在职人数总计
    total_month_end = out.loc[data_mask, "月末在职人数"].sum()
    total_on_job = out.loc[data_mask, "当月总在职人数"].sum()
    if total_on_job > 0:
        total_row["稳岗率"] = round_half_up(total_month_end / total_on_job, 2)
    else:
        total_row["稳岗率"] = 0.0
    # 折算后单价的总计 = 该列数据的总和
    total_row["折算后单价5元/人 (期初)-稳岗"] = out.loc[data_mask, "折算后单价5元/人 (期初)-稳岗"].sum()

    # 更新总计行数据
    total_row_idx = out[out["项目驻场"] == "总计"].index
    if len(total_row_idx) > 0:
        for col, val in total_row.items():
            out.loc[total_row_idx[0], col] = val

    # 确保空白行、平均数行、中位数行的其他列为空（不显示数字）
    # 找到空白行、平均数行、中位数行的索引
    for idx in out.index:
        project_field = out.loc[idx, "项目驻场"]
        if pd.isna(project_field) or project_field == "平均数" or project_field == "中位数":
            # 只保留稳岗率列的值（如果是平均数或中位数行），其他列设为 NaN
            if project_field == "平均数" or project_field == "中位数":
                for col in cols:
                    if col != "项目驻场" and col != "稳岗率":
                        out.loc[idx, col] = np.nan
            else:
                # 空白行所有列都设为 NaN
                for col in cols:
                    out.loc[idx, col] = np.nan

    return out


def build_stability_sheet_combined(stability_zc: pd.DataFrame, stability_pm: pd.DataFrame) -> pd.DataFrame:
    """
    构建稳岗率表，包含项目驻场和项目经理两部分（在同一张表中，左右排列）
    项目驻场在左侧，项目经理在右侧，中间空 3 列
    """
    # 重置索引以确保对齐
    stability_zc = stability_zc.reset_index(drop=True)
    stability_pm = stability_pm.reset_index(drop=True)

    # 创建 3 个空白列（使用 NaN 作为列名，导出时不会显示）
    num_rows = len(stability_zc)
    blank_col1 = pd.DataFrame({np.nan: [np.nan] * num_rows})
    blank_col2 = pd.DataFrame({np.nan: [np.nan] * num_rows})
    blank_col3 = pd.DataFrame({np.nan: [np.nan] * num_rows})

    # 左右拼接：项目驻场表 + 3 个空白列 + 项目经理表
    combined = pd.concat([
        stability_zc,
        blank_col1,
        blank_col2,
        blank_col3,
        stability_pm
    ], axis=1, ignore_index=False)

    return combined


def build_stability_sheet_pm(df: pd.DataFrame, month: str, pm_refresh_df: pd.DataFrame) -> pd.DataFrame:
    """
    构建项目经理稳岗率表
    数据来源：项目经理刷新表
    单价固定为 1 元/人（锦途为 0 元）
    """
    # 从刷新表中提取数据
    # 获取所有项目经理人员姓名
    actors = sorted([x for x in df["项目经理"].dropna().unique().tolist() if str(x).strip()])

    # 从刷新表中提取需要的数据
    # 查找"月末在职人数"行
    month_end_row = pm_refresh_df[pm_refresh_df["指标"] == "月末在职人数"]
    # 查找"当月总在职人数"行
    total_on_job_row = pm_refresh_df[pm_refresh_df["指标"] == "当月总在职人数"]
    # 查找"当月入职"部分的数据（使用传入的结束月份）
    month_in_data = pm_refresh_df[pm_refresh_df["指标"] == month].iloc[0] if len(
        pm_refresh_df[pm_refresh_df["指标"] == month]) > 0 else None

    # 构建数据
    records = []
    for actor in actors:
        rec = {"项目经理": actor}
        # 从月末在职人数行获取数据
        if not month_end_row.empty and actor in month_end_row.iloc[0]:
            rec["月末在职人数"] = int(month_end_row.iloc[0][actor]) if pd.notna(month_end_row.iloc[0][actor]) else 0
        else:
            rec["月末在职人数"] = 0

        # 从当月总在职人数行获取数据
        if not total_on_job_row.empty and actor in total_on_job_row.iloc[0]:
            rec["当月总在职人数"] = int(total_on_job_row.iloc[0][actor]) if pd.notna(
                total_on_job_row.iloc[0][actor]) else 0
        else:
            rec["当月总在职人数"] = 0

        # 计算稳岗率
        if rec["当月总在职人数"] > 0:
            rec["稳岗率"] = round_half_up(rec["月末在职人数"] / rec["当月总在职人数"], 2)
        else:
            rec["稳岗率"] = 0.0

        # 从当月入职数据中获取 2026/01 的数据
        if month_in_data is not None and actor in month_in_data:
            rec["当月入职人数"] = int(month_in_data[actor]) if pd.notna(month_in_data[actor]) else 0
        else:
            rec["当月入职人数"] = 0

        records.append(rec)

    out = pd.DataFrame(records)

    # 固定单价为 1，但如果项目经理包含"锦途"则设为 0
    def get_unit_price(actor):
        if "锦途" in str(actor):
            return 0.0
        return 1.0

    out["单价"] = out["项目经理"].apply(get_unit_price)

    # 入职补贴 = 当月入职人数 * 单价
    out["入职补贴（1元/人）"] = out["当月入职人数"] * out["单价"]

    # 当月结算总人数：从原始数据统计每个项目经理的在职人数
    g = df.dropna(subset=["项目经理"]).copy()
    g["项目经理"] = g["项目经理"].astype(str).str.strip()
    g = g[g["项目经理"] != ""]

    # 动态查找包含"工厂实际结算"字样的列
    settlement_col = None
    for col in g.columns:
        if "工厂实际结算" in str(col):
            settlement_col = col
            break

    # 检查是否找到列，只要是有效值就计数
    if settlement_col is not None:
        def is_valid_value(val):
            if pd.isna(val) or val == '':
                return False
            str_val = str(val).strip()
            if len(str_val) == 0:
                return False
            if len(str_val) > 50:
                return False
            return True

        g_valid = g[g[settlement_col].apply(is_valid_value)].copy()
        print(f"项目经理原始数据行数：{len(g)}, 过滤后的有效数据行数：{len(g_valid)}")
    else:
        print("警告：未找到包含'工厂实际结算'的列，使用所有数据")
        g_valid = g.copy()

    # 统计每个项目经理的总人数
    settlement_counts = g_valid.groupby("项目经理").size()
    print(f"各项目经理人数统计：{settlement_counts.to_dict()}")

    out["当月结算总人数"] = out["项目经理"].map(settlement_counts).fillna(0).astype(int)

    # 分离"锦途"行和其他行
    jintu_rows = out[out["项目经理"].str.contains("锦途", na=False)]
    other_rows = out[~out["项目经理"].str.contains("锦途", na=False)]

    # 对其他行按稳岗率升序排列
    other_rows_sorted = other_rows.sort_values("稳岗率", ascending=True)

    # 计算平均数（只针对其他行，不包含"锦途"行）
    valid_rates_for_avg = other_rows_sorted["稳岗率"].dropna()
    if len(valid_rates_for_avg) > 0:
        avg_rate_calc = valid_rates_for_avg.mean()
    else:
        avg_rate_calc = 0.0

    # 计算中位数（只针对其他行，不包含"锦途"行）
    valid_rates_for_median = other_rows_sorted["稳岗率"].dropna()
    if len(valid_rates_for_median) > 0:
        median_rate_calc = valid_rates_for_median.median()
    else:
        median_rate_calc = 0.0

    # 创建空白行数据
    blank_row_data = {col: np.nan for col in out.columns}

    # 创建总计行
    total_row = {col: np.nan for col in out.columns}
    total_row["项目经理"] = "总计"

    # 创建平均数行
    avg_row_data = {col: np.nan for col in out.columns}
    avg_row_data["项目经理"] = "平均数"
    avg_row_data["稳岗率"] = round_half_up(avg_rate_calc, 2)

    # 创建中位数行
    median_row_data = {col: np.nan for col in out.columns}
    median_row_data["项目经理"] = "中位数"
    median_row_data["稳岗率"] = round_half_up(median_rate_calc, 2)

    # 按顺序拼接：排序后的其他行 + 锦途行 + 总计行 + 空行 + 平均数 + 中位数
    out = pd.concat([other_rows_sorted, jintu_rows], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([blank_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([blank_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([avg_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([median_row_data])], ignore_index=True)

    # 从平均数行和中位数行提取稳岗率值，取较大值作为基准数
    avg_rate = avg_row_data["稳岗率"]
    median_rate = median_row_data["稳岗率"]
    base_rate = max(avg_rate, median_rate)
    print(f"项目经理 - 平均数：{avg_rate:.2f}, 中位数：{median_rate:.2f}, 基准数：{base_rate:.2f}")

    # 折算后单价 5 元/人（期初）- 稳岗
    # 根据图片 Excel 公式：IF(INT((E2-$O$5)/0.05)>=6,2,IF(INT((E2-$O$5)/0.05)<-5,0,INT((E2-$O$5)/0.05)*5*0.1+0.5))
    # E2 是稳岗率，$O$5 是基准数（0.73）
    # 规则说明：
    #   - 48% 以下（不含）：0 元
    #   - 48%（含）~53%（不含）：0.5 元
    #   - 53%（含）~58%（不含）：0.6 元
    #   - ...每 5% 递增 0.1 元
    #   - 68%（含）~73%（不含）：1 元
    #   - ...
    #   - 98%（含）~103%（不含）：1.5 元
    #   - 103% 及以上：2 元（封顶）
    def calc_adjusted_unit_pm(rate, actor):
        if pd.isna(rate):
            return 0.0
        # 如果项目经理包含"锦途"，则折算后单价为 0
        if "锦途" in str(actor):
            return 0.0

        # 计算稳岗率与基准数的差值，每 5% 为一档
        diff = rate - base_rate
        # 计算档位：除以 0.05 后向下取整
        gear = math.floor(diff / 0.05)

        # 根据档位计算单价
        if gear >= 6:
            return 2.0  # 103% 及以上，封顶 2 元
        elif gear < -5:
            return 0.0  # 48% 以下，不结算
        else:
            return gear * 0.1 + 1

    # 只对实际的数据行（非空白行、非总计行、非平均数行、非中位数行）计算折算后单价和正确档位
    data_mask = ~out["项目经理"].isin(["", "总计", "平均数", "中位数"]) & out["项目经理"].notna()

    # 计算折算后单价
    out.loc[data_mask, "折算后单价5元/人 (期初)-稳岗"] = out[data_mask].apply(
        lambda row: calc_adjusted_unit_pm(row["稳岗率"], row["项目经理"]), axis=1
    )

    # 计算稳岗补贴 = 当月结算总人数 * 折算后单价
    out.loc[data_mask, "稳岗补贴"] = out.loc[data_mask, "当月结算总人数"] * out.loc[
        data_mask, "折算后单价5元/人 (期初)-稳岗"]

    # 填充总计行数据
    for col in ["月末在职人数", "当月总在职人数", "当月入职人数", "单价", "入职补贴（1元/人）", "当月结算总人数",
                "稳岗补贴"]:
        total_row[col] = out.loc[data_mask, col].sum()
    # 稳岗率列的总计
    total_month_end = out.loc[data_mask, "月末在职人数"].sum()
    total_on_job = out.loc[data_mask, "当月总在职人数"].sum()
    if total_on_job > 0:
        total_row["稳岗率"] = round_half_up(total_month_end / total_on_job, 2)
    else:
        total_row["稳岗率"] = 0.0
    # 折算后单价的总计
    total_row["折算后单价5元/人 (期初)-稳岗"] = out.loc[data_mask, "折算后单价5元/人 (期初)-稳岗"].sum()

    # 更新总计行数据
    total_row_idx = out[out["项目经理"] == "总计"].index
    if len(total_row_idx) > 0:
        for col, val in total_row.items():
            out.loc[total_row_idx[0], col] = val

    # 选择需要的列（调整顺序）
    cols = [
        "项目经理", "月末在职人数", "当月总在职人数", "稳岗率",
        "当月入职人数", "单价", "入职补贴（1元/人）",
        "当月结算总人数", "折算后单价5元/人 (期初)-稳岗", "稳岗补贴"
    ]
    out = out[cols]

    return out


def build_reward_sheet(stability_zc: pd.DataFrame, stability_pm: pd.DataFrame, raw_data: pd.DataFrame) -> pd.DataFrame:
    """
    构建奖励汇总表
    表结构：姓名 | 入职补贴（5元/人）| 稳岗补贴_驻场 | 入职补贴（1元/人）| 稳岗补贴_经理 | 合计 | 备注
    数据来源：
    - 入职补贴（5元/人）和稳岗补贴_驻场：来自稳岗率表项目驻场部分
    - 入职补贴（1元/人）和稳岗补贴_经理：来自稳岗率表项目经理部分
    - 合计：四个补贴列的总和
    - 备注：
    """
    # 从项目驻场部分提取数据
    valid_rows_zc = stability_zc[
        ~stability_zc["项目驻场"].isin(["总计", "平均数", "中位数"]) &
        ~stability_zc["项目驻场"].str.contains("对比结果", na=False) &
        stability_zc["项目驻场"].notna()
        ].copy()

    zc_data = valid_rows_zc[
        ["项目驻场", "入职补贴（5元/人）", "稳岗补贴"]].copy()
    zc_data.columns = ["姓名", "入职补贴_驻场", "稳岗补贴_驻场"]
    zc_data = zc_data.fillna(0)

    # 从项目经理部分提取数据
    valid_rows_pm = stability_pm[
        ~stability_pm["项目经理"].isin(["总计", "平均数", "中位数"]) &
        ~stability_pm["项目经理"].str.contains("这是对比的结果", na=False) &
        stability_pm["项目经理"].notna()
        ].copy()

    pm_data = valid_rows_pm[
        ["项目经理", "入职补贴（1元/人）", "稳岗补贴"]].copy()
    pm_data.columns = ["姓名", "入职补贴_经理", "稳岗补贴_经理"]
    pm_data = pm_data.fillna(0)

    # 确保姓名列都是字符串类型
    zc_data["姓名"] = zc_data["姓名"].astype(str)
    pm_data["姓名"] = pm_data["姓名"].astype(str)

    # 合并驻场和项目经理数据（全外连接，同一个姓名可能在两表中都存在）
    reward_merged = pd.merge(zc_data, pm_data, on="姓名", how="outer").fillna(0)

    # 计算合计列（四个补贴列的总和）
    reward_merged["合计"] = (
        reward_merged["入职补贴_驻场"] +
        reward_merged["稳岗补贴_驻场"] +
        reward_merged["入职补贴_经理"] +
        reward_merged["稳岗补贴_经理"]
    )

    # 添加备注列，先初始化为空
    reward_merged["备注"] = ""

    # 按姓名排序
    reward_merged = reward_merged.sort_values("姓名").reset_index(drop=True)

    # 调整列顺序并重命名
    reward_merged = reward_merged[[
        "姓名",
        "入职补贴_驻场",
        "稳岗补贴_驻场",
        "入职补贴_经理",
        "稳岗补贴_经理",
        "合计",
        "备注"
    ]].rename(columns={
        "入职补贴_驻场": "入职补贴（5元/人）",
        "稳岗补贴_驻场": "稳岗补贴_驻场",
        "入职补贴_经理": "入职补贴（1元/人）",
        "稳岗补贴_经理": "稳岗补贴_经理"
    })

    # 添加总计行
    total_row = {
        "姓名": "总计",
        "入职补贴（5元/人）": reward_merged["入职补贴（5元/人）"].sum(),
        "稳岗补贴_驻场": reward_merged["稳岗补贴_驻场"].sum(),
        "入职补贴（1元/人）": reward_merged["入职补贴（1元/人）"].sum(),
        "稳岗补贴_经理": reward_merged["稳岗补贴_经理"].sum(),
        "合计": reward_merged["合计"].sum(),
        "备注": ""
    }

    # 使用 concat 添加总计行
    total_df = pd.DataFrame([total_row])
    reward_merged = pd.concat([reward_merged, total_df], ignore_index=True)
    return reward_merged


def get_target_month(months: list[str]) -> str:
    # 默认取最新月份，例如 2026/01
    if not months:
        return ""
    return months[-1]


def save_output(input_file: Path, outputs: dict[str, pd.DataFrame], selected_sheet: str = None, copy_original_sheets: bool = True) -> Path:
    """
    保存输出文件
    Args:
        input_file: 源 Excel 文件路径
        outputs: 要保存的工作表字典 {sheet_name: DataFrame}
        selected_sheet: 被选中的工作表名称，如果为 None 则不复制原工作表
        copy_original_sheets: 是否复制原始工作表（默认 True，设为 False 可显著提升速度）
    """
    from tkinter import Tk, filedialog
    import shutil
    from openpyxl import load_workbook

    # 弹窗让用户选择保存位置和文件名
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    out_path = filedialog.asksaveasfilename(
        title="请选择保存位置和文件名",
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx")],
        initialfile=f"{input_file.stem}_处理结果"
    )
    root.destroy()
    if not out_path:
        raise SystemExit("未选择保存位置，程序退出。")
    out_path = Path(out_path)

    # 优化方案：先复制文件，再添加新的工作表
    if copy_original_sheets:
        try:
            # 直接复制整个文件（二进制复制，速度极快）
            shutil.copy2(input_file, out_path)
            print(f"已复制原始文件")

            # 加载复制后的文件
            wb = load_workbook(out_path)

            # 添加新生成的工作表（不删除任何原始工作表）
            for sheet_name, df in outputs.items():
                # 如果同名工作表已存在，删除它
                if sheet_name in wb.sheetnames:
                    del wb[sheet_name]

                ws = wb.create_sheet(title=sheet_name)

                # 将 DataFrame 写入工作表
                if sheet_name == "基准数":
                    # 基准数表不需要列名
                    for r_idx, row in df.iterrows():
                        for c_idx, value in enumerate(row.values, start=1):
                            if pd.notna(value):
                                ws.cell(row=r_idx+1, column=c_idx, value=value)
                else:
                    # 其他表格需要列名
                    # 写入列名
                    for c_idx, col_name in enumerate(df.columns, start=1):
                        ws.cell(row=1, column=c_idx, value=col_name)
                    # 写入数据
                    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
                        for c_idx, value in enumerate(row.values, start=1):
                            if pd.notna(value):
                                ws.cell(row=r_idx, column=c_idx, value=value)

            wb.save(out_path)
            wb.close()
            print(f"已添加新工作表：{list(outputs.keys())}")

        except Exception as e:
            print(f"错误：复制工作表失败：{e}")
            print("将使用备用方案（不复制原始工作表）")
            # 如果复制失败，使用备用方案
            copy_original_sheets = False

    # 备用方案：不复制原始工作表，只保存新生成的工作表
    if not copy_original_sheets:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            # 写入新生成的工作表
            for sheet, df in outputs.items():
                # 基准数表不写入列名和索引
                if sheet == "基准数":
                    df.to_excel(writer, sheet_name=sheet, index=False, header=False)
                else:
                    df.to_excel(writer, sheet_name=sheet, index=False)

    return out_path


def main() -> None:
    input_file = choose_excel_file()
    raw = read_workbook(input_file)
    data = prepare_raw(raw)

    months = extract_months(data)

    # 默认使用最新月份
    target_month = get_target_month(months)
    start_month = None
    end_month = None

    zc_refresh = build_refresh(data, months, "项目驻场", start_month, end_month)
    pm_refresh = build_refresh(data, months, "项目经理", start_month, end_month)
    stability_zc = build_stability_sheet(data, target_month, zc_refresh)
    stability_pm = build_stability_sheet_pm(data, target_month, pm_refresh)
    stability_combined = build_stability_sheet_combined(stability_zc.copy(), stability_pm.copy())

    reward = build_reward_sheet(stability_zc, stability_pm, data)

    # 从稳岗率表中提取基准数
    # 项目驻场基准数：查找"平均数"或"中位数"行的稳岗率值
    zc_avg_row = stability_zc[stability_zc["项目驻场"] == "平均数"]
    zc_median_row = stability_zc[stability_zc["项目驻场"] == "中位数"]

    zc_avg_rate = zc_avg_row["稳岗率"].iloc[0] if not zc_avg_row.empty else 0.0
    zc_median_rate = zc_median_row["稳岗率"].iloc[0] if not zc_median_row.empty else 0.0
    zc_base_rate = max(zc_avg_rate, zc_median_rate)

    # 项目经理基准数：查找"平均数"或"中位数"行的稳岗率值
    pm_avg_row = stability_pm[stability_pm["项目经理"] == "平均数"]
    pm_median_row = stability_pm[stability_pm["项目经理"] == "中位数"]

    pm_avg_rate = pm_avg_row["稳岗率"].iloc[0] if not pm_avg_row.empty else 0.0
    pm_median_rate = pm_median_row["稳岗率"].iloc[0] if not pm_median_row.empty else 0.0
    pm_base_rate = max(pm_avg_rate, pm_median_rate)

    # 生成基准数表
    base_rate_sheet = build_base_rate_sheet(zc_base_rate, pm_base_rate)

    out_file = save_output(
        input_file,
        {
            "项目驻场": zc_refresh,
            "项目经理": pm_refresh,
            "稳岗率": stability_combined,
            "驻场&项目经理 - 奖励汇总": reward,
            "基准数": base_rate_sheet,
        },
        selected_sheet=None,  # 简单脚本模式不复制原始工作表
    )

    print(f"处理完成：{out_file}")
    print(f"核算月份：{target_month}")
    print(f"项目驻场基准数：{zc_base_rate:.2f}")
    print(f"项目经理基准数：{pm_base_rate:.2f}")
