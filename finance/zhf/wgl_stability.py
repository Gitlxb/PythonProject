"""
稳岗率计算 - 稳岗率计算模块
包含：build_stability_sheet, build_stability_sheet_pm, build_stability_sheet_combined, build_base_rate_sheet
"""

import math
import numpy as np
import pandas as pd

from .wgl_config import *
from .wgl_utils import (round_half_up, is_valid_settlement_value, 
                        find_settlement_column, calc_gear, 
                        calc_unit_price_zc, calc_unit_price_pm)


def build_stability_sheet(df: pd.DataFrame, month: str, zc_refresh_df: pd.DataFrame) -> pd.DataFrame:
    """
    构建稳岗率表（项目驻场）
    
    Args:
        df: 原始数据框
        month: 目标月份
        zc_refresh_df: 项目驻场刷新表
    
    Returns:
        稳岗率表数据框
    """
    return _build_stability_sheet_common(
        df=df,
        month=month,
        refresh_df=zc_refresh_df,
        role_col="项目驻场",
        unit_price_default=DEFAULT_UNIT_ZC,
        unit_price_jintu=DEFAULT_UNIT_JINTU,
        in_subsidy_col=COL_IN_SUBSIDY_ZC,
        is_zc=True
    )


def build_stability_sheet_pm(df: pd.DataFrame, month: str, pm_refresh_df: pd.DataFrame) -> pd.DataFrame:
    """
    构建项目经理稳岗率表
    
    Args:
        df: 原始数据框
        month: 目标月份
        pm_refresh_df: 项目经理刷新表
    
    Returns:
        项目经理稳岗率表数据框
    """
    return _build_stability_sheet_common(
        df=df,
        month=month,
        refresh_df=pm_refresh_df,
        role_col="项目经理",
        unit_price_default=DEFAULT_UNIT_PM,
        unit_price_jintu=DEFAULT_UNIT_JINTU,
        in_subsidy_col=COL_IN_SUBSIDY_PM,
        is_zc=False
    )


def _build_stability_sheet_common(df: pd.DataFrame, month: str, refresh_df: pd.DataFrame,
                                   role_col: str, unit_price_default: float, 
                                   unit_price_jintu: float, in_subsidy_col: str,
                                   is_zc: bool) -> pd.DataFrame:
    """
    构建稳岗率表（通用函数）
    
    Args:
        df: 原始数据框
        month: 目标月份
        refresh_df: 刷新表
        role_col: 角色列名（"项目驻场"或"项目经理"）
        unit_price_default: 默认单价
        unit_price_jintu: 锦途单价
        in_subsidy_col: 入职补贴列名
        is_zc: 是否为项目驻场（True=项目驻场，False=项目经理）
    
    Returns:
        稳岗率表数据框
    """
    # 获取所有角色人员姓名
    actors = sorted([x for x in df[role_col].dropna().unique().tolist() if str(x).strip()])
    
    # 从刷新表中提取需要的数据
    month_end_row = refresh_df[refresh_df["指标"] == "月末在职人数"]
    total_on_job_row = refresh_df[refresh_df["指标"] == "当月总在职人数"]
    
    # 构建数据
    records = _build_stability_records(actors, month_end_row, total_on_job_row, role_col)
    
    out = pd.DataFrame(records)
    
    # 重新计算当月入职人数
    out = _calc_in_count(df, out, month, role_col)
    
    # 设置单价和入职补贴
    out = _set_unit_price_and_subsidy(out, role_col, unit_price_default, 
                                       unit_price_jintu, in_subsidy_col)
    
    # 计算当月结算总人数
    out = _calc_settlement_count(df, out, role_col)
    
    # 分离"锦途"行和其他行，排序
    out, jintu_rows, other_rows_sorted = _sort_and_split_rows(out, role_col)
    
    # 计算平均数和中位数
    avg_rate_calc, median_rate_calc = _calc_avg_median(other_rows_sorted)
    
    # 构建最终表格（包含总计、平均数、中位数）
    out = _build_final_table(out, jintu_rows, other_rows_sorted, 
                              avg_rate_calc, median_rate_calc, role_col)
    
    # 计算基准数和折算后单价
    base_rate = max(avg_rate_calc, median_rate_calc)
    out = _calc_adjusted_unit_and_subsidy(out, role_col, base_rate, is_zc)
    
    # 选择需要的列
    out = _select_output_columns(out, role_col, in_subsidy_col)
    
    # 填充总计行
    out = _fill_total_row(out, role_col)
    
    # 清理特殊行的数据
    out = _clean_special_rows(out, role_col)
    
    return out


def _build_stability_records(actors: list, month_end_row: pd.DataFrame, 
                              total_on_job_row: pd.DataFrame, role_col: str) -> list[dict]:
    """
    构建稳岗率记录
    
    Args:
        actors: 角色列表
        month_end_row: 月末在职人数行
        total_on_job_row: 当月总在职人数行
        role_col: 角色列名
    
    Returns:
        记录列表
    """
    records = []
    for actor in actors:
        rec = {role_col: actor}
        
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
        
        # 当月入职人数初始化为0，后续从原始数据重新计算
        rec["当月入职人数"] = 0
        
        records.append(rec)
    
    return records


def _calc_in_count(df: pd.DataFrame, out: pd.DataFrame, month: str, role_col: str) -> pd.DataFrame:
    """
    计算当月入职人数
    
    Args:
        df: 原始数据框
        out: 输出数据框
        month: 目标月份
        role_col: 角色列名
    
    Returns:
        更新后的输出数据框
    """
    settlement_col = find_settlement_column(df)
    
    # 筛选当月入职_月为目标月份且工厂实际结算有效的行
    in_mask = df["当月入职_月"] == month
    if settlement_col is not None:
        in_mask = in_mask & df[settlement_col].apply(is_valid_settlement_value)
    
    month_in_counts = df[in_mask].dropna(subset=[role_col]).groupby(role_col).size()
    out["当月入职人数"] = out[role_col].map(month_in_counts).fillna(0).astype(int)
    
    return out


def _set_unit_price_and_subsidy(out: pd.DataFrame, role_col: str, 
                                 unit_price_default: float, 
                                 unit_price_jintu: float, 
                                 in_subsidy_col: str) -> pd.DataFrame:
    """
    设置单价和入职补贴
    
    Args:
        out: 输出数据框
        role_col: 角色列名
        unit_price_default: 默认单价
        unit_price_jintu: 锦途单价
        in_subsidy_col: 入职补贴列名
    
    Returns:
        更新后的输出数据框
    """
    def get_unit_price(actor):
        if JINTU_KEYWORD in str(actor):
            return unit_price_jintu
        return unit_price_default
    
    out["单价"] = out[role_col].apply(get_unit_price)
    out[in_subsidy_col] = out["当月入职人数"] * out["单价"]
    
    return out


def _calc_settlement_count(df: pd.DataFrame, out: pd.DataFrame, role_col: str) -> pd.DataFrame:
    """
    计算当月结算总人数
    
    Args:
        df: 原始数据框
        out: 输出数据框
        role_col: 角色列名
    
    Returns:
        更新后的输出数据框
    """
    g = df.dropna(subset=[role_col]).copy()
    g[role_col] = g[role_col].astype(str).str.strip()
    g = g[g[role_col] != ""]
    
    settlement_col = find_settlement_column(g)
    
    # 检查是否找到列，只要是有效值就计数（不需要是数字）
    if settlement_col is not None:
        g_valid = g[g[settlement_col].apply(is_valid_settlement_value)].copy()
        print(f"原始数据行数：{len(g)}, 过滤后的有效数据行数：{len(g_valid)}")
    else:
        # 如果没有该列，使用所有数据
        print("警告：未找到包含'工厂实际结算'的列，使用所有数据")
        g_valid = g.copy()
    
    # 统计每个角色的总人数
    settlement_counts = g_valid.groupby(role_col).size()
    print(f"各{role_col}人数统计：{settlement_counts.to_dict()}")
    
    out["当月结算总人数"] = out[role_col].map(settlement_counts).fillna(0).astype(int)
    
    return out


def _sort_and_split_rows(out: pd.DataFrame, role_col: str):
    """
    排序和分离行
    
    Args:
        out: 输出数据框
        role_col: 角色列名
    
    Returns:
        (out, jintu_rows, other_rows_sorted)
    """
    # 分离"锦途"行和其他行
    jintu_rows = out[out[role_col].str.contains(JINTU_KEYWORD, na=False)]
    other_rows = out[~out[role_col].str.contains(JINTU_KEYWORD, na=False)]
    
    # 对其他行按稳岗率升序排列
    other_rows_sorted = other_rows.sort_values("稳岗率", ascending=True)
    
    return out, jintu_rows, other_rows_sorted


def _calc_avg_median(other_rows_sorted: pd.DataFrame) -> tuple[float, float]:
    """
    计算平均数和中位数
    
    Args:
        other_rows_sorted: 排序后的其他行
    
    Returns:
        (avg_rate, median_rate)
    """
    valid_rates_for_avg = other_rows_sorted["稳岗率"].dropna()
    if len(valid_rates_for_avg) > 0:
        avg_rate_calc = valid_rates_for_avg.mean()
    else:
        avg_rate_calc = 0.0
    
    valid_rates_for_median = other_rows_sorted["稳岗率"].dropna()
    if len(valid_rates_for_median) > 0:
        median_rate_calc = valid_rates_for_median.median()
    else:
        median_rate_calc = 0.0
    
    return avg_rate_calc, median_rate_calc


def _build_final_table(out: pd.DataFrame, jintu_rows: pd.DataFrame, 
                        other_rows_sorted: pd.DataFrame, 
                        avg_rate_calc: float, median_rate_calc: float,
                        role_col: str) -> pd.DataFrame:
    """
    构建最终表格（包含总计、平均数、中位数）
    
    Args:
        out: 输出数据框
        jintu_rows: 锦途行
        other_rows_sorted: 排序后的其他行
        avg_rate_calc: 平均数
        median_rate_calc: 中位数
        role_col: 角色列名
    
    Returns:
        更新后的输出数据框
    """
    # 创建空白行数据
    blank_row_data = {role_col: np.nan}
    
    # 创建总计行
    total_row = {role_col: "总计"}
    
    # 创建平均数行
    avg_row_data = {role_col: "平均数", "稳岗率": round_half_up(avg_rate_calc, 2)}
    
    # 创建中位数行
    median_row_data = {role_col: "中位数", "稳岗率": round_half_up(median_rate_calc, 2)}
    
    # 按顺序拼接
    out = pd.concat([other_rows_sorted, jintu_rows], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([blank_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([blank_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([avg_row_data])], ignore_index=True)
    out = pd.concat([out, pd.DataFrame([median_row_data])], ignore_index=True)
    
    return out


def _calc_adjusted_unit_and_subsidy(out: pd.DataFrame, role_col: str, 
                                     base_rate: float, is_zc: bool) -> pd.DataFrame:
    """
    计算折算后单价和稳岗补贴
    
    Args:
        out: 输出数据框
        role_col: 角色列名
        base_rate: 基准数
        is_zc: 是否为项目驻场
    
    Returns:
        更新后的输出数据框
    """
    # 确保所有需要的列都存在
    if "单价" not in out.columns:
        out["单价"] = 0.0
    if COL_STABILITY_SUBSIDY not in out.columns:
        out[COL_STABILITY_SUBSIDY] = 0.0
    if COL_ADJUSTED_UNIT not in out.columns:
        out[COL_ADJUSTED_UNIT] = 0.0
    
    # 只对实际的数据行计算折算后单价和正确档位
    data_mask = ~out[role_col].isin(EXCLUDE_NAMES) & out[role_col].notna()
    
    # 计算折算后单价
    if is_zc:
        out.loc[data_mask, COL_ADJUSTED_UNIT] = out[data_mask].apply(
            lambda row: _calc_adjusted_unit_zc(row["稳岗率"], row[role_col], base_rate), axis=1
        )
    else:
        out.loc[data_mask, COL_ADJUSTED_UNIT] = out[data_mask].apply(
            lambda row: _calc_adjusted_unit_pm(row["稳岗率"], row[role_col], base_rate), axis=1
        )
    
    # 计算稳岗补贴 = 当月结算总人数 * 折算后单价
    out.loc[data_mask, COL_STABILITY_SUBSIDY] = out.loc[data_mask, "当月结算总人数"] * out.loc[
        data_mask, COL_ADJUSTED_UNIT]
    
    return out


def _calc_adjusted_unit_zc(rate, actor, base_rate: float) -> float:
    """
    计算项目驻场的折算后单价
    
    Args:
        rate: 稳岗率
        actor: 角色名
        base_rate: 基准数
    
    Returns:
        折算后单价
    """
    if pd.isna(rate):
        return 0.0
    # 如果项目驻场包含"锦途"，则折算后单价为 0
    if JINTU_KEYWORD in str(actor):
        return 0.0
    
    # 计算档位
    gear = calc_gear(rate - base_rate)
    
    # 根据档位计算单价
    unit, _ = calc_unit_price_zc(gear)
    return unit


def _calc_adjusted_unit_pm(rate, actor, base_rate: float) -> float:
    """
    计算项目经理的折算后单价
    
    Args:
        rate: 稳岗率
        actor: 角色名
        base_rate: 基准数
    
    Returns:
        折算后单价
    """
    if pd.isna(rate):
        return 0.0
    # 如果项目经理包含"锦途"，则折算后单价为 0
    if JINTU_KEYWORD in str(actor):
        return 0.0
    
    # 计算档位
    gear = calc_gear(rate - base_rate)
    
    # 根据档位计算单价
    unit, _ = calc_unit_price_pm(gear)
    return unit


def _select_output_columns(out: pd.DataFrame, role_col: str, in_subsidy_col: str) -> pd.DataFrame:
    """
    选择输出列
    
    Args:
        out: 输出数据框
        role_col: 角色列名
        in_subsidy_col: 入职补贴列名
    
    Returns:
        选择列后的数据框
    """
    cols = [
        role_col, "月末在职人数", "当月总在职人数", "稳岗率",
        "当月入职人数", "单价", in_subsidy_col,
        "当月结算总人数", COL_ADJUSTED_UNIT, COL_STABILITY_SUBSIDY
    ]
    out = out[cols]
    return out


def _fill_total_row(out: pd.DataFrame, role_col: str) -> pd.DataFrame:
    """
    填充总计行数据
    
    Args:
        out: 输出数据框
        role_col: 角色列名
    
    Returns:
        更新后的输出数据框
    """
    data_mask = ~out[role_col].isin(EXCLUDE_NAMES) & out[role_col].notna()
    
    # 动态构建总计行，只包含 out 中实际存在的列（避免 pandas 自动创建新列）
    total_row = {}
    sum_cols = [
        "月末在职人数",
        "当月总在职人数",
        "当月入职人数",
        "单价",
        "当月结算总人数",
        COL_IN_SUBSIDY_ZC,      # 项目驻场入职补贴
        COL_IN_SUBSIDY_PM,      # 项目经理入职补贴
        COL_STABILITY_SUBSIDY,
        COL_ADJUSTED_UNIT,
    ]
    for col in sum_cols:
        if col in out.columns:
            total_row[col] = out.loc[data_mask, col].sum()
    
    # 稳岗率列的总计 = 月末在职人数总计 / 当月总在职人数总计
    if "月末在职人数" in out.columns and "当月总在职人数" in out.columns:
        total_month_end = out.loc[data_mask, "月末在职人数"].sum()
        total_on_job = out.loc[data_mask, "当月总在职人数"].sum()
        if total_on_job > 0:
            total_row["稳岗率"] = round_half_up(total_month_end / total_on_job, 2)
        else:
            total_row["稳岗率"] = 0.0
    
    # 更新总计行数据（只更新 total_row 中存在的列，避免自动创建新列）
    total_row_idx = out[out[role_col] == "总计"].index
    if len(total_row_idx) > 0:
        for col, val in total_row.items():
            if col in out.columns:
                out.loc[total_row_idx[0], col] = val
    
    return out


def _clean_special_rows(out: pd.DataFrame, role_col: str) -> pd.DataFrame:
    """
    清理特殊行的数据（空白行、平均数行、中位数行）
    
    Args:
        out: 输出数据框
        role_col: 角色列名
    
    Returns:
        清理后的数据框
    """
    cols = out.columns.tolist()
    
    # 找到空白行、平均数行、中位数行的索引
    for idx in out.index:
        project_field = out.loc[idx, role_col]
        if pd.isna(project_field) or project_field == "平均数" or project_field == "中位数":
            # 只保留稳岗率列的值（如果是平均数或中位数行），其他列设为 NaN
            if project_field == "平均数" or project_field == "中位数":
                for col in cols:
                    if col != role_col and col != "稳岗率":
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
    
    Args:
        stability_zc: 项目驻场稳岗率表
        stability_pm: 项目经理稳岗率表
    
    Returns:
        合并后的稳岗率表
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


def build_base_rate_sheet(zc_base_rate: float, pm_base_rate: float) -> pd.DataFrame:
    """
    构建基准数表，包含项目驻场和项目经理的基准数规则
    左侧：项目驻场；右侧：项目经理
    
    Args:
        zc_base_rate: 项目驻场基准数
        pm_base_rate: 项目经理基准数
    
    Returns:
        基准数表数据框
    """
    # 项目驻场规则（左侧）
    # 档位从 -8 到 10
    rows = []
    
    for gear in range(GEAR_MAX, GEAR_MIN - 1, -1):  # 10, 9, ..., 0, -1, ..., -8
        # 计算对应的稳岗率
        rate_zc = zc_base_rate + gear * GEAR_STEP
        rate_pct_zc = f"{rate_zc*100:.0f}%"
        
        # 计算单价（项目驻场）
        unit_zc, note_zc = calc_unit_price_zc(gear)
        
        # 如果是基准数所在的行（gear=0），在第一列显示基准数
        col1 = zc_base_rate if gear == 0 else ""
        
        # 项目经理部分
        rate_pm = pm_base_rate + gear * GEAR_STEP
        rate_pct_pm = f"{rate_pm*100:.0f}%"
        
        # 计算单价（项目经理）
        unit_pm, note_pm = calc_unit_price_pm(gear)
        
        # 如果是基准数所在的行（gear=0），在第八列显示基准数
        col9 = pm_base_rate if gear == 0 else ""
        
        row = {
            1: col1,                          # 第一列：项目驻场基准数（只在 gear=0 时显示）
            2: gear,                          # 第三列：档位
            3: rate_pct_zc,                   # 第四列：项目驻场稳岗率百分比
            4: unit_zc,                       # 第五列：项目驻场单价
            5: note_zc,                       # 第六列：项目驻场备注
            6: "",                            # 第七列：空白（用于分隔左右）
            8: col9,                          # 第八列：项目经理基准数（只在 gear=0 时显示）
            9: gear,                          # 第十列：档位
            10: rate_pct_pm,                  # 第十一列：项目经理稳岗率百分比
            11: unit_pm,                      # 第十二列：项目经理单价
            12: note_pm,                      # 第十三列：项目经理备注
        }
        rows.append(row)
    
    # 创建 DataFrame
    df = pd.DataFrame(rows)
    
    # 添加标题行（第一行）
    title_row = {i: "" for i in range(BASE_RATE_COL_COUNT)}
    title_row[BASE_RATE_TITLE_COL_ZC] = "项目驻场"
    title_row[BASE_RATE_TITLE_COL_PM] = "项目经理"
    
    # 将标题行添加到数据中
    df_with_title = pd.concat([
        pd.DataFrame([title_row]),
        df
    ], ignore_index=True)
    
    return df_with_title
