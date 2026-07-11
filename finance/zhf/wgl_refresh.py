"""
稳岗率计算 - 刷新表构建
包含：build_refresh 函数及其辅助函数
"""

import numpy as np
import pandas as pd

from .wgl_config import *
from .wgl_utils import is_valid_settlement_value, round_half_up


def build_refresh(df: pd.DataFrame, months: list[str], role_col: str, 
                  start_month: str, end_month: str) -> pd.DataFrame:
    """
    构建刷新表（项目驻场或项目经理）
    
    Args:
        df: 原始数据框
        months: 月份列表
        role_col: 角色列名（"项目驻场"或"项目经理"）
        start_month: 起始月份
        end_month: 结束月份
    
    Returns:
        刷新表数据框
    """
    # 第一重过滤：筛选"总人数"列为有效值的行（全局过滤，影响所有后续统计）
    total_col = None
    for col in df.columns:
        if "总人数" in str(col):
            total_col = col
            break
    
    df_filtered = df.copy()
    if total_col is not None:
        original_len = len(df_filtered)
        df_filtered = df_filtered[df_filtered[total_col].apply(is_valid_settlement_value)].copy()
        print(f"[build_refresh] 总人数过滤前：{original_len} 行，过滤后：{len(df_filtered)} 行")
    else:
        print("[build_refresh] 警告：未找到包含'总人数'的列，使用所有数据")
    
    actors = sorted([x for x in df_filtered[role_col].dropna().unique().tolist() if str(x).strip()])
    
    # 当月入职统计
    in_tmp = _build_in_stats(df_filtered, months, role_col, actors)
    
    # 当月离职统计
    out_tmp = _build_out_stats(df_filtered, months, role_col, actors)
    
    # 计算稳岗率
    stability_rows = _build_stability_rows(df_filtered, in_tmp, out_tmp, months, actors, role_col, start_month, end_month)
    
    # 拼接所有部分
    rows = _build_final_rows(in_tmp, out_tmp, months, actors, role_col, stability_rows)
    
    out = pd.DataFrame(rows)
    return out


def _build_in_stats(df: pd.DataFrame, months: list[str], role_col: str, actors: list) -> pd.DataFrame:
    """
    构建入职统计表
    
    Args:
        df: 过滤后的数据框
        months: 月份列表
        role_col: 角色列名
        actors: 角色列表
    
    Returns:
        入职统计表
    """
    in_tmp = (
        df.dropna(subset=["当月入职_月", role_col])
        .groupby(["当月入职_月", role_col], dropna=False)
        .size()
        .unstack(fill_value=0)
        .reindex(index=months, fill_value=0)
    )
    if actors:
        in_tmp = in_tmp.reindex(columns=actors, fill_value=0)
    in_tmp["总计"] = in_tmp.sum(axis=1)
    return in_tmp


def _build_out_stats(df: pd.DataFrame, months: list[str], role_col: str, actors: list) -> pd.DataFrame:
    """
    构建离职统计表
    
    Args:
        df: 过滤后的数据框
        months: 月份列表
        role_col: 角色列名
        actors: 角色列表
    
    Returns:
        离职统计表
    """
    out_tmp = (
        df.dropna(subset=["当月离职_月", role_col])
        .groupby(["当月离职_月", role_col], dropna=False)
        .size()
        .unstack(fill_value=0)
        .reindex(index=months, fill_value=0)
    )
    if actors:
        out_tmp = out_tmp.reindex(columns=actors, fill_value=0)
    out_tmp["总计"] = out_tmp.sum(axis=1)
    return out_tmp


def _build_stability_rows(df: pd.DataFrame, in_tmp: pd.DataFrame, out_tmp: pd.DataFrame, 
                          months: list[str], actors: list, role_col: str,
                          start_month: str, end_month: str) -> list[dict]:
    """
    构建稳岗率计算行
    
    Args:
        df: 过滤后的数据框
        in_tmp: 入职统计表
        out_tmp: 离职统计表
        months: 月份列表
        actors: 角色列表
        role_col: 角色列名
        start_month: 起始月份
        end_month: 结束月份
    
    Returns:
        稳岗率行列表
    """
    # 使用指定的时间范围
    target_months = [m for m in months if m >= start_month and m <= end_month]
    
    # 计算离职月份（不包含结束月份）
    out_months_until_dec = [m for m in target_months if m < end_month]
    
    stability_rows = []
    
    # 第一行：时间 + 姓名标题
    name_row = {"指标": end_month}
    for actor in actors:
        name_row[actor] = actor
    name_row["总计"] = "总计"
    stability_rows.append(name_row)
    
    # 第二行：月末在职人数
    month_end_row = {"指标": "月末在职人数"}
    for actor in actors:
        in_total = in_tmp.loc[target_months, actor].sum() if actor in in_tmp.columns else 0
        out_total = out_tmp.loc[target_months, actor].sum() if actor in out_tmp.columns else 0
        month_end_row[actor] = in_total - out_total
    month_end_row["总计"] = sum(month_end_row[actor] for actor in actors)
    stability_rows.append(month_end_row)
    
    # 第三行：当月总在职人数
    total_on_job_row = {"指标": "当月总在职人数"}
    for actor in actors:
        in_total = in_tmp.loc[target_months, actor].sum() if actor in in_tmp.columns else 0
        out_total = out_tmp.loc[out_months_until_dec, actor].sum() if actor in out_tmp.columns else 0
        total_on_job_row[actor] = in_total - out_total
    total_on_job_row["总计"] = sum(total_on_job_row[actor] for actor in actors)
    stability_rows.append(total_on_job_row)
    
    # 第四行：稳岗率
    stability_rate_row = {"指标": "稳岗率"}
    for actor in actors:
        month_end = month_end_row[actor]
        total_on_job = total_on_job_row[actor]
        if total_on_job > 0:
            rate = month_end / total_on_job
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
    
    return stability_rows


def _build_final_rows(in_tmp: pd.DataFrame, out_tmp: pd.DataFrame, 
                      months: list[str], actors: list, role_col: str,
                      stability_rows: list[dict]) -> list[dict]:
    """
    构建最终输出行
    
    Args:
        in_tmp: 入职统计表
        out_tmp: 离职统计表
        months: 月份列表
        actors: 角色列表
        role_col: 角色列名
        stability_rows: 稳岗率行列表
    
    Returns:
        最终输出行列表
    """
    rows = []
    
    # 当月入职行
    month_in_row = {"指标": "当月入职"}
    for actor in actors:
        month_in_row[actor] = actor
    month_in_row["总计"] = "总计"
    rows.append(month_in_row)
    
    for m in months:
        rec = {"指标": m}
        rec.update(in_tmp.loc[m].to_dict())
        rows.append(rec)
    rec_total_in = {"指标": "总计"}
    rec_total_in.update(in_tmp.sum(axis=0).to_dict())
    rows.append(rec_total_in)
    
    rows.append({"指标": ""})
    
    # 当月离职行
    month_out_row = {"指标": "当月离职"}
    for actor in actors:
        month_out_row[actor] = actor
    month_out_row["总计"] = "总计"
    rows.append(month_out_row)
    
    for m in months:
        rec = {"指标": m}
        rec.update(out_tmp.loc[m].to_dict())
        rows.append(rec)
    rec_total_out = {"指标": "总计"}
    rec_total_out.update(out_tmp.sum(axis=0).to_dict())
    rows.append(rec_total_out)
    
    # 添加稳岗率部分
    rows.append({"指标": ""})
    rows.extend(stability_rows)
    
    return rows
