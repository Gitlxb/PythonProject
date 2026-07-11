"""
稳岗率计算 - 奖励汇总模块
包含：build_reward_sheet 函数
"""

import pandas as pd

from .wgl_config import *


def build_reward_sheet(stability_zc: pd.DataFrame, stability_pm: pd.DataFrame) -> pd.DataFrame:
    """
    构建奖励汇总表
    表结构：姓名 | 入职补贴（5元/人）| 稳岗补贴_驻场 | 入职补贴（1元/人）| 稳岗补贴_经理 | 合计 | 备注
    数据来源：
    - 入职补贴（5元/人）和稳岗补贴_驻场：来自稳岗率表项目驻场部分
    - 入职补贴（1元/人）和稳岗补贴_经理：来自稳岗率表项目经理部分
    - 合计：四个补贴列的总和
    - 备注：
    
    Args:
        stability_zc: 项目驻场稳岗率表
        stability_pm: 项目经理稳岗率表
    
    Returns:
        奖励汇总表数据框
    """
    # 从项目驻场部分提取数据
    zc_data = _extract_zc_data(stability_zc)
    
    # 从项目经理部分提取数据
    pm_data = _extract_pm_data(stability_pm)
    
    # 合并驻场和项目经理数据
    reward_merged = _merge_reward_data(zc_data, pm_data)
    
    # 计算合计列
    reward_merged = _calc_total_reward(reward_merged)
    
    # 调整列顺序并重命名
    reward_merged = _reorder_and_rename_columns(reward_merged)
    
    # 添加总计行
    reward_merged = _add_total_row(reward_merged)
    
    return reward_merged


def _extract_zc_data(stability_zc: pd.DataFrame) -> pd.DataFrame:
    """
    从项目驻场稳岗率表中提取数据
    
    Args:
        stability_zc: 项目驻场稳岗率表
    
    Returns:
        提取后的数据框
    """
    valid_rows_zc = stability_zc[
        ~stability_zc["项目驻场"].isin(EXCLUDE_NAMES) &
        ~stability_zc["项目驻场"].str.contains("对比结果", na=False) &
        stability_zc["项目驻场"].notna()
        ].copy()
    
    zc_data = valid_rows_zc[
        ["项目驻场", COL_IN_SUBSIDY_ZC, COL_STABILITY_SUBSIDY]].copy()
    zc_data.columns = ["姓名", "入职补贴_驻场", "稳岗补贴_驻场"]
    zc_data = zc_data.fillna(0)
    
    # 确保姓名列都是字符串类型
    zc_data["姓名"] = zc_data["姓名"].astype(str)
    
    return zc_data


def _extract_pm_data(stability_pm: pd.DataFrame) -> pd.DataFrame:
    """
    从项目经理稳岗率表中提取数据
    
    Args:
        stability_pm: 项目经理稳岗率表
    
    Returns:
        提取后的数据框
    """
    valid_rows_pm = stability_pm[
        ~stability_pm["项目经理"].isin(EXCLUDE_NAMES) &
        ~stability_pm["项目经理"].str.contains("这是对比的结果", na=False) &
        stability_pm["项目经理"].notna()
        ].copy()
    
    pm_data = valid_rows_pm[
        ["项目经理", COL_IN_SUBSIDY_PM, COL_STABILITY_SUBSIDY]].copy()
    pm_data.columns = ["姓名", "入职补贴_经理", "稳岗补贴_经理"]
    pm_data = pm_data.fillna(0)
    
    # 确保姓名列都是字符串类型
    pm_data["姓名"] = pm_data["姓名"].astype(str)
    
    return pm_data


def _merge_reward_data(zc_data: pd.DataFrame, pm_data: pd.DataFrame) -> pd.DataFrame:
    """
    合并驻场和项目经理数据
    
    Args:
        zc_data: 项目驻场数据
        pm_data: 项目经理数据
    
    Returns:
        合并后的数据框
    """
    # 合并驻场和项目经理数据（全外连接，同一个姓名可能在两表中都存在）
    reward_merged = pd.merge(zc_data, pm_data, on="姓名", how="outer").fillna(0)
    return reward_merged


def _calc_total_reward(reward_merged: pd.DataFrame) -> pd.DataFrame:
    """
    计算合计列（四个补贴列的总和）
    
    Args:
        reward_merged: 合并后的数据框
    
    Returns:
        添加合计列后的数据框
    """
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
    
    return reward_merged


def _reorder_and_rename_columns(reward_merged: pd.DataFrame) -> pd.DataFrame:
    """
    调整列顺序并重命名
    
    Args:
        reward_merged: 数据框
    
    Returns:
        调整列顺序后的数据框
    """
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
        "入职补贴_驻场": COL_IN_SUBSIDY_ZC,
        "稳岗补贴_驻场": "稳岗补贴_驻场",
        "入职补贴_经理": COL_IN_SUBSIDY_PM,
        "稳岗补贴_经理": "稳岗补贴_经理"
    })
    
    return reward_merged


def _add_total_row(reward_merged: pd.DataFrame) -> pd.DataFrame:
    """
    添加总计行
    
    Args:
        reward_merged: 数据框
    
    Returns:
        添加总计行后的数据框
    """
    # 添加总计行
    total_row = {
        "姓名": "总计",
        COL_IN_SUBSIDY_ZC: reward_merged[COL_IN_SUBSIDY_ZC].sum(),
        "稳岗补贴_驻场": reward_merged["稳岗补贴_驻场"].sum(),
        COL_IN_SUBSIDY_PM: reward_merged[COL_IN_SUBSIDY_PM].sum(),
        "稳岗补贴_经理": reward_merged["稳岗补贴_经理"].sum(),
        "合计": reward_merged["合计"].sum(),
        "备注": ""
    }
    
    # 使用 concat 添加总计行
    total_df = pd.DataFrame([total_row])
    reward_merged = pd.concat([reward_merged, total_df], ignore_index=True)
    
    return reward_merged
