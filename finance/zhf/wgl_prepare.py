"""
稳岗率计算 - 数据预处理
包含：列名标准化、日期转换、数据准备等
"""

import re
from decimal import Decimal, ROUND_HALF_UP

import numpy as np
import pandas as pd

from .wgl_config import *


def normalize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    """
    标准化列名（去除前后空格）
    
    Args:
        df: 原始数据框
    
    Returns:
        列名标准化后的数据框
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_month_str(series: pd.Series) -> pd.Series:
    """
    将系列转换为月份字符串格式 (YYYY/MM)
    
    Args:
        series: 输入的系列
    
    Returns:
        月份字符串系列
    """
    s = series.astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "NaT": np.nan})
    # 已是 YYYY/MM
    ok = s.str.match(r"^\d{4}/\d{2}$", na=False)
    out = s.where(ok, np.nan)
    
    dt = pd.to_datetime(s, errors="coerce")
    out = out.fillna(dt.dt.strftime("%Y/%m"))
    return out


def to_date(series: pd.Series) -> pd.Series:
    """
    将系列转换为日期格式
    
    Args:
        series: 输入的系列
    
    Returns:
        日期系列
    """
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        s = series.astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "NaT": np.nan})
        return pd.to_datetime(s, errors="coerce")


def extract_months(raw: pd.DataFrame) -> list[str]:
    """
    从原始数据中提取所有月份
    
    Args:
        raw: 原始数据框
    
    Returns:
        排序后的月份列表
    """
    m1 = to_month_str(raw.get("当月入职", pd.Series(dtype=str)))
    m2 = to_month_str(raw.get("当月离职", pd.Series(dtype=str)))
    months = pd.concat([m1, m2], ignore_index=True).dropna().unique().tolist()
    months = sorted(months)
    return months


def prepare_raw(raw: pd.DataFrame) -> pd.DataFrame:
    """
    准备原始数据（添加计算列）
    
    Args:
        raw: 原始数据框
    
    Returns:
        处理后的数据框
    """
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


def get_target_month(months: list[str]) -> str:
    """
    获取目标月份（默认取最新月份）
    
    Args:
        months: 月份列表
    
    Returns:
        目标月份字符串，如果列表为空则返回空字符串
    """
    if not months:
        return ""
    return months[-1]
