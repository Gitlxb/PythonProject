"""
稳岗率计算 - 工具函数
包含：四舍五入、数据验证等通用工具函数
"""

import math
from decimal import Decimal, ROUND_HALF_UP

import numpy as np
import pandas as pd


def round_half_up(value, decimals=2):
    """
    四舍五入（银行家舍入法，与Excel一致）
    
    Args:
        value: 要四舍五入的值
        decimals: 保留小数位数
    
    Returns:
        四舍五入后的浮点数
    """
    if pd.isna(value) or value is None:
        return 0.0
    d = Decimal(str(value))
    rounded = d.quantize(Decimal(10) ** -decimals, rounding=ROUND_HALF_UP)
    return float(rounded)


def is_valid_settlement_value(val) -> bool:
    """
    判断工厂实际结算列是否为有效值（非空、非乱码）
    
    Args:
        val: 要检查的值
    
    Returns:
        是否有效
    """
    if pd.isna(val) or val == '':
        return False
    str_val = str(val).strip()
    if len(str_val) == 0:
        return False
    if len(str_val) > 50:
        return False
    return True


def is_valid_total_number(val) -> bool:
    """
    判断总人数列是否为有效数字值（排除空白、乱码、非数字内容）
    
    Args:
        val: 要检查的值
    
    Returns:
        是否为有效数字
    """
    if pd.isna(val):
        return False
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def find_settlement_column(df: pd.DataFrame) -> str | None:
    """
    动态查找包含"工厂实际结算"字样的列
    
    Args:
        df: 数据框
    
    Returns:
        列名，如果未找到返回None
    """
    for col in df.columns:
        if "工厂实际结算" in str(col):
            return col
    return None


def calc_gear(diff: float) -> int:
    """
    计算档位（与Excel INT函数保持一致）
    
    Args:
        diff: 稳岗率与基准数的差值
    
    Returns:
        档位（整数）
    """
    return math.floor(round(diff / 0.05, 10))


def calc_unit_price_zc(gear: int) -> tuple[float, str]:
    """
    计算项目驻场的单价
    
    Args:
        gear: 档位
    
    Returns:
        (单价, 备注)
    """
    if gear >= 6:
        return 10.0, "≥8元时，实际按照10元结算" if gear == 6 else ""
    elif gear < -5:
        return 0.0, "＜2元时，不结算" if gear == -6 else ""
    else:
        return gear * 0.5 + 5, ""


def calc_unit_price_pm(gear: int) -> tuple[float, str]:
    """
    计算项目经理的单价
    
    Args:
        gear: 档位
    
    Returns:
        (单价, 备注)
    """
    if gear >= 6:
        return 2.0, "≥1.6元时，实际按照2元结算" if gear == 6 else ""
    elif gear < -5:
        return 0.0, "＜0.4元时，不结算" if gear == -6 else ""
    else:
        return gear * 0.1 + 1, ""
