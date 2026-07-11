"""
从"员工信息汇总-xx年xx月"生成：
1) 项目驻场-刷新
2) 项目经理-刷新
3) 稳岗率
4) 驻场&项目经理-奖励汇总

主入口文件 - 导入各模块完成功能
"""

from pathlib import Path

from .wgl_config import *
from .wgl_io import choose_excel_file, read_workbook, save_output
from .wgl_prepare import prepare_raw, extract_months, get_target_month
from .wgl_refresh import build_refresh
from .wgl_stability import (build_stability_sheet, build_stability_sheet_pm, 
                             build_stability_sheet_combined, build_base_rate_sheet)
from .wgl_reward import build_reward_sheet


def main() -> None:
    """
    主函数：执行稳岗率计算完整流程
    """
    # 1. 选择文件
    input_file = choose_excel_file()
    if input_file is None:
        print("未选择文件，程序退出。")
        return
    
    # 2. 读取数据
    raw = read_workbook(input_file)
    data = prepare_raw(raw)
    
    # 3. 提取月份
    months = extract_months(data)
    if not months:
        print("未找到有效的月份数据！")
        return
    
    # 4. 获取目标月份
    target_month = get_target_month(months)
    start_month = None
    end_month = None
    
    print(f"找到月份：{months}")
    print(f"目标月份：{target_month}")
    
    # 5. 构建刷新表
    print("正在构建项目驻场刷新表...")
    zc_refresh = build_refresh(data, months, "项目驻场", start_month, end_month)
    
    print("正在构建项目经理刷新表...")
    pm_refresh = build_refresh(data, months, "项目经理", start_month, end_month)
    
    # 6. 构建稳岗率表
    print("正在构建项目驻场稳岗率表...")
    stability_zc = build_stability_sheet(data, target_month, zc_refresh)
    
    print("正在构建项目经理稳岗率表...")
    stability_pm = build_stability_sheet_pm(data, target_month, pm_refresh)
    
    # 7. 合并稳岗率表
    print("正在合并稳岗率表...")
    stability_combined = build_stability_sheet_combined(stability_zc.copy(), stability_pm.copy())
    
    # 8. 构建奖励汇总表
    print("正在构建奖励汇总表...")
    reward = build_reward_sheet(stability_zc, stability_pm)
    
    # 9. 提取基准数
    zc_base_rate, pm_base_rate = _extract_base_rates(stability_zc, stability_pm)
    
    # 10. 生成基准数表
    print("正在生成基准数表...")
    base_rate_sheet = build_base_rate_sheet(zc_base_rate, pm_base_rate)
    
    # 11. 保存输出文件
    print("正在保存文件...")
    out_file = save_output(
        input_file,
        {
            SHEET_ZC: zc_refresh,
            SHEET_PM: pm_refresh,
            SHEET_STABILITY: stability_combined,
            SHEET_REWARD: reward,
            SHEET_BASE_RATE: base_rate_sheet,
        },
        selected_sheet=None,  # 简单脚本模式不复制原始工作表
    )
    
    # 12. 输出结果
    print(f"\n处理完成：{out_file}")
    print(f"核算月份：{target_month}")
    print(f"项目驻场基准数：{zc_base_rate:.2f}")
    print(f"项目经理基准数：{pm_base_rate:.2f}")


def _extract_base_rates(stability_zc: "pd.DataFrame", stability_pm: "pd.DataFrame") -> tuple[float, float]:
    """
    从稳岗率表中提取基准数
    
    Args:
        stability_zc: 项目驻场稳岗率表
        stability_pm: 项目经理稳岗率表
    
    Returns:
        (zc_base_rate, pm_base_rate)
    """
    from .wgl_utils import round_half_up
    
    # 项目驻场基准数
    zc_avg_row = stability_zc[stability_zc["项目驻场"] == "平均数"]
    zc_median_row = stability_zc[stability_zc["项目驻场"] == "中位数"]
    
    zc_avg_rate = zc_avg_row["稳岗率"].iloc[0] if not zc_avg_row.empty else 0.0
    zc_median_rate = zc_median_row["稳岗率"].iloc[0] if not zc_median_row.empty else 0.0
    zc_base_rate = max(zc_avg_rate, zc_median_rate)
    
    # 项目经理基准数
    pm_avg_row = stability_pm[stability_pm["项目经理"] == "平均数"]
    pm_median_row = stability_pm[stability_pm["项目经理"] == "中位数"]
    
    pm_avg_rate = pm_avg_row["稳岗率"].iloc[0] if not pm_avg_row.empty else 0.0
    pm_median_rate = pm_median_row["稳岗率"].iloc[0] if not pm_median_row.empty else 0.0
    pm_base_rate = max(pm_avg_rate, pm_median_rate)
    
    return zc_base_rate, pm_base_rate


if __name__ == "__main__":
    main()
