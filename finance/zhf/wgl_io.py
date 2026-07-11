"""
稳岗率计算 - IO操作
包含：文件选择、Excel读取、文件保存等IO操作
"""

from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd
from openpyxl import load_workbook, Workbook

from .wgl_config import SHEET_ZC, SHEET_PM


def choose_excel_file() -> Path | None:
    """
    弹出文件选择对话框，选择Excel文件
    
    Returns:
        选择的文件路径，如果取消则返回None
    """
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


def read_workbook(input_file: Path, sheet_name: str = None) -> pd.DataFrame:
    """
    读取 Excel 文件
    
    Args:
        input_file: Excel 文件路径
        sheet_name: 工作表名称，如果为 None 则默认读取第一个工作表
    
    Returns:
        数据框
    """
    if sheet_name is None:
        # 默认读取第一个工作表
        excel_file = pd.ExcelFile(input_file)
        sheet_name = excel_file.sheet_names[0]
        excel_file.close()
    
    raw = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    from .wgl_prepare import normalize_colnames
    raw = normalize_colnames(raw)
    return raw


def save_output(input_file: Path, outputs: dict[str, pd.DataFrame],
                selected_sheet: str = None, copy_original_sheets: bool = True) -> Path:
    """
    保存输出文件
    
    Args:
        input_file: 源 Excel 文件路径
        outputs: 要保存的工作表字典 {sheet_name: DataFrame}
        selected_sheet: 被选中的工作表名称，如果为 None 则不复制原工作表
        copy_original_sheets: 是否复制原始工作表（默认 True，设为 False 可显著提升速度）
    
    Returns:
        保存的文件路径
    """
    import shutil
    
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
    
    # 优化方案：复制临时文件，删除不在白名单里的原始工作表，再添加新的工作表
    if copy_original_sheets:
        try:
            # 白名单：保留的工作表
            KEEP_EXACT = {"运营中心-项目经理", "员工信息汇总"}
            KEEP_PREFIXES = ("代工费-", "百事通报备名单-", "员工信息汇总-")

            # 直接复制整个文件（二进制复制，速度极快）
            shutil.copy2(input_file, out_path)
            print(f"已复制原始文件")

            # 加载复制后的文件
            wb = load_workbook(out_path)

            # 删除不在白名单里的原始工作表
            for sheet_name in list(wb.sheetnames):
                keep = (
                    sheet_name in KEEP_EXACT or
                    any(sheet_name.startswith(prefix) for prefix in KEEP_PREFIXES)
                )
                if not keep:
                    del wb[sheet_name]
                    print(f"  已删除原始工作表: {sheet_name}")
            print(f"  保留的原始工作表: {wb.sheetnames}")

            # 添加新生成的工作表（不删除任何保留的原始工作表）
            for sheet_name, df in outputs.items():
                # 如果同名工作表已存在，删除它
                if sheet_name in wb.sheetnames:
                    del wb[sheet_name]

                ws = wb.create_sheet(title=sheet_name)

                # 将 DataFrame 写入工作表
                if sheet_name in (SHEET_ZC, SHEET_PM, "基准数"):
                    # 项目驻场、项目经理、基准数表都不需要列名
                    _write_dataframe_to_sheet_no_header(ws, df)
                else:
                    # 其他表格需要列名
                    _write_dataframe_to_sheet(ws, df)
            
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
        wb = Workbook()
        # 删除默认的空白 Sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        for sheet_name, df in outputs.items():
            ws = wb.create_sheet(title=sheet_name)
            if sheet_name in (SHEET_ZC, SHEET_PM, "基准数"):
                _write_dataframe_to_sheet_no_header(ws, df)
            else:
                _write_dataframe_to_sheet(ws, df)

        wb.save(out_path)
        wb.close()
    
    return out_path


def _write_dataframe_to_sheet(ws, df: pd.DataFrame) -> None:
    """
    将DataFrame写入工作表（包含列名）
    
    Args:
        ws: openpyxl工作表对象
        df: 要写入的数据框
    """
    # 写入列名
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx, value=col_name)
    # 写入数据
    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for c_idx, value in enumerate(row.values, start=1):
            if pd.notna(value):
                ws.cell(row=r_idx, column=c_idx, value=value)


def _write_dataframe_to_sheet_no_header(ws, df: pd.DataFrame) -> None:
    """
    将DataFrame写入工作表（不包含列名）
    
    Args:
        ws: openpyxl工作表对象
        df: 要写入的数据框
    """
    for r_idx, row in df.iterrows():
        for c_idx, value in enumerate(row.values, start=1):
            if pd.notna(value):
                ws.cell(row=r_idx+1, column=c_idx, value=value)
