# -- coding: utf-8 --
# @Time : 2025-10-29 10:55
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cw_hbxmtjb_yjl.py
# @Software: PyCharm

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def select_folder_and_process_excel():
    """选择文件夹并处理Excel文件（包含所有子文件夹）"""
    # 创建根窗口并隐藏
    root = tk.Tk()
    root.withdraw()

    # 选择源文件夹
    folder_path = filedialog.askdirectory(title="请选择包含Excel文件的源文件夹")
    if not folder_path:
        print("未选择源文件夹，程序退出")
        return

    print(f"已选择源文件夹: {folder_path}")

    # 选择输出文件保存路径
    output_file = filedialog.asksaveasfilename(
        title="请选择结果文件保存位置",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile="项目统计汇总表.xlsx"
    )

    if not output_file:
        print("未选择保存路径，程序退出")
        return

    print(f"结果文件将保存到: {output_file}")

    # 递归获取文件夹中所有的Excel文件（包括子文件夹）
    excel_files = []
    for root_dir, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(root_dir, file))

    if not excel_files:
        messagebox.showwarning("警告", "选择的文件夹及其子文件夹中没有Excel文件")
        return

    print(f"找到 {len(excel_files)} 个Excel文件（包含所有子文件夹）")

    # 创建新的工作簿用于存储合并结果
    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.title = "项目统计汇总"

    # 设置表头
    headers = [
        "项目名称", "支付费用人数", "商保人数", "社保人数", "收入-项目收入",
        "工人商保", "工人实发工资(工资加需扣合计)", "工人预支", "供应商预支",
        "业务代工费", "业务返费", "个税", "增值税", "增值税附加",
        "发票类型(专票6%，专票5%，差额普票5%、差额普票6%、普票全差额)", "所属公司"
    ]

    # 写入表头
    for col, header in enumerate(headers, 1):
        cell = result_ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 处理每个Excel文件
    processed_count = 0
    skipped_count = 0
    current_row = 2

    for excel_file in excel_files:
        try:
            # 检查文件是否包含"项目统计表"工作表
            if not check_sheet_exists(excel_file, '项目统计表'):
                print(f"跳过文件: {os.path.basename(excel_file)} - 未找到'项目统计表'工作表")
                skipped_count += 1
                continue

            if process_single_excel_file(excel_file, result_ws, current_row):
                processed_count += 1
                print(f"成功处理: {os.path.relpath(excel_file, folder_path)}")
                current_row += 1
            else:
                skipped_count += 1
                print(f"跳过处理: {os.path.relpath(excel_file, folder_path)} - 处理失败")
        except Exception as e:
            skipped_count += 1
            print(f"处理文件 {os.path.relpath(excel_file, folder_path)} 时出错: {str(e)}")

    # 调整格式
    setup_excel_format(result_ws, len(headers))

    # 保存结果文件
    try:
        result_wb.save(output_file)
        result_wb.close()
        print(f"结果文件已保存到: {output_file}")

    except Exception as e:
        error_msg = f"保存文件时出错: {str(e)}"
        print(error_msg)
        messagebox.showerror("错误", error_msg)
        return

    # 显示处理结果统计
    result_message = f"处理完成！\n成功处理: {processed_count} 个文件\n跳过处理: {skipped_count} 个文件\n结果文件保存到: {output_file}"

    # 询问是否打开结果文件
    if processed_count > 0:
        open_file = messagebox.askyesno("处理完成", result_message + "\n\n是否要打开结果文件？")
        if open_file:
            try:
                os.startfile(output_file)  # Windows系统
            except:
                # 如果是其他操作系统，尝试用默认程序打开
                import subprocess
                import platform
                system = platform.system()
                try:
                    if system == "Darwin":  # macOS
                        subprocess.call(["open", output_file])
                    elif system == "Linux":  # Linux
                        subprocess.call(["xdg-open", output_file])
                    else:  # Windows或其他
                        os.startfile(output_file)
                except:
                    print("无法自动打开文件，请手动打开查看")
    else:
        messagebox.showinfo("处理完成", result_message)


def check_sheet_exists(file_path, sheet_name):
    """检查Excel文件中是否存在指定名称的工作表"""
    try:
        # 使用pandas读取Excel文件的所有工作表名称
        xl = pd.ExcelFile(file_path)
        return sheet_name in xl.sheet_names
    except Exception as e:
        print(f"检查工作表时出错 {os.path.basename(file_path)}: {str(e)}")
        return False


def process_single_excel_file(file_path, result_ws, current_row):
    """处理单个Excel文件并将结果写入结果工作表"""
    try:
        # 使用pandas读取Excel文件，不设置列名，直接按位置索引
        df = pd.read_excel(file_path, sheet_name='项目统计表', header=None, nrows=16)

        # 检查数据是否足够
        if df.empty or len(df) < 16:
            print(f"警告: {os.path.basename(file_path)} 中数据不足16行")
            return True

        # 提取A列和B列的数据（按位置索引，不按列名）
        a_column_data = []
        b_column_data = []

        # 读取A1:A16和B1:B16的数据
        for i in range(16):
            # A列数据（索引0）
            a_val = df.iloc[i, 0] if len(df.columns) > 0 and pd.notna(df.iloc[i, 0]) else ""
            a_column_data.append(a_val)

            # B列数据（索引1）
            b_val = df.iloc[i, 1] if len(df.columns) > 1 and pd.notna(df.iloc[i, 1]) else ""
            b_column_data.append(b_val)

        # 1. 将A1单元格内容（项目名称）放在第1列
        project_name = a_column_data[0] if len(a_column_data) > 0 else ""
        result_ws.cell(row=current_row, column=1, value=project_name)

        # 2. 将B1:B15进行转置，从第2列开始向右排列（B2:O2）
        for i in range(min(14, len(b_column_data))):
            result_ws.cell(row=current_row, column=i + 2, value=b_column_data[i])

        # 3. 将A2:A16进行转置，从B2开始向右添加（B2:O2），与B列数据在同一行
        # 注意：这里A2:A16有15个数据，但B2:O2只有14个位置，所以只取前14个
        for i in range(1, min(16, len(a_column_data))):  # 从A2开始，共15个数据
            if i <= 16:  # 只取前14个数据，放到B2:O2
                result_ws.cell(row=current_row, column=i + 1, value=a_column_data[i])

        return True

    except Exception as e:
        error_msg = f"处理文件 {os.path.basename(file_path)} 时发生错误: {str(e)}"
        print(error_msg)
        return False


def setup_excel_format(worksheet, num_columns):
    """设置Excel格式"""
    # 设置列宽
    column_widths = [15, 12, 10, 10, 15, 10, 20, 10, 12, 12, 10, 10, 10, 12, 25]

    for col, width in enumerate(column_widths, 1):
        if col <= len(column_widths):
            worksheet.column_dimensions[get_column_letter(col)].width = width

    # 设置行高
    worksheet.row_dimensions[1].height = 30

    # 设置边框
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    # 为所有有数据的单元格添加边框
    max_col = worksheet.max_column
    max_row = worksheet.max_row
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.border = thin_border
            if cell.value is not None:
                cell.alignment = Alignment(horizontal='center', vertical='center')


def get_column_letter(col_idx):
    """将列索引转换为字母"""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def main():
    """主函数"""
    print("Excel文件处理程序启动")
    print("功能说明:")
    print("1. 首先选择包含Excel文件的源文件夹")
    print("2. 然后选择结果文件的保存路径和文件名")
    print("3. 程序会递归读取源文件夹及其所有子文件夹中每个Excel文件中'项目统计表'的A1:B16数据")
    print("4. 如果Excel文件中没有'项目统计表'工作表，将自动跳过该文件")
    print("5. 按照转置逻辑处理数据：")
    print("   - A1放在第1列（项目名称）")
    print("   - B1:B15转置后从第2列开始向右排列")
    print("   - A2:A16转置后从B2开始向右添加（B2:O2）")
    print("6. 生成项目统计汇总表")

    try:
        select_folder_and_process_excel()
    except Exception as e:
        error_msg = f"程序执行过程中发生错误: {str(e)}"
        print(error_msg)
        messagebox.showerror("错误", error_msg)
    finally:
        print("程序执行完毕")


if __name__ == "__main__":
    main()