# -- coding: utf-8 --
# @Time : 2025-09-19 10:54
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cf_zdnrfyq_nrcs_wsx.py
# @Software: PyCharm

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re


def column_letter_to_index(column_letter):
    """将Excel列字母（如A, AA, AAA）转换为列索引（从0开始，pandas格式）"""
    result = 0
    for char in column_letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1  # 转换为0-based索引


def auto_adjust_column_width(worksheet):
    """为不同列设置不同的固定宽度"""
    # 您可以在这里自定义每列的宽度
    column_widths = {
        'A': 10,  # 第一列宽度10
        'B': 21,  # 第二列宽度15
        'C': 29,  # 第三列宽度20
        'D': 21,  # 第四列宽度12
        'E': 37,  # 第五列宽度25
        # 可以继续添加其他列...
    }

    for column in worksheet.columns:
        column_letter = get_column_letter(column[0].column)
        # 如果指定了该列的宽度，使用指定值，否则使用默认值15
        width = column_widths.get(column_letter, 15)
        worksheet.column_dimensions[column_letter].width = width


def main_cf_zdnr_ygbg():
    # 创建Tkinter根窗口并隐藏
    root = tk.Tk()
    root.withdraw()

    try:
        # 1. 选择Excel文件
        file_path = filedialog.askopenfilename(
            title="选择要拆分的Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            print("未选择文件")
            return

        # 获取原文件名（不含扩展名）
        original_filename = os.path.splitext(os.path.basename(file_path))[0]

        # 2. 选择存储文件夹
        save_folder = filedialog.askdirectory(title="选择拆分文件的存储文件夹")
        if not save_folder:
            print("未选择存储文件夹")
            return

        # 3. 输入要拆分的列字母（支持AAA格式）
        column_letter = simpledialog.askstring("输入列字母", "请输入要拆分的列字母（如E, F, AA等）:")
        if not column_letter:
            print("未输入列字母")
            return

        # 验证列字母格式
        if not re.match(r'^[A-Za-z]{1,3}$', column_letter):
            messagebox.showerror("错误", "请输入有效的列字母（1-3个字母）")
            return

        # 读取Excel文件
        print("正在读取Excel文件...")
        excel_file = pd.ExcelFile(file_path)

        # 检查是否存在"拆分参数"工作表
        if '拆分参数' not in excel_file.sheet_names:
            messagebox.showerror("错误", "Excel文件中找不到'拆分参数'工作表")
            return

        # 读取拆分参数工作表
        split_params_df = pd.read_excel(file_path, sheet_name='拆分参数')

        # 创建分组字典 {分组数字: [公司名称列表]}
        split_groups = {}
        for _, row in split_params_df.iterrows():
            company_name = row.iloc[0]  # A列：公司名称
            group_number = row.iloc[1]  # B列：分组数字

            if pd.notna(company_name) and pd.notna(group_number):
                if group_number not in split_groups:
                    split_groups[group_number] = []
                split_groups[group_number].append(company_name)

        print(f"找到拆分参数分组:")
        for group_num, companies in split_groups.items():
            print(f"  分组 {group_num}: {companies}")

        # 获取数据工作表（排除'拆分参数'）
        data_sheet_name = None
        for sheet in excel_file.sheet_names:
            if sheet != '拆分参数':
                data_sheet_name = sheet
                break

        if not data_sheet_name:
            messagebox.showerror("错误", "找不到数据工作表")
            return

        print(f"正在处理工作表: {data_sheet_name}")

        # 读取数据工作表，将所有列作为文本读取，防止长数字变为科学记数法
        # 首先获取所有列名
        df_sample = pd.read_excel(file_path, sheet_name=data_sheet_name, nrows=1)
        column_names = df_sample.columns.tolist()

        # 创建一个字典，指定所有列都以字符串形式读取
        dtype_dict = {col: str for col in column_names}

        # 重新读取整个工作表，所有列作为文本
        df = pd.read_excel(file_path, sheet_name=data_sheet_name, dtype=dtype_dict)
        print(f"数据表总行数: {len(df)}")

        # 转换列字母为列索引
        try:
            column_index = column_letter_to_index(column_letter)
            if column_index >= len(df.columns):
                messagebox.showerror("错误", f"列{column_letter}不存在于工作表中")
                return
        except Exception as e:
            messagebox.showerror("错误", f"列字母转换失败: {str(e)}")
            return

        split_column_name = df.columns[column_index]
        print(f"拆分列: {column_letter} -> {split_column_name}")

        # 显示拆分列的一些样本值
        unique_values = df[split_column_name].dropna().unique()
        print(f"拆分列前10个唯一值: {unique_values[:10]}")

        # 按分组进行拆分和保存
        created_files = []

        for group_num, company_list in split_groups.items():
            print(f"\n正在处理分组 {group_num}:")
            print(f"  目标公司: {company_list}")

            # 简单条件：只要拆分列的内容在目标公司列表中
            condition = df[split_column_name].isin(company_list)
            filtered_df = df[condition]

            print(f"  找到匹配行数: {len(filtered_df)}")

            if not filtered_df.empty:
                # 显示匹配的公司分布
                company_counts = filtered_df[split_column_name].value_counts()
                print(f"  各公司匹配情况:")
                for company, count in company_counts.items():
                    print(f"    - {company}: {count} 行")

                # 创建新的Excel文件，文件名格式为：原文件名_分组数字.xlsx
                output_filename = f"{group_num}_{original_filename}.xlsx"
                output_path = os.path.join(save_folder, output_filename)

                # 保存到新文件，并自动调整列宽
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Data')

                    # 获取工作表并调整列宽
                    worksheet = writer.sheets['Data']
                    auto_adjust_column_width(worksheet)

                    # 设置所有单元格格式为文本，确保长数字不会显示为科学记数法
                    for row in worksheet.iter_rows():
                        for cell in row:
                            cell.number_format = '@'  # '@' 表示文本格式

                created_files.append(output_filename)
                print(f"✓ 已创建文件: {output_filename}")
            else:
                print(f"✗ 未找到匹配分组 {group_num} 的数据")

        # 重新加载文件以删除"拆分参数"工作表
        wb = load_workbook(file_path)
        if '拆分参数' in wb.sheetnames:
            del wb['拆分参数']

        # 统计结果
        total_processed = 0
        for file in created_files:
            file_path = os.path.join(save_folder, file)
            try:
                file_df = pd.read_excel(file_path)
                total_processed += len(file_df)
            except Exception as e:
                print(f"读取文件 {file} 时出错: {e}")

        print(f"\n拆分完成！")
        print(f"  创建文件数: {len(created_files)}")
        print(f"  处理总行数: {total_processed}")
        print(f"  保存位置: {save_folder}")

        messagebox.showinfo("完成",
                            f"文件拆分完成！\n"
                            f"创建了 {len(created_files)} 个分组文件\n"
                            f"处理了 {total_processed} 行数据\n"
                            f"保存位置: {save_folder}")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中发生错误: {str(e)}")
        print(f"错误详情: {str(e)}")
        import traceback
        traceback.print_exc()

    finally:
        root.destroy()


# if __name__ == "__main__":
#     main_cf_zdnr_ygbg()