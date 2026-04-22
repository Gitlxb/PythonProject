# -- coding: utf-8 --
# @Time : 2025-03-12 9:03
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : 拆分_拆分在表格里.py
# @Software: PyCharm

import pandas as pd
import os
from tkinter import Tk, filedialog, simpledialog, messagebox
from datetime import datetime, date
import re


def excel_column_to_index(column_letter):
    """
    将 Excel 列字母（如 A, B, AA, AB, XFD 等）转换为列索引（从 0 开始）。
    """
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char.upper()) - ord('A') + 1)
    return index - 1  # 转换为从 0 开始的索引


def index_to_excel_column(index):
    """
    将列索引（从 0 开始）转换为 Excel 列字母（如 A, B, AA, AB, XFD 等）。
    """
    column_letter = ""
    while index >= 0:
        column_letter = chr((index % 26) + ord('A')) + column_letter
        index = (index // 26) - 1
    return column_letter


def format_sheet_name(name):
    """
    格式化工作表名称，确保日期和日期时间类型不会触发Excel的自动格式化。
    """
    if pd.isna(name):
        # 处理空值
        return "空值"
    elif isinstance(name, (pd.Timestamp, datetime, date)):
        # 对于日期时间类型，使用格式化字符串
        return name.strftime('%Y-%m-%d')
    else:
        # 其他类型直接转换为字符串
        return str(name)


def is_valid_sheet_name(name):
    """
    检查工作表名称是否有效
    """
    if not name or len(name) > 31:
        return False

    # Excel不允许的字符
    forbidden_chars = ['\\', '/', '*', '[', ']', ':', '?']
    for char in forbidden_chars:
        if char in name:
            return False

    return True


def clean_sheet_name(name):
    """
    清理工作表名称，确保符合Excel要求
    """
    # 替换不允许的字符
    forbidden_chars = ['\\', '/', '*', '[', ']', ':', '?']
    for char in forbidden_chars:
        name = name.replace(char, '_')

    # 截断长度
    if len(name) > 31:
        name = name[:31]

    return name


def split_excel_by_column():
    # 隐藏Tkinter根窗口
    root = Tk()
    root.withdraw()

    # 选择要读取的Excel文件
    file_path = filedialog.askopenfilename(title="选择要读取的Excel文件", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        print("未选择文件，程序退出。")
        return

    # 选择保存路径
    save_path = filedialog.asksaveasfilename(title="选择保存路径", defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        print("未选择保存路径，程序退出。")
        return

    try:
        # 读取Excel文件，保持原始数据类型
        df = pd.read_excel(file_path)

        # 获取列名
        columns = df.columns

        # 显示列名供用户选择
        print("可用的列名：")
        for i, col in enumerate(columns):
            print(f"{index_to_excel_column(i)}: {col}")

        # 用户输入要拆分的列字母（如 A, B, AA, XFD 等）
        col_letter = simpledialog.askstring("输入列字母", "请输入要拆分的列字母（例如：A, B, AA, XFD）：")
        if not col_letter:
            print("未输入列字母，程序退出。")
            return

        col_letter = col_letter.upper()

        # 将列字母转换为列索引
        try:
            col_index = excel_column_to_index(col_letter)
        except ValueError:
            print("输入的列字母无效，程序退出。")
            return

        # 检查列索引是否有效
        if col_index < 0 or col_index >= len(columns):
            print("输入的列字母无效，程序退出。")
            return

        # 获取要拆分的列名
        split_column = columns[col_index]

        # 根据指定列拆分数据
        grouped = df.groupby(split_column)

        # 使用xlsxwriter引擎，它支持日期格式设置
        with pd.ExcelWriter(save_path, engine='xlsxwriter', datetime_format='yyyy-mm-dd',
                            date_format='yyyy-mm-dd') as writer:
            workbook = writer.book

            # 将原始数据写入名为"原始表"的工作表
            df.to_excel(writer, sheet_name="原始表", index=False)

            # 处理每个分组
            sheet_names = set()
            for name, group in grouped:
                # 格式化工作表名称
                sheet_name = format_sheet_name(name)
                sheet_name = clean_sheet_name(sheet_name)

                # 确保工作表名称唯一
                original_name = sheet_name
                counter = 1
                while sheet_name in sheet_names or not is_valid_sheet_name(sheet_name):
                    sheet_name = f"{original_name}_{counter}"
                    counter += 1

                sheet_names.add(sheet_name)

                # 将每个分组写入不同的工作表
                group.to_excel(writer, sheet_name=sheet_name, index=False)

            # 确保至少有一个可见的工作表
            if len(sheet_names) == 0:
                # 如果没有分组，创建一个默认的工作表
                pd.DataFrame().to_excel(writer, sheet_name="数据", index=False)

        # 弹出"处理完成"的提示
        messagebox.showinfo("处理完成", f"文件已成功拆分并保存到：{save_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")
        print(f"错误：{str(e)}")


# if __name__ == "__main__":
#     split_excel_by_column()