# -- coding: utf-8 --
# @Time : 2025-03-12 10:18
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : 拆分_拆分多个表格文件.py
# @Software: PyCharm

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers


def excel_column_to_index(column_letter):
    """将Excel列字母（如A, B, AA, AB）转换为0-based列索引"""
    index = 0
    for char in column_letter.upper():
        if not char.isalpha():
            raise ValueError(f"无效的列字母: {column_letter}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1  # 转换为0-based索引


def apply_date_formatting(worksheet, df):
    """为日期列应用正确的格式"""
    for col_idx, col_name in enumerate(df.columns, 1):
        if pd.api.types.is_datetime64_any_dtype(df[col_name]):
            for row_idx in range(2, len(df) + 2):  # 从第2行开始（第1行是标题）
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if cell.value and isinstance(cell.value, datetime):
                    # 检查是否是纯日期（时间为00:00:00）
                    if cell.value.time() == datetime.min.time():
                        # 纯日期格式
                        cell.number_format = 'yyyy-mm-dd'
                    else:
                        # 日期时间格式
                        cell.number_format = 'yyyy-mm-dd hh:mm:ss'


def auto_adjust_column_widths(file_path):
    """自动调整Excel文件的列宽"""
    try:
        # 加载Excel文件
        wb = load_workbook(file_path)

        # 处理所有工作表
        for ws_name in wb.sheetnames:
            ws = wb[ws_name]

            # 遍历所有列，调整宽度
            for col in ws.columns:
                max_length = 0
                column_letter = get_column_letter(col[0].column)

                # 计算列中单元格的最大宽度
                for cell in col:
                    try:
                        if cell.value is not None and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass

                # 设置列宽（加一点缓冲空间）
                adjusted_width = min((max_length + 2) * 1.2, 50)  # 最大宽度限制为50
                ws.column_dimensions[column_letter].width = adjusted_width

        # 保存调整后的文件
        wb.save(file_path)
    except Exception as e:
        print(f"调整列宽时出错: {e}")


def split_excel_by_column(parent=None):
    """
    按列拆分Excel文件
    :param parent: 父窗口，如果提供则使用Toplevel，否则使用独立的Tk实例
    """
    # 初始化Tkinter - 如果有父窗口则创建Toplevel，否则创建临时Tk
    if parent:
        # 如果有父窗口，创建一个临时的Toplevel用于文件对话框
        dialog_parent = tk.Toplevel(parent)
        dialog_parent.withdraw()
    else:
        # 否则创建独立的Tk实例
        dialog_parent = tk.Tk()
        dialog_parent.withdraw()  # 隐藏Tkinter主窗口

    try:
        # 选择要读取的Excel文件
        file_path = filedialog.askopenfilename(
            title="选择要拆分的Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            parent=dialog_parent if parent else None
        )
        if not file_path:
            messagebox.showerror("错误", "未选择文件")
            return

        # 选择保存路径
        save_dir = filedialog.askdirectory(title="选择保存路径")
        if not save_dir:
            messagebox.showerror("错误", "未选择保存路径")
            return

        # 读取Excel文件
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {e}")
            return

        # 弹出对话框，输入要拆分的列字母（如 A, B, AA, AB）
        columns_input = simpledialog.askstring(
            "输入列字母",
            "请输入要拆分的列（字母，多个列用逗号分隔，例如：A,B,AA）："
        )
        if not columns_input:
            messagebox.showerror("错误", "未输入列字母")
            return

        # 将输入的字母转换为列索引（0-based）
        try:
            columns = [excel_column_to_index(col.strip()) for col in columns_input.split(',') if col.strip()]
        except ValueError as e:
            messagebox.showerror("错误", str(e))
            return

        # 验证列索引是否有效
        for col_idx in columns:
            if col_idx < 0 or col_idx >= len(df.columns):
                messagebox.showerror("错误", f"列 '{columns_input}' 不存在于文件中")
                return

        # 根据指定的列进行拆分
        for col_idx in columns:
            col_name = df.columns[col_idx]  # 获取列名
            unique_values = df[col_name].unique()  # 获取该列的唯一值

            for value in unique_values:
                # 筛选数据
                filtered_df = df[df[col_name] == value]

                # 处理文件名中的特殊字符
                if pd.isna(value):
                    value_str = "空值"
                elif isinstance(value, (pd.Timestamp, datetime)):
                    # 日期类型：格式化为字符串用于文件名
                    value_str = value.strftime('%Y-%m-%d')
                else:
                    value_str = str(value)

                # 清理文件名中的非法字符
                invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
                for char in invalid_chars:
                    value_str = value_str.replace(char, '_')

                # 限制文件名长度
                if len(value_str) > 50:
                    value_str = value_str[:50]

                # 保存到新的Excel文件
                save_path = os.path.join(save_dir, f"{col_name}_{value_str}.xlsx")

                # 使用ExcelWriter来保持日期格式
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Data')

                    # 获取工作表并应用日期格式
                    worksheet = writer.sheets['Data']
                    apply_date_formatting(worksheet, filtered_df)

                # 调整新文件的列宽
                auto_adjust_column_widths(save_path)

        # 显示完成提示并等待用户确认
        messagebox.showinfo("完成", f"文件拆分完成！已生成 {len(columns) * len(unique_values)} 个文件")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现错误: {e}")
    finally:
        # 关闭临时窗口
        if not parent:
            dialog_parent.quit()
            dialog_parent.destroy()


# if __name__ == "__main__":
#     split_excel_by_column()