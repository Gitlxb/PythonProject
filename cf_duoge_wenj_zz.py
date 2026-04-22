# -- coding: utf-8 --
# @Time : 2025-09-18 15:35
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cf_duoge_wenj_zz.py
# @Software: PyCharm

import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox, simpledialog
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


def format_date_columns(df):
    """格式化DataFrame中的日期列，只保留年月日"""
    for col in df.columns:
        # 检查列是否是日期类型
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            # 将日期格式化为YYYY-MM-DD字符串，移除时间部分
            df[col] = df[col].dt.strftime('%Y-%m-%d')
        else:
            # 检查字符串列中是否包含日期时间格式
            try:
                # 尝试将字符串列转换为日期时间，然后再格式化为日期
                temp_col = pd.to_datetime(df[col], errors='coerce')
                if not temp_col.isna().all():  # 如果成功转换了一些日期
                    df[col] = temp_col.dt.strftime('%Y-%m-%d')
            except:
                pass
    return df


def auto_adjust_column_widths(file_path):
    """自动调整Excel文件的列宽"""
    try:
        # 加载Excel文件
        wb = load_workbook(file_path)
        ws = wb.active

        # 遍历所有列，调整宽度
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # 获取列字母

            # 计算列中单元格的最大宽度
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            # 设置列宽（加一点缓冲空间）
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width

        # 保存调整后的文件
        wb.save(file_path)
    except Exception as e:
        print(f"调整列宽时出错: {e}")


def save_dataframe_with_text_format(df, file_path):
    """保存DataFrame到Excel，确保长数字以文本格式存储"""
    # 创建一个ExcelWriter对象，指定引擎为openpyxl
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # 获取工作簿和工作表
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # 遍历所有列，检测可能包含长数字的列
        for col_idx, col_name in enumerate(df.columns, 1):
            # 检查列是否可能包含长数字（全数字且长度较长）
            col_data = df[col_name].dropna()
            if len(col_data) > 0:
                # 尝试将值转换为数字，检查是否可能为长数字
                try:
                    sample_value = col_data.iloc[0]
                    if (isinstance(sample_value, (int, float)) and
                            not pd.isna(sample_value) and
                            sample_value > 1e10):
                        # 对于可能的长数字列，设置单元格格式为文本
                        for row_idx in range(2, len(df) + 2):  # 从第2行开始（第1行是标题）
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.number_format = numbers.FORMAT_TEXT
                except:
                    # 如果转换失败，跳过此列
                    pass

    # 再次调整列宽
    auto_adjust_column_widths(file_path)


def split_excel_by_column_zcfgzb():
    # 初始化Tkinter
    root = Tk()
    root.withdraw()  # 隐藏Tkinter主窗口

    # 选择要读取的Excel文件
    file_path = filedialog.askopenfilename(
        title="选择要拆分的Excel文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
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
        # 首先读取列名，确定哪些列可能是日期列
        header_df = pd.read_excel(file_path, nrows=0)
        date_parser_cols = []

        # 创建一个转换器字典，将所有列都作为字符串读取
        converters = {}
        for col in header_df.columns:
            converters[col] = str

        # 使用converters参数确保所有列都以字符串形式读取
        df = pd.read_excel(file_path, converters=converters)

    except Exception as e:
        messagebox.showerror("错误", f"读取文件失败: {e}")
        return

    # 格式化日期列（只保留年月日）
    df = format_date_columns(df)

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

    # 询问是否要对所有拆分后的文件进一步拆分工作表
    split_all_files = messagebox.askyesno("进一步拆分", "是否要对所有拆分后的文件进一步拆分工作表？")

    # 如果选择是，则询问要拆分的列
    split_columns = None
    if split_all_files:
        split_columns_input = simpledialog.askstring(
            "输入工作表拆分列字母",
            "请输入要拆分工作表的列（字母，多个列用逗号分隔）："
        )
        if split_columns_input:
            try:
                split_columns = [excel_column_to_index(col.strip()) for col in split_columns_input.split(',') if
                                 col.strip()]
            except ValueError as e:
                messagebox.showerror("错误", str(e))
                return

    # 根据指定的列进行拆分
    try:
        created_files = []  # 存储创建的文件路径

        for col_idx in columns:
            col_name = df.columns[col_idx]  # 获取列名
            unique_values = df[col_name].unique()  # 获取该列的唯一值

            # 按字母倒序排序唯一值
            unique_values = sorted(unique_values, key=lambda x: str(x).lower(), reverse=True)

            for value in unique_values:
                # 筛选数据
                filtered_df = df[df[col_name] == value]

                # 生成文件名（只保留年月日）
                value_str = str(value)
                # 如果值是日期类型，格式化为年月日
                try:
                    # 先尝试直接解析
                    dt_value = pd.to_datetime(value_str, errors='coerce')
                    if not pd.isna(dt_value):
                        value_str = dt_value.strftime('%Y-%m-%d')
                    else:
                        # 如果包含时间部分，尝试去除
                        if ' ' in value_str:
                            date_part = value_str.split(' ')[0]
                            dt_value = pd.to_datetime(date_part, errors='coerce')
                            if not pd.isna(dt_value):
                                value_str = dt_value.strftime('%Y-%m-%d')
                except:
                    # 如果不是日期，保持原样
                    pass

                # 清理文件名中的非法字符
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    value_str = value_str.replace(char, '_')

                # 保存到新的Excel文件，使用自定义函数确保文本格式
                save_path = os.path.join(save_dir, f"{value_str}.xlsx")
                save_dataframe_with_text_format(filtered_df, save_path)
                created_files.append(save_path)

        # 如果选择了对所有文件进一步拆分工作表
        if split_all_files and split_columns:
            for file_path in created_files:
                try:
                    split_worksheet_in_file(file_path, split_columns)
                except Exception as e:
                    messagebox.showerror("错误", f"拆分文件 '{os.path.basename(file_path)}' 时出错: {e}")

        # 显示完成提示并等待用户确认
        messagebox.showinfo("完成", "文件拆分完成并已自动调整列宽！")
    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现错误: {e}")
    finally:
        # 确保在任何情况下都会显示完成或错误信息，并给用户时间查看
        root.quit()


def split_worksheet_in_file(file_path, columns=None):
    """在已拆分的Excel文件中进一步拆分工作表"""
    try:
        # 读取Excel文件，使用converters确保所有列都以字符串形式读取
        converters = {}
        header_df = pd.read_excel(file_path, nrows=0)
        for col in header_df.columns:
            converters[col] = str

        df = pd.read_excel(file_path, converters=converters)

        # 格式化日期列
        df = format_date_columns(df)

        # 如果没有提供列，则询问用户
        if columns is None:
            columns_input = simpledialog.askstring(
                "输入列字母",
                f"请选择要拆分工作表 '{os.path.basename(file_path)}' 的列（字母，多个列用逗号分隔）："
            )
            if not columns_input:
                return

            columns = [excel_column_to_index(col.strip()) for col in columns_input.split(',') if col.strip()]

        # 验证列索引是否有效
        for col_idx in columns:
            if col_idx < 0 or col_idx >= len(df.columns):
                messagebox.showerror("错误", f"列 '{get_column_letter(col_idx + 1)}' 不存在于文件中")
                return

        # 创建新的工作簿
        wb = load_workbook(file_path)
        original_sheet = wb.active
        original_sheet_name = original_sheet.title

        # 存储创建的工作表信息
        created_sheets = []

        # 根据指定的列进行拆分
        for col_idx in columns:
            col_name = df.columns[col_idx]

            # 只清理要拆分的这一列
            cleaned_series = df[col_name].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

            unique_values = cleaned_series.unique()

            print(f"\n=== 调试信息 ===")
            print(f"列名: {col_name}")
            print(f"清理后的唯一值: {unique_values}")

            for value in unique_values:
                if pd.isna(value) or value == 'nan':
                    continue

                # 使用清理后的值进行匹配，但保持原始数据写入
                matching_indices = cleaned_series[cleaned_series == value].index
                filtered_df = df.loc[matching_indices]

                print(f"值 '{value}' 匹配到 {len(filtered_df)} 行数据")

                if len(filtered_df) == 0:
                    print(f"警告: 值 '{value}' 没有匹配到任何数据，跳过创建工作表")
                    continue

                # 创建工作表名称（只保留年月日）
                sheet_name = str(value)
                # 如果值是日期类型，格式化为年月日
                try:
                    dt_value = pd.to_datetime(sheet_name, errors='coerce')
                    if not pd.isna(dt_value):
                        sheet_name = dt_value.strftime('%Y-%m-%d')
                    else:
                        # 如果包含时间部分，尝试去除
                        if ' ' in sheet_name:
                            date_part = sheet_name.split(' ')[0]
                            dt_value = pd.to_datetime(date_part, errors='coerce')
                            if not pd.isna(dt_value):
                                sheet_name = dt_value.strftime('%Y-%m-%d')
                except:
                    # 如果不是日期，保持原样
                    pass

                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]

                new_sheet = wb.create_sheet(sheet_name)

                # 写入数据
                for c_idx, col_name in enumerate(filtered_df.columns, 1):
                    new_sheet.cell(row=1, column=c_idx, value=col_name)

                for r_idx, row in enumerate(filtered_df.values, 2):
                    for c_idx, cell_value in enumerate(row, 1):
                        cell = new_sheet.cell(row=r_idx, column=c_idx, value=cell_value)
                        # 对于可能的长数字，设置单元格格式为文本
                        try:
                            if (isinstance(cell_value, (int, float)) and
                                    not pd.isna(cell_value) and
                                    cell_value > 1e10):
                                cell.number_format = numbers.FORMAT_TEXT
                            elif (isinstance(cell_value, str) and
                                  cell_value.isdigit() and
                                  len(cell_value) > 10):
                                cell.number_format = numbers.FORMAT_TEXT
                        except:
                            # 如果格式设置失败，继续处理下一个单元格
                            pass

                # 记录工作表信息
                created_sheets.append(sheet_name)

        # 删除原始工作表
        if original_sheet_name in wb.sheetnames:
            del wb[original_sheet_name]

        # 调整工作表位置（按第一个字的字母排序）
        adjust_worksheet_order(wb)

        wb.save(file_path)
        auto_adjust_column_widths(file_path)
        print(f"文件 '{os.path.basename(file_path)}' 的工作表拆分完成！")

    except Exception as e:
        raise Exception(f"拆分工作表时出错: {e}")


def adjust_worksheet_order(wb):
    """调整工作表位置，按第一个字的字母倒序排序（Z在最左边，A在最右边）"""
    try:
        # 获取所有工作表
        sheets = wb.sheetnames

        if len(sheets) <= 1:
            return  # 只有一个工作表不需要排序

        print(f"排序前的工作表顺序: {sheets}")

        # 按第一个字母倒序排序（Z→A）
        sorted_sheets = sorted(sheets, key=lambda x: x[0].lower() if x else '', reverse=False)

        print(f"按字母倒序排序: {sorted_sheets}")

        # 重新排列工作表位置
        # 由于openpyxl的工作表顺序是第一个在最左边，我们直接重新排序
        new_order = []
        for sheet_name in sorted_sheets:
            sheet = wb[sheet_name]
            new_order.append(sheet)

        # 替换原始的工作表列表
        wb._sheets = new_order

        print(f"最终工作表顺序: {[sheet.title for sheet in wb._sheets]}")
        print("工作表位置调整完成！")

    except Exception as e:
        print(f"调整工作表顺序时出错: {e}")


# if __name__ == "__main__":
#     split_excel_by_column_zcfgzb()