# -- coding: utf-8 --
# @Time : 2025-11-19 15:35
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cw_bxcfyqh_xzy.py
# @Software: PyCharm

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings

warnings.filterwarnings('ignore')


class ExcelProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel文件处理器")
        self.root.geometry("600x500")
            
        # 设置窗口居中
        self.center_window(self.root, 600, 500)
    
        self.file_path = None
        self.original_df = None
        self.processed_df = None
        self.sheet_name = None
        self.original_wb = None  # 保存原始工作簿对象
    
        self.setup_ui()

    def center_window(self, window, width, height):
        """设置窗口居中"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        """设置用户界面"""
        # 文件选择区域
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10)

        tk.Button(file_frame, text="选择Excel文件", command=self.select_file,
                  bg="#4CAF50", fg="white", font=("Arial", 12)).pack(pady=5)

        self.file_label = tk.Label(file_frame, text="未选择文件", wraplength=500)
        self.file_label.pack(pady=5)

        # 处理按钮
        process_frame = tk.Frame(self.root)
        process_frame.pack(pady=10)

        tk.Button(process_frame, text="1. 处理Excel文件（前三步）", command=self.process_excel,
                  bg="#2196F3", fg="white", font=("Arial", 12), state="disabled").pack(pady=5)
        self.process_btn = process_frame.winfo_children()[0]

        # 列选择区域
        column_frame = tk.Frame(self.root)
        column_frame.pack(pady=10, fill="both", expand=True)

        tk.Label(column_frame, text="列预览界面（处理后）", font=("Arial", 14, "bold")).pack()

        # 创建树形视图显示列预览
        self.tree = ttk.Treeview(column_frame, columns=('Index', 'ColumnName'), show='headings', height=10)
        self.tree.heading('Index', text='序号')
        self.tree.heading('ColumnName', text='列名')

        scrollbar = ttk.Scrollbar(column_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")

        # 拆分按钮
        split_frame = tk.Frame(self.root)
        split_frame.pack(pady=10)

        tk.Button(split_frame, text="2. 选择拆分列并执行", command=self.show_column_selector,
                  bg="#FF9800", fg="white", font=("Arial", 12), state="disabled").pack(pady=5)
        self.split_btn = split_frame.winfo_children()[0]

    def select_file(self):
        """选择Excel文件"""
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if self.file_path:
            self.file_label.config(text=f"已选择: {os.path.basename(self.file_path)}")
            self.process_btn.config(state="normal")

            # 预览原始文件的列名
            try:
                # 使用openpyxl加载原始工作簿以保留格式
                self.original_wb = load_workbook(self.file_path)
                self.sheet_name = self.original_wb.sheetnames[0]

                # 使用pandas预览列名
                df_preview = pd.read_excel(self.file_path, sheet_name=0, nrows=1)
                self.preview_columns(df_preview.columns.tolist(), "原始文件列预览")
            except Exception as e:
                messagebox.showerror("错误", f"读取文件失败: {str(e)}")

    def preview_columns(self, columns, title="列预览"):
        """预览列名"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 更新标题
        for child in self.root.winfo_children():
            if isinstance(child, tk.Frame) and hasattr(child, 'winfo_children'):
                for grandchild in child.winfo_children():
                    if isinstance(grandchild, tk.Label) and "列预览" in grandchild.cget("text"):
                        grandchild.config(text=title)
                        break

        # 添加新数据
        for i, col in enumerate(columns, 1):
            self.tree.insert("", "end", values=(i, col))

    def process_excel(self):
        """处理Excel文件：完成前三步操作"""
        try:
            # 重新加载原始工作簿
            self.original_wb = load_workbook(self.file_path)
            self.sheet_name = self.original_wb.sheetnames[0]
            ws = self.original_wb[self.sheet_name]

            # ---------- 步骤1：拆分所有合并单元格 ----------
            merged_ranges = list(ws.merged_cells.ranges)
            for merged_range in merged_ranges:
                # 获取合并区域的范围
                min_row, min_col, max_row, max_col = (
                    merged_range.min_row, merged_range.min_col,
                    merged_range.max_row, merged_range.max_col
                )
                # 获取左上角单元格的值
                top_left_value = ws.cell(row=min_row, column=min_col).value
                # 取消合并
                ws.unmerge_cells(str(merged_range))
                # 填充值到所有单元格
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        ws.cell(row=row, column=col).value = top_left_value

            # ---------- 步骤2：拆分第二行的多字段单元格 ----------
            max_col = ws.max_column
            for col in range(max_col, 0, -1):
                cell = ws.cell(row=2, column=col)
                if cell.value and isinstance(cell.value, str) and ' ' in cell.value.strip():
                    fields = [f.strip() for f in cell.value.split(' ') if f.strip()]
                    if len(fields) > 1:
                        # 在当前位置右侧插入新列
                        ws.insert_cols(col + 1, amount=len(fields) - 1)
                        # 设置原单元格为第一个字段
                        ws.cell(row=2, column=col).value = fields[0]
                        # 设置新插入的列单元格
                        for i, field in enumerate(fields[1:], start=1):
                            ws.cell(row=2, column=col + i).value = field

            # ---------- 步骤3：删除第一行 ----------
            ws.delete_rows(1)

            # 保存处理后的临时文件，以便pandas读取
            temp_file = "temp_processed.xlsx"
            self.original_wb.save(temp_file)
            self.original_wb.close()  # 关闭工作簿，因为后面会重新加载

            # 读取处理后的数据为DataFrame
            self.processed_df = pd.read_excel(temp_file, sheet_name=self.sheet_name)

            # 删除临时文件
            if os.path.exists(temp_file):
                os.remove(temp_file)

            # ---------- 步骤4：移动"申请报销金额"列到C列右侧 ----------
            if "申请报销金额" in self.processed_df.columns:
                amount_col = self.processed_df["申请报销金额"]
                self.processed_df = self.processed_df.drop("申请报销金额", axis=1)

                # 插入到C列右边（第3列之后，索引为3）
                if len(self.processed_df.columns) >= 3:
                    self.processed_df.insert(3, "申请报销金额", amount_col)
                else:
                    self.processed_df["申请报销金额"] = amount_col

            # 重新加载原始工作簿（用于后续复制原始数据）
            self.original_wb = load_workbook(self.file_path)
            self.sheet_name = self.original_wb.sheetnames[0]

            messagebox.showinfo("处理完成",
                                "前三步处理已完成：\n"
                                "✓ 所有合并单元格已拆分并填充\n"
                                "✓ 第二行中由空格分隔的字段已拆分到多列\n"
                                "✓ 第一行已删除\n"
                                "✓ 申请报销金额列已移动到C列右侧\n\n"
                                "现在请选择拆分列进行下一步操作。")

            # 启用拆分按钮
            self.split_btn.config(state="normal")

            # 预览处理后的列
            self.preview_columns(self.processed_df.columns.tolist(), "处理后列预览（请选择拆分列）")

        except Exception as e:
            messagebox.showerror("错误", f"处理文件失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def simplify_account_name(self, name):
        """将支付账户的长名称简写为短名称"""
        # 定义映射字典
        mapping = {
            "刘先锋个人现金卡": "刘先锋现金卡",
            "温州鑫锦途供应链服务有限公司": "温州鑫锦途",
            "浙江锦途人力资源有限公司": "浙江锦途",
            "温州锦途服务外包有限公司": "温州锦途",
            "衢州鑫途服务外包有限公司": "衢州鑫途",
            "浙江锦途人力资源有限公司乐清分公司": "浙江锦途（乐清分公司）",
            "芜湖才库人力资源有限公司": "芜湖才库",
            "衢州驰锦企业服务有限公司": "衢州驰锦",
            "安徽锦途企业服务集团有限公司": "安徽锦途",
            "安庆市同益劳务服务有限公司": "安庆同益",
            "台州众通人才服务有限公司": "台州众通",
            "台州锦途人力资源服务外包有限公司": "台州锦途",
            "温州鑫途企业服务有限公司": "温州鑫途",
            "温州驰锦企业服务有限公司": "温州驰锦",
            "浙江盛威安保服务有限公司": "盛威安保",
            "温州市合静人力资源有限公司": "合静",
            "铜陵锦途劳务有限公司": "铜陵锦途"
        }
        # 如果名称在映射中，返回简写，否则返回原名称
        return mapping.get(name, name)

    def show_column_selector(self):
        """显示列选择器界面"""
        if self.processed_df is None:
            messagebox.showwarning("警告", "请先完成前三步处理")
            return

        # 创建列选择对话框
        selector = ColumnSelector(self.root, self.processed_df.columns.tolist(), self.execute_split)
        self.root.wait_window(selector.window)

    def execute_split(self, split_column):
        """执行拆分操作"""
        try:
            # 检查列是否存在
            if split_column not in self.processed_df.columns:
                messagebox.showerror("错误", f"列 '{split_column}' 不存在")
                return

            # 选择保存路径
            save_path = filedialog.asksaveasfilename(
                title="保存拆分后的文件",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )

            if not save_path:
                return

            # 创建新的工作簿
            new_wb = Workbook()
            # 删除默认创建的工作表
            new_wb.remove(new_wb.active)

            # 1. 复制原始数据工作表（完全保留格式）
            if self.original_wb is not None:
                original_ws = self.original_wb[self.sheet_name]
                original_copy = new_wb.create_sheet("原始数据")

                # 复制所有内容和格式
                for row in original_ws.iter_rows():
                    for cell in row:
                        new_cell = original_copy.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            new_cell.font = cell.font.copy()
                            new_cell.border = cell.border.copy()
                            new_cell.fill = cell.fill.copy()
                            new_cell.number_format = cell.number_format
                            new_cell.protection = cell.protection.copy()
                            new_cell.alignment = cell.alignment.copy()

                # 复制列宽
                for col_letter, dimension in original_ws.column_dimensions.items():
                    if dimension.width is not None:
                        original_copy.column_dimensions[col_letter].width = dimension.width

                # 复制行高
                for row, dimension in original_ws.row_dimensions.items():
                    if dimension.height is not None:
                        original_copy.row_dimensions[row].height = dimension.height

                # 复制合并单元格
                for merged_range in original_ws.merged_cells.ranges:
                    original_copy.merge_cells(str(merged_range))

            # 2. 添加处理后数据工作表
            processed_ws = new_wb.create_sheet("处理后数据")
            for r_idx, row in enumerate(dataframe_to_rows(self.processed_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    processed_ws.cell(row=r_idx, column=c_idx, value=value)
            
            # 对处理后数据工作表添加合计行并填充黄色
            if "申请报销金额" in self.processed_df.columns:
                # 计算合计值
                total_amount = self.processed_df["申请报销金额"].sum()
                # 在最后一行后添加合计行（数据行数 +1 为表头，+1 为合计行）
                total_row_idx = len(self.processed_df) + 2
                # "合计"写在 C 列（第 3 列）
                processed_ws.cell(row=total_row_idx, column=3, value="合计")
                # 金额写在 D 列（第 4 列）
                processed_ws.cell(row=total_row_idx, column=4, value=total_amount)
                
                # 填充黄色背景
                from openpyxl.styles import PatternFill
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                processed_ws.cell(row=total_row_idx, column=3).fill = yellow_fill
                processed_ws.cell(row=total_row_idx, column=4).fill = yellow_fill

            # 3. 按选定列拆分数据并创建工作表
            grouped = self.processed_df.groupby(split_column)

            for group_name, group_data in grouped:
                # 创建每个组的副本
                sheet_df = group_data.copy().reset_index(drop=True)

                # 添加合计行
                if "申请报销金额" in sheet_df.columns:
                    # 确保金额列是数值类型
                    sheet_df["申请报销金额"] = pd.to_numeric(sheet_df["申请报销金额"], errors='coerce').fillna(0)

                    # 添加合计行
                    total_row = {}
                    for col in sheet_df.columns:
                        if col == "报销成员":
                            total_row[col] = "合计"
                        elif col == "申请报销金额":
                            total_row[col] = sheet_df[col].sum()
                        else:
                            total_row[col] = ""

                    total_row_df = pd.DataFrame([total_row])
                    sheet_df = pd.concat([sheet_df, total_row_df], ignore_index=True)

                    # 添加三行空行
                    for _ in range(3):
                        empty_row = {col: "" for col in sheet_df.columns}
                        sheet_df = pd.concat([sheet_df, pd.DataFrame([empty_row])], ignore_index=True)

                    # 获取去重的报销成员（排除空值、标题和合计）
                    if "报销成员" in sheet_df.columns:
                        unique_members = group_data["报销成员"].dropna().unique()
                        unique_members = [m for m in unique_members if m != "合计" and m != "" and pd.notna(m)]

                        # 添加去重统计
                        for member in unique_members:
                            member_data = group_data[group_data["报销成员"] == member]
                            member_amount = member_data["申请报销金额"].sum() if "申请报销金额" in member_data.columns else 0

                            stat_row = {}
                            for col in sheet_df.columns:
                                if col == "报销成员":
                                    stat_row[col] = member
                                elif col == "申请报销金额":
                                    stat_row[col] = member_amount
                                else:
                                    stat_row[col] = ""

                            sheet_df = pd.concat([sheet_df, pd.DataFrame([stat_row])], ignore_index=True)

                        # 添加总计行
                        total_stat_row = {}
                        for col in sheet_df.columns:
                            if col == "报销成员":
                                total_stat_row[col] = "总计"
                            elif col == "申请报销金额":
                                total_amount = sum([group_data[group_data["报销成员"] == m]["申请报销金额"].sum()
                                                    for m in unique_members if m in group_data["报销成员"].values])
                                total_stat_row[col] = total_amount
                            else:
                                total_stat_row[col] = ""

                        sheet_df = pd.concat([sheet_df, pd.DataFrame([total_stat_row])], ignore_index=True)

                # 确定工作表名称
                if split_column == "支付账户":
                    # 对“支付账户”列进行简写处理
                    raw_name = str(group_name)
                    sheet_name = self.simplify_account_name(raw_name)
                else:
                    sheet_name = str(group_name)

                # 限制长度和去除非法字符
                sheet_name = sheet_name[:31]
                invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, '_')

                # 创建拆分的工作表
                split_ws = new_wb.create_sheet(sheet_name)
                for r_idx, row in enumerate(dataframe_to_rows(sheet_df, index=False, header=True), 1):
                    for c_idx, value in enumerate(row, 1):
                        split_ws.cell(row=r_idx, column=c_idx, value=value)
                
                # 对拆分工作表中的合计行进行黄色填充（只填充第一个合计行的 C 列和 D 列）
                from openpyxl.styles import PatternFill
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                # 找到第一个合计行
                total_row_found = False
                for row_idx in range(1, split_ws.max_row + 1):
                    # 检查 C 列是否为"合计"
                    c_cell = split_ws.cell(row=row_idx, column=3)
                    if c_cell.value == "合计" and not total_row_found:
                        # 填充 C 列的"合计"和 D 列的金额
                        split_ws.cell(row=row_idx, column=3).fill = yellow_fill
                        split_ws.cell(row=row_idx, column=4).fill = yellow_fill
                        total_row_found = True
                        break  # 只处理第一个合计行

            # 保存工作簿
            new_wb.save(save_path)
            new_wb.close()

            messagebox.showinfo("成功",
                                f"文件已保存到: {save_path}\n"
                                f"包含以下工作表：\n"
                                f"- 原始数据：完全未处理的原始数据（保留所有格式）\n"
                                f"- 处理后数据：经过前三步处理的数据\n"
                                f"- 按 '{split_column}' 列拆分的 {len(grouped)} 个工作表（含统计信息）")

        except Exception as e:
            messagebox.showerror("错误", f"拆分操作失败: {str(e)}")

    def run(self):
        """运行应用程序"""
        self.root.mainloop()


class ColumnSelector:
    """列选择器类"""

    def __init__(self, parent, columns, callback):
        self.columns = columns
        self.callback = callback

        self.window = tk.Toplevel(parent)
        self.window.title("选择拆分列")
        self.window.geometry("500x400")
        self.window.transient(parent)
        self.window.grab_set()
        
        # 设置弹窗居中
        self.center_window(self.window, 500, 400)

        self.setup_ui()

    def center_window(self, window, width, height):
        """设置窗口居中"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        """设置选择器界面"""
        # 标题
        tk.Label(self.window, text="请选择要拆分的列",
                 font=("Arial", 14, "bold")).pack(pady=10)

        # 说明文字
        tk.Label(self.window, text="以下为处理后的列列表，请选择一列作为拆分依据：",
                 font=("Arial", 10)).pack(pady=5)

        # 列选择框架
        frame = tk.Frame(self.window)
        frame.pack(fill="both", expand=True, padx=20, pady=10)

        # 创建树形视图
        self.tree = ttk.Treeview(frame, columns=('Index', 'ColumnName'), show='headings', height=12)
        self.tree.heading('Index', text='序号')
        self.tree.heading('ColumnName', text='列名')

        # 添加滚动条
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 添加列数据
        for i, col in enumerate(self.columns, 1):
            self.tree.insert("", "end", values=(i, col))

        # 绑定双击事件
        self.tree.bind("<Double-1>", self.on_double_click)

        # 按钮框架
        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="确认选择", command=self.confirm_selection,
                  bg="#2196F3", fg="white", font=("Arial", 11)).pack(side="left", padx=5)
        tk.Button(btn_frame, text="取消", command=self.window.destroy,
                  font=("Arial", 11)).pack(side="left", padx=5)

    def on_double_click(self, event):
        """双击选择列"""
        self.confirm_selection()

    def confirm_selection(self):
        """确认选择"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择一列进行拆分")
            return

        selected_item = self.tree.item(selection[0])
        selected_col = selected_item['values'][1]  # 获取列名

        self.window.destroy()
        self.callback(selected_col)


# 运行程序
if __name__ == "__main__":
    app = ExcelProcessor()
    app.run()