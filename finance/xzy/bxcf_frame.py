# -- coding: utf-8 --
# @File : bxcf_frame.py
# @Description: 报销拆分工具 - Frame嵌入版
# @Note: 核心逻辑从 cw_bxcfyqh_xzy.py 提取并重构，UI适配为Frame嵌入模式

import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# 支付账户名称简写映射（与原始代码一致）
ACCOUNT_NAME_MAP = {
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


def simplify_account_name(name):
    """将支付账户的长名称简写为短名称"""
    return ACCOUNT_NAME_MAP.get(name, name)


class BxcfCore:
    """报销拆分核心逻辑"""

    def __init__(self):
        self.file_path = None
        self.original_wb = None
        self.sheet_name = None

    def load_file(self, file_path):
        self.file_path = file_path
        self.original_wb = load_workbook(file_path)
        self.sheet_name = self.original_wb.sheetnames[0]

    def process_excel(self):
        """
        执行前三步处理：拆分合并单元格、拆分第二行字段、删除第一行、移动申请报销金额列
        返回: 处理后的 DataFrame
        """
        wb = self.original_wb
        ws = wb[self.sheet_name]

        # 步骤1：拆分所有合并单元格
        merged_ranges = list(ws.merged_cells.ranges)
        for merged_range in merged_ranges:
            min_row, min_col, max_row, max_col = (
                merged_range.min_row, merged_range.min_col,
                merged_range.max_row, merged_range.max_col
            )
            top_left_value = ws.cell(row=min_row, column=min_col).value
            ws.unmerge_cells(str(merged_range))
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    ws.cell(row=row, column=col).value = top_left_value

        # 步骤2：拆分第二行的多字段单元格
        max_col = ws.max_column
        for col in range(max_col, 0, -1):
            cell = ws.cell(row=2, column=col)
            if cell.value and isinstance(cell.value, str) and ' ' in cell.value.strip():
                fields = [f.strip() for f in cell.value.split(' ') if f.strip()]
                if len(fields) > 1:
                    ws.insert_cols(col + 1, amount=len(fields) - 1)
                    ws.cell(row=2, column=col).value = fields[0]
                    for i, field in enumerate(fields[1:], start=1):
                        ws.cell(row=2, column=col + i).value = field

        # 步骤3：删除第一行
        ws.delete_rows(1)

        # 保存临时文件供 pandas 读取
        temp_file = "_temp_bxcf_processed.xlsx"
        wb.save(temp_file)
        wb.close()

        # 读取为 DataFrame
        processed_df = pd.read_excel(temp_file, sheet_name=self.sheet_name)

        # 清理临时文件
        if os.path.exists(temp_file):
            os.remove(temp_file)

        # 步骤4：移动"申请报销金额"列到C列右侧
        if "申请报销金额" in processed_df.columns:
            amount_col = processed_df["申请报销金额"]
            processed_df = processed_df.drop("申请报销金额", axis=1)
            if len(processed_df.columns) >= 3:
                processed_df.insert(3, "申请报销金额", amount_col)
            else:
                processed_df["申请报销金额"] = amount_col

        # 重新加载原始工作簿用于后续操作
        self.original_wb = load_workbook(self.file_path)
        self.sheet_name = self.original_wb.sheetnames[0]

        self.processed_df = processed_df
        return processed_df

    def execute_split(self, split_column, save_path):
        """
        按指定列执行拆分，保存到新文件
        """
        df = self.processed_df

        new_wb = Workbook()
        new_wb.remove(new_wb.active)

        # 1. 复制原始数据工作表（完全保留格式）
        original_ws = self.original_wb[self.sheet_name]
        original_copy = new_wb.create_sheet("原始数据")

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

        for col_letter, dim in original_ws.column_dimensions.items():
            if dim.width is not None:
                original_copy.column_dimensions[col_letter].width = dim.width
        for row_num, dim in original_ws.row_dimensions.items():
            if dim.height is not None:
                original_copy.row_dimensions[row_num].height = dim.height
        for merged_range in original_ws.merged_cells.ranges:
            original_copy.merge_cells(str(merged_range))

        # 2. 处理后数据工作表
        processed_ws = new_wb.create_sheet("处理后数据")
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                processed_ws.cell(row=r_idx, column=c_idx, value=value)

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        if "申请报销金额" in df.columns:
            total_amount = df["申请报销金额"].sum()
            total_row_idx = len(df) + 2
            processed_ws.cell(row=total_row_idx, column=3, value="合计")
            processed_ws.cell(row=total_row_idx, column=4, value=total_amount)
            processed_ws.cell(row=total_row_idx, column=3).fill = yellow_fill
            processed_ws.cell(row=total_row_idx, column=4).fill = yellow_fill

        # 3. 按选定列拆分
        grouped = df.groupby(split_column)

        for group_name, group_data in grouped:
            sheet_df = group_data.copy().reset_index(drop=True)

            # 添加合计行
            if "申请报销金额" in sheet_df.columns:
                sheet_df["申请报销金额"] = pd.to_numeric(sheet_df["申请报销金额"],
                                                           errors='coerce').fillna(0)

                total_row = {}
                for col in sheet_df.columns:
                    if col == "报销成员":
                        total_row[col] = "合计"
                    elif col == "申请报销金额":
                        total_row[col] = sheet_df[col].sum()
                    else:
                        total_row[col] = ""

                sheet_df = pd.concat([sheet_df, pd.DataFrame([total_row])], ignore_index=True)

                for _ in range(3):
                    empty_row = {col: "" for col in sheet_df.columns}
                    sheet_df = pd.concat([sheet_df, pd.DataFrame([empty_row])], ignore_index=True)

                # 报销成员去重统计
                if "报销成员" in sheet_df.columns:
                    unique_members = group_data["报销成员"].dropna().unique()
                    unique_members = [m for m in unique_members
                                      if m not in ("合计", "", None) and pd.notna(m)]

                    for member in unique_members:
                        member_data = group_data[group_data["报销成员"] == member]
                        member_amount = (member_data["申请报销金额"].sum()
                                         if "申请报销金额" in member_data.columns else 0)

                        stat_row = {}
                        for col in sheet_df.columns:
                            if col == "报销成员":
                                stat_row[col] = member
                            elif col == "申请报销金额":
                                stat_row[col] = member_amount
                            else:
                                stat_row[col] = ""
                        sheet_df = pd.concat([sheet_df, pd.DataFrame([stat_row])], ignore_index=True)

                    total_stat_row = {}
                    for col in sheet_df.columns:
                        if col == "报销成员":
                            total_stat_row[col] = "总计"
                        elif col == "申请报销金额":
                            total_stat_row[col] = sum([
                                group_data[group_data["报销成员"] == m]["申请报销金额"].sum()
                                for m in unique_members
                                if m in group_data["报销成员"].values
                            ])
                        else:
                            total_stat_row[col] = ""
                    sheet_df = pd.concat([sheet_df, pd.DataFrame([total_stat_row])], ignore_index=True)

            # 确定工作表名称
            if split_column == "支付账户":
                raw_name = str(group_name)
                sheet_name_out = simplify_account_name(raw_name)
            else:
                sheet_name_out = str(group_name)

            sheet_name_out = sheet_name_out[:31]
            for ch in ['\\', '/', '*', '?', ':', '[', ']']:
                sheet_name_out = sheet_name_out.replace(ch, '_')

            split_ws = new_wb.create_sheet(sheet_name_out)
            for r_idx, row in enumerate(dataframe_to_rows(sheet_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    split_ws.cell(row=r_idx, column=c_idx, value=value)

            # 合计行标黄
            for row_idx in range(1, split_ws.max_row + 1):
                c_cell = split_ws.cell(row=row_idx, column=3)
                if c_cell.value == "合计":
                    split_ws.cell(row=row_idx, column=3).fill = yellow_fill
                    split_ws.cell(row=row_idx, column=4).fill = yellow_fill
                    break

        new_wb.save(save_path)
        new_wb.close()


class ColumnSelectorDialog:
    """列选择器弹窗对话框"""

    def __init__(self, parent, columns, callback):
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("选择拆分列")
        self.window.geometry("500x400")
        self.window.transient(parent)
        self.window.grab_set()

        # 居中
        self.window.update_idletasks()
        w = self.window.winfo_width()
        h = self.window.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.window.geometry(f'+{x}+{y}')

        self._setup_ui(columns)

    def _setup_ui(self, columns):
        tk.Label(self.window, text="请选择要拆分的列",
                 font=("Arial", 14, "bold")).pack(pady=10)
        tk.Label(self.window,
                 text="以下为处理后的列列表，请选择一列作为拆分依据：",
                 font=("Arial", 10)).pack(pady=5)

        frame = tk.Frame(self.window)
        frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.tree = ttk.Treeview(frame, columns=('Index', 'ColumnName'),
                                 show='headings', height=12)
        self.tree.heading('Index', text='序号')
        self.tree.heading('ColumnName', text='列名')

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for i, col in enumerate(columns, 1):
            self.tree.insert("", "end", values=(i, col))

        self.tree.bind("<Double-1>", lambda e: self.confirm())

        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="确认选择", command=self.confirm,
                  bg="#2196F3", fg="white", font=("Arial", 11),
                  padx=15, pady=3).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=self.window.destroy,
                  font=("Arial", 11), padx=15, pady=3).pack(side=tk.LEFT, padx=5)

    def confirm(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择一列进行拆分")
            return
        item = self.tree.item(selection[0])
        selected_col = item['values'][1]
        self.window.destroy()
        self.callback(selected_col)


class BxcfFrame(ttk.Frame):
    """报销拆分功能 - 作为Frame嵌入主界面"""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.core = BxcfCore()
        self.processed_df = None
        self._build_ui()

    def _get_toplevel(self):
        return self.winfo_toplevel()

    def _build_ui(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        ttk.Label(
            main_frame, text="Excel报销拆分工具",
            font=("Microsoft YaHei", 16, "bold")
        ).pack(pady=(0, 12))

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=12)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        file_row = ttk.Frame(file_frame)
        file_row.pack(fill=tk.X, pady=5)
        ttk.Label(file_row, text="选择文件：", font=("Microsoft YaHei", 10), width=10).pack(side=tk.LEFT)
        self.lbl_file = ttk.Label(file_row, text="未选择文件", font=("Microsoft YaHei", 9),
                                  foreground="gray")
        self.lbl_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.btn_select_file = ttk.Button(file_row, text="选择", command=self.select_file, width=8)
        self.btn_select_file.pack(side=tk.RIGHT)

        # 操作按钮区域
        btn_area = ttk.Frame(main_frame)
        btn_area.pack(pady=8)

        self.btn_process = ttk.Button(
            btn_area, text="1. 处理Excel文件（前三步）",
            command=self.process_excel, width=30
        )
        self.btn_process.pack(pady=3)
        self.btn_process.config(state=tk.DISABLED)

        self.btn_split = ttk.Button(
            btn_area, text="2. 选择拆分列并执行",
            command=self.show_column_selector, width=30
        )
        self.btn_split.pack(pady=3)
        self.btn_split.config(state=tk.DISABLED)

        # 列预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="列预览界面（处理后）", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.preview_title_label = tk.Label(preview_frame, text="", font=("Microsoft YaHei", 11, "bold"))
        self.preview_title_label.pack(anchor="w")

        self.tree = ttk.Treeview(preview_frame, columns=('Index', 'ColumnName'), show='headings', height=10)
        self.tree.heading('Index', text='序号')
        self.tree.heading('ColumnName', text='列名')

        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")

    # ==================== 文件选择 ====================

    def select_file(self):
        filepath = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not filepath:
            return

        try:
            self.core.load_file(filepath)
            filename = os.path.basename(filepath)
            self.lbl_file.config(text=filename, foreground="black")
            self.btn_process.config(state=tk.NORMAL)
            self.preview_columns([], "原始文件列预览（请先执行步骤1处理）")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")

    # ==================== 预览 ====================

    def preview_columns(self, columns, title="列预览"):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.preview_title_label.config(text=title)
        for i, col in enumerate(columns, 1):
            self.tree.insert("", "end", values=(i, col))

    # ==================== 步骤1：处理Excel ====================

    def process_excel(self):
        try:
            self.core.load_file(self.core.file_path)  # 重新加载确保干净状态
            self.processed_df = self.core.process_excel()

            messagebox.showinfo(
                "处理完成",
                "前三步处理已完成：\n"
                "- 所有合并单元格已拆分并填充\n"
                "- 第二行中由空格分隔的字段已拆分到多列\n"
                "- 第一行已删除\n"
                "- 申请报销金额列已移动到C列右侧\n\n"
                "现在请点击'选择拆分列并执行'进行下一步。"
            )

            self.btn_split.config(state=tk.NORMAL)
            self.preview_columns(self.processed_df.columns.tolist(), "处理后列预览（请选择拆分列）")
            self.app.status_label.config(text="报销拆分 - 已处理，等待选择拆分列")

        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("错误", f"处理文件失败: {str(e)}")

    # ==================== 步骤2：拆分 ====================

    def show_column_selector(self):
        if self.processed_df is None:
            messagebox.showwarning("警告", "请先完成前三步处理")
            return

        dialog = ColumnSelectorDialog(
            self._get_toplevel(),
            self.processed_df.columns.tolist(),
            self.execute_split
        )
        self._get_toplevel().wait_window(dialog.window)

    def execute_split(self, split_column):
        save_path = filedialog.asksaveasfilename(
            title="保存拆分后的文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not save_path:
            return

        try:
            self.core.execute_split(split_column, save_path)

            grouped_count = len(self.processed_df.groupby(split_column))
            messagebox.showinfo(
                "成功",
                f"文件已保存到: {save_path}\n\n"
                f"包含以下工作表：\n"
                f"- 原始数据：完全未处理的原始数据（保留所有格式）\n"
                f"- 处理后数据：经过前三步处理的数据\n"
                f"- 按 '{split_column}' 列拆分的 {grouped_count} 个工作表（含统计信息）"
            )

            self.app.status_label.config(text=f"报销拆分 - 已保存 ({grouped_count} 个工作表)")

        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("错误", f"拆分操作失败: {str(e)}")
