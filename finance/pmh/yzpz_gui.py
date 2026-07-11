"""
预支平账 GUI - 业务逻辑 + 主入口
继承 yzpz_gui_base.ExcelProcessorGUI，添加核对平账、保存、重置等逻辑
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from finance.pmh.yzpz_gui_base import ExcelProcessorGUI
from finance.pmh.yzpz_core import excel_processor
import os


class ExcelProcessorGUI(ExcelProcessorGUI):
    """预支平账处理工具 GUI（完整版，包含业务逻辑）"""
    pass


# ==================== 业务逻辑方法（动态绑定到类） ====================

def check_and_balance(self):
    """核对 - 平账：先核对，核对成功后立即平账"""
    # 同步输入框的值（支持直接输入）
    self.selected_unit = self.unit_var.get().strip()
    self.selected_month = self.month_var.get().strip()

    # 检查预支明细表是否已选择
    if not self.advance_file_path:
        messagebox.showwarning("警告", "请先选择预支明细表文件")
        return

    # 检查预支明细表工作簿是否已选择
    advance_sheet = self.advance_sheet_var.get()
    if not advance_sheet:
        messagebox.showwarning("警告", "请先选择预支明细表的工作簿")
        return

    # 检查工资表工作簿是否已选择
    if not self.check_workbook or not self.check_sheet_name:
        messagebox.showwarning("警告", "请先选择工资表文件和工作簿")
        return

    # 检查单位是否已选择
    if not self.selected_unit:
        messagebox.showwarning("警告", "请先在第三行选择单位")
        return

    # 检查平账时间是否已选择
    if not self.selected_month:
        messagebox.showwarning("警告", "请先在第三行选择平账时间")
        return

    self.log_status(f"已选择工资表工作簿：{self.check_sheet_name}")
    self.log_status(f"已选择单位：{self.selected_unit}")
    self.log_status(f"已选择平账时间：{self.selected_month}")

    # ==================== 第一步：核对 ====================
    current_round = excel_processor.get_current_check_round()
    names_num, data_num = excel_processor._get_round_sheet_numbers()

    self.log_status(f"\n===== 第 {current_round} 轮核对开始 =====")
    self.log_status(f"本轮将使用：Sheet{names_num}(姓名)、Sheet{data_num}(数据)")

    required_sheets = [f'Sheet{names_num}', f'Sheet{data_num}']

    missing_sheets = [s for s in required_sheets if s not in excel_processor.sheet_names]
    if missing_sheets:
        success, message = excel_processor._create_check_sheets()
        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", f"创建工作表失败：\n{message}")
            return
        excel_processor.sheet_names = excel_processor.workbook.sheetnames
        self.log_status(f"✓ 已创建缺失的工作表：{', '.join(missing_sheets)}")
    else:
        self.log_status(f"✓ 已确认 Sheet{names_num}、Sheet{data_num}")

    success, message = excel_processor.copy_names_from_workbook(
        self.check_workbook,
        self.check_sheet_name
    )

    if not success:
        self.log_status(f"✗ {message}")
        messagebox.showerror("错误", f"复制失败：\n{message}")
        return

    self.log_status(f"✓ 姓名已复制到 Sheet{names_num} 工作表中")

    merged_sheet = self.advance_sheet_var.get()
    if not merged_sheet:
        messagebox.showwarning("警告", "请先在预支明细表的工作簿下拉框中选择目标工作表")
        return

    success, message = excel_processor.apply_vlookup_formula(merged_sheet)
    if not success:
        self.log_status(f"✗ {message}")
        messagebox.showerror("错误", message)
        return

    # 单位已提前选好，直接传入
    if message.startswith("UNITS:"):
        success, message = excel_processor.apply_vlookup_formula(merged_sheet, [self.selected_unit])
        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", message)
            return

    self.log_status(f"✓ {message}")

    self.log_status(f"开始处理需扣回/预支扣款/预支扣回数据...")

    selected_column = self.select_deduction_column()

    success, message = excel_processor.process_deduction_data(
        self.check_workbook,
        self.check_sheet_name,
        selected_column
    )

    if not success:
        self.log_status(f"✗ {message}")
        messagebox.showerror("错误", f"处理需扣回/预支扣款/预支扣回数据失败：\n{message}")
        return

    self.log_status(f"✓ {message}")
    self.log_status(f"===== 第 {current_round} 轮核对完成 =====\n")

    # ==================== 第二步：平账 ====================
    salary_sheet_name = self.advance_sheet_var.get()
    month = self.selected_month

    self.log_status(f"\n===== 第 {current_round} 轮平账开始 =====")
    self.log_status(f"使用数据：Sheet{data_num}")
    self.log_status(f"平账月份：{month}")
    self.log_status(f"平账目标：{salary_sheet_name}")
    self.log_status(f"平账单位：{self.selected_unit}")

    success, message = excel_processor.balance_accounts(salary_sheet_name, month, [self.selected_unit])

    if success:
        self.log_status(f"✓ {message}")

        messagebox.showinfo("成功",
            f"第 {current_round} 轮核对 - 平账完成！\n{message}\n\n"
            f"如需继续，请切换工作簿选择新数据源，再点击【核对 - 平账】。")

        next_round = excel_processor.increment_round()
        self.log_status(f"✓ 轮次已递增：{current_round} -> {next_round}")
        self.log_status(f"===== 第 {current_round} 轮全部完成 =====\n")
    else:
        self.log_status(f"✗ {message}")
        messagebox.showerror("错误", message)


def select_deduction_column(self):
    """选择需扣回数据类型（当有多个列有数据时）"""
    available_columns = []

    name_col = None
    deduction_col = None
    advance_deduction_col = None
    advance_return_col = None

    for col in range(1, self.check_workbook.max_column + 1):
        cell_value = self.check_workbook.cell(row=2, column=col).value
        if cell_value:
            cell_str = str(cell_value).strip()
            if cell_str == "姓名":
                name_col = col
            elif "需扣回" in cell_str:
                deduction_col = col
            elif "预支扣款" in cell_str:
                advance_deduction_col = col
            elif "预支扣回" in cell_str:
                advance_return_col = col

    if deduction_col:
        for row in range(3, self.check_workbook.max_row + 1):
            val = self.check_workbook.cell(row=row, column=deduction_col).value
            if val is not None and str(val).strip() != '':
                available_columns.append('需扣回')
                break

    if advance_deduction_col:
        for row in range(3, self.check_workbook.max_row + 1):
            val = self.check_workbook.cell(row=row, column=advance_deduction_col).value
            if val is not None and str(val).strip() != '':
                available_columns.append('预支扣款')
                break

    if advance_return_col:
        for row in range(3, self.check_workbook.max_row + 1):
            val = self.check_workbook.cell(row=row, column=advance_return_col).value
            if val is not None and str(val).strip() != '':
                available_columns.append('预支扣回')
                break

    if len(available_columns) <= 1:
        return None

    dialog = tk.Toplevel(self.root)
    dialog.title("选择数据列")
    dialog.geometry("320x200")
    dialog.transient(self.root)
    dialog.grab_set()

    self.center_dialog(dialog, 320, 200)

    tk.Label(dialog, text="检测到多个数据列，请选择使用哪一列：").pack(pady=(16, 10))

    selected = [None]

    def on_select(col_type):
        selected[0] = col_type
        dialog.destroy()

    for col_type in available_columns:
        btn = ttk.Button(dialog, text=col_type, command=lambda c=col_type: on_select(c))
        btn.pack(pady=4, fill=tk.X, padx=30)

    self.root.wait_window(dialog)
    return selected[0]


def select_save_location(self):
    """选择保存位置并保存文件"""
    if not self.advance_file_path:
        messagebox.showwarning("警告", "请先选择预支明细表文件")
        return

    dialog = tk.Toplevel(self.root)
    dialog.title("选择保存方式")
    dialog.geometry("320x140")
    dialog.transient(self.root)
    dialog.grab_set()

    self.center_dialog(dialog, 320, 140)

    tk.Label(dialog, text="请选择保存方式：").pack(pady=(16, 12))

    save_path = [None]

    def save_to_original():
        save_path[0] = self.advance_file_path
        dialog.destroy()

    def save_as_new():
        path = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            save_path[0] = path
        dialog.destroy()

    def cancel_save():
        save_path[0] = None
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=8)

    ttk.Button(btn_frame, text="保存到原文件", command=save_to_original).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="另存为新文件", command=save_as_new).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="取消", command=cancel_save).pack(side=tk.LEFT, padx=5)

    self.root.wait_window(dialog)

    if not save_path[0]:
        return

    success, message = excel_processor.save_workbook(save_path[0])
    if success:
        self.last_save_path = save_path[0]
        self.log_status(f"✓ {message}")
        messagebox.showinfo("成功", f"文件已保存到:\n{save_path[0]}")
    else:
        self.log_status(f"✗ {message}")
        messagebox.showerror("错误", message)


def open_save_directory(self):
    """打开保存文件所在的目录"""
    target_path = self.last_save_path or self.advance_file_path
    if not target_path:
        messagebox.showwarning("警告", "尚未保存文件，无法打开目录")
        return

    directory = os.path.dirname(target_path)
    if os.path.exists(directory):
        os.startfile(directory)
        self.log_status(f"✓ 已打开目录：{directory}")
    else:
        messagebox.showerror("错误", "目录不存在")


def reset_all(self):
    """重置所有状态，允许换文件重新操作"""
    if messagebox.askyesno("确认", "确定要重置吗？当前未保存的数据将丢失。"):
        excel_processor.close_workbook()

        if self.salary_workbook:
            try:
                self.salary_workbook.close()
            except:
                pass

        self.advance_file_path = None
        self.sheet_names = []
        self.check_sheet_name = None
        self.check_workbook = None
        self.selected_unit = None
        self.selected_month = None
        self.available_units = []
        self.last_save_path = None

        self.advance_sheet_var.set("")
        self.advance_sheets = []
        self.advance_entry_var.set("尚未选择文件")

        self.salary_file_path = None
        self.salary_sheet_var.set("")
        self.salary_sheets = []
        self.salary_workbook = None
        self.salary_entry_var.set("尚未选择文件")

        self.unit_var.set("")
        self.month_var.set("")
        self.unit_combo.config(state="disabled")
        self.month_combo.config(state="disabled")

        self.log_status("✓ 已重置，可以开始新的操作")


# 将业务逻辑方法绑定到 ExcelProcessorGUI 类
ExcelProcessorGUI.check_and_balance = check_and_balance
ExcelProcessorGUI.select_deduction_column = select_deduction_column
ExcelProcessorGUI.select_save_location = select_save_location
ExcelProcessorGUI.open_save_directory = open_save_directory
ExcelProcessorGUI.reset_all = reset_all


# ==================== 主入口 ====================

def main():
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
