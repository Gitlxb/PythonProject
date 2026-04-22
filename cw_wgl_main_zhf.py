#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
稳岗率分析工具 - GUI 启动界面
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from pathlib import Path
import sys
import pandas as pd

# 导入分析函数
from cw_wgl_zhf import (
    choose_excel_file,
    read_workbook,
    prepare_raw,
    extract_months,
    build_refresh,
    build_stability_sheet,
    build_stability_sheet_pm,
    build_stability_sheet_combined,
    build_reward_sheet,
    build_base_rate_sheet,
    save_output,
    round_half_up
)


class SheetSelectionDialog(simpledialog.Dialog):
    """工作表选择对话框"""

    def __init__(self, parent, sheet_names):
        self.sheet_names = sheet_names
        super().__init__(parent, title="选择工作表")

    def body(self, master):
        ttk.Label(master, text="请选择要处理的工作表:").grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        # 创建列表框
        self.listbox = tk.Listbox(master, width=50, height=10, font=('微软雅黑', 10))
        self.listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        # 填充工作表名称
        for sheet in self.sheet_names:
            self.listbox.insert(tk.END, sheet)

        # 默认选择第一个
        if self.sheet_names:
            self.listbox.selection_set(0)

        return self.listbox

    def validate(self):
        if not self.listbox.curselection():
            messagebox.showwarning("警告", "请选择一个工作表", parent=self.parent)
            return False
        return True

    def apply(self):
        idx = self.listbox.curselection()[0]
        self.result = self.sheet_names[idx]


class DateRangeDialog(simpledialog.Dialog):
    """日期范围选择对话框"""

    def __init__(self, parent, months_list):
        self.months_list = months_list
        super().__init__(parent, title="选择时间范围")

    def body(self, master):
        ttk.Label(master, text="起始月份:").grid(row=0, column=0, padx=10, pady=10)
        ttk.Label(master, text="结束月份:").grid(row=1, column=0, padx=10, pady=10)

        # 创建下拉列表
        self.start_combo = ttk.Combobox(master, values=self.months_list, state="readonly", width=15)
        self.end_combo = ttk.Combobox(master, values=self.months_list, state="readonly", width=15)

        # 默认选择第一个和最后一个
        if self.months_list:
            self.start_combo.set(self.months_list[0])
            self.end_combo.set(self.months_list[-1])

        self.start_combo.grid(row=0, column=1, padx=10, pady=10)
        self.end_combo.grid(row=1, column=1, padx=10, pady=10)

        return self.start_combo

    def validate(self):
        if not self.start_combo.get() or not self.end_combo.get():
            messagebox.showwarning("警告", "请选择起始和结束月份", parent=self.parent)
            return False

        if self.start_combo.get() > self.end_combo.get():
            messagebox.showwarning("警告", "起始月份不能晚于结束月份", parent=self.parent)
            return False

        return True

    def apply(self):
        self.result = {
            'start_month': self.start_combo.get(),
            'end_month': self.end_combo.get()
        }


class ExcelAnalysisApp:
    """稳岗率分析工具 GUI 应用"""

    def __init__(self, root):
        self.root = root
        self.root.title("分析工具")
        self.root.geometry("600x650")  # 增加高度以容纳复选框

        # 变量
        self.input_file = None
        self.output_file = None
        self.raw_data = None
        self.months_list = []
        self.target_month = None
        self.start_month = None
        self.end_month = None
        self.copy_original_sheets = tk.BooleanVar(value=True)  # 默认复制原始工作表

        # 创建界面
        self.create_widgets()

    def create_widgets(self):
        """创建界面组件"""
        # 标题
        title_label = ttk.Label(
            self.root,
            text="分析工具",
            font=("微软雅黑", 22, "bold")
        )
        title_label.pack(pady=30)

        # 按钮框架
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.BOTH, expand=True, padx=80, pady=20)

        # 按钮样式配置
        button_style = {
            'width': 35,
            'takefocus': True
        }

        # 按钮 1：选择 Excel 文件
        self.btn_select_file = tk.Button(
            btn_frame,
            text="1. 选择 Excel 文件",
            command=self.select_file,
            font=('微软雅黑', 12),
            pady=10,
            **button_style
        )
        self.btn_select_file.pack(pady=15)

        # 按钮 2：数据处理
        self.btn_process = tk.Button(
            btn_frame,
            text="2. 数据处理",
            command=self.process_data,
            font=('微软雅黑', 12),
            width=35,
            pady=10,
            state=tk.DISABLED,
            takefocus=True
        )
        self.btn_process.pack(pady=15)

        # 按钮 3：选择保存位置
        self.btn_save = tk.Button(
            btn_frame,
            text="3. 选择保存位置",
            command=self.save_file,
            font=('微软雅黑', 12),
            width=35,
            pady=10,
            state=tk.DISABLED,
            takefocus=True
        )
        self.btn_save.pack(pady=15)

        # 添加复选框：是否复制原始工作表
        option_frame = ttk.Frame(btn_frame)
        option_frame.pack(fill=tk.X, pady=10)

        self.chk_copy_sheets = tk.Checkbutton(
            option_frame,
            text="复制原始工作表（数据量大时建议取消勾选，可显著提升速度）",
            variable=self.copy_original_sheets,
            onvalue=True,
            offvalue=False,
            font=('微软雅黑', 10)
        )
        self.chk_copy_sheets.pack(anchor=tk.W)

        # 按钮 4：退出
        self.btn_exit = tk.Button(
            btn_frame,
            text="退出",
            command=self.exit_app,
            font=('微软雅黑', 12),
            width=35,
            pady=10,
            takefocus=True
        )
        self.btn_exit.pack(pady=15)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(10, 5),
            font=('微软雅黑', 10)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 5))

    def select_file(self):
        """选择 Excel 文件"""
        try:
            self.input_file = choose_excel_file()
                
            # 如果用户取消选择,返回初始界面
            if self.input_file is None:
                self.status_var.set("已取消选择")
                return
                
            self.status_var.set(f"已选择:{self.input_file.name}")

            # 读取 Excel 文件的所有工作表
            excel_file = pd.ExcelFile(self.input_file)
            sheet_names = excel_file.sheet_names
            excel_file.close()

            # 显示工作表选择对话框
            dialog = SheetSelectionDialog(self.root, sheet_names)
            if dialog.result is None:
                self.status_var.set("未选择工作表")
                return

            self.selected_sheet = dialog.result
            self.status_var.set(f"已选择：{self.input_file.name} - 工作表：{self.selected_sheet}")
            self.btn_select_file.config(text=f"✓ 已选择文件")
            self.btn_process.config(state=tk.NORMAL)

            # 读取数据获取月份列表
            self.raw_data = read_workbook(self.input_file, self.selected_sheet)
            self.raw_data = prepare_raw(self.raw_data)
            self.months_list = extract_months(self.raw_data)

            if not self.months_list:
                messagebox.showwarning("警告", "未找到有效的月份数据！")
                self.btn_process.config(state=tk.DISABLED)

        except Exception as e:
            messagebox.showerror("错误", f"选择文件失败：{str(e)}")
            self.status_var.set("选择文件失败")

    def process_data(self):
        """处理数据"""
        if not self.months_list:
            messagebox.showerror("错误", "没有可用的月份数据！")
            return

        # 显示日期范围选择对话框
        dialog = DateRangeDialog(self.root, self.months_list)
        if dialog.result is None:
            return

        self.start_month = dialog.result['start_month']
        self.end_month = dialog.result['end_month']
        self.target_month = self.end_month

        try:
            # 更新状态
            self.status_var.set(f"正在处理数据...")
            self.root.update()

            # 执行数据处理
            months_filtered = [m for m in self.months_list
                               if m >= self.start_month and m <= self.end_month]

            zc_refresh = build_refresh(self.raw_data, months_filtered, "项目驻场", self.start_month, self.end_month)
            pm_refresh = build_refresh(self.raw_data, months_filtered, "项目经理", self.start_month, self.end_month)
            stability_zc = build_stability_sheet(self.raw_data, self.target_month, zc_refresh)
            stability_pm = build_stability_sheet_pm(self.raw_data, self.target_month, pm_refresh)
            stability_combined = build_stability_sheet_combined(stability_zc.copy(), stability_pm.copy())
            reward = build_reward_sheet(stability_zc, stability_pm, self.raw_data)

            # 提取基准数
            zc_avg_row = stability_zc[stability_zc["项目驻场"] == "平均数"]
            zc_median_row = stability_zc[stability_zc["项目驻场"] == "中位数"]
            zc_avg_rate = zc_avg_row["稳岗率"].iloc[0] if not zc_avg_row.empty else 0.0
            zc_median_rate = zc_median_row["稳岗率"].iloc[0] if not zc_median_row.empty else 0.0
            zc_base_rate = max(zc_avg_rate, zc_median_rate)

            pm_avg_row = stability_pm[stability_pm["项目经理"] == "平均数"]
            pm_median_row = stability_pm[stability_pm["项目经理"] == "中位数"]
            pm_avg_rate = pm_avg_row["稳岗率"].iloc[0] if not pm_avg_row.empty else 0.0
            pm_median_rate = pm_median_row["稳岗率"].iloc[0] if not pm_median_row.empty else 0.0
            pm_base_rate = max(pm_avg_rate, pm_median_rate)

            base_rate_sheet = build_base_rate_sheet(zc_base_rate, pm_base_rate)

            # 存储处理结果
            self.outputs = {
                "项目驻场": zc_refresh,
                "项目经理": pm_refresh,
                "稳岗率": stability_combined,
                "驻场&项目经理 - 奖励汇总": reward,
                "基准数": base_rate_sheet,
            }

            # 显示处理完成对话框
            msg = f"数据处理完成！\n\n" \
                  f"核算月份：{self.target_month}\n" \
                  f"时间范围：{self.start_month} 至 {self.end_month}\n" \
                  f"项目驻场基准数：{zc_base_rate:.2f}\n" \
                  f"项目经理基准数：{pm_base_rate:.2f}"

            messagebox.showinfo("处理完成", msg)

            # 更新状态
            self.status_var.set(f"数据已处理完毕，请选择保存位置")
            self.btn_process.config(text=f"✓ 数据已处理 ({self.target_month})")
            self.btn_save.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("错误", f"数据处理失败：{str(e)}")
            self.status_var.set("数据处理失败")

    def save_file(self):
        """保存文件"""
        try:
            # 传递选中的工作表名称和复制选项
            self.output_file = save_output(
                self.input_file,
                self.outputs,
                self.selected_sheet,
                copy_original_sheets=self.copy_original_sheets.get()
            )
            self.status_var.set(f"已保存：{self.output_file.name}")

            # 显示保存成功对话框
            if self.copy_original_sheets.get():
                msg = f"文件已保存至：\n{self.output_file}\n\n已包含所有原始工作表。"
            else:
                msg = f"文件已保存至：\n{self.output_file}\n\n仅包含处理结果的工作表。"

            messagebox.showinfo("保存成功", msg)

            # 重置状态，可以重新处理
            self.btn_save.config(text="✓ 已保存")
            self.btn_process.config(text="2. 数据处理", state=tk.NORMAL)
            self.btn_select_file.config(text="1. 选择 Excel 文件")

        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败：{str(e)}")
            self.status_var.set("保存文件失败")

    def exit_app(self):
        """退出应用"""
        if messagebox.askyesno("确认退出", "确定要退出程序吗？"):
            self.root.quit()
            self.root.destroy()


def main():
    """主函数"""
    root = tk.Tk()


    app = ExcelAnalysisApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
