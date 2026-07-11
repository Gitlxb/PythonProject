#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
稳岗率计算功能 Frame
分析项目驻场和项目经理的稳岗率，生成奖励汇总
基于 wgl_main_zhf.py 的业务逻辑封装的 UI 界面
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import pandas as pd

# 导入稳岗率业务逻辑函数（从各模块导入）
from .wgl_io import choose_excel_file, read_workbook, save_output
from .wgl_prepare import prepare_raw, extract_months
from .wgl_refresh import build_refresh
from .wgl_stability import (build_stability_sheet, build_stability_sheet_pm,
                              build_stability_sheet_combined, build_base_rate_sheet)
from .wgl_reward import build_reward_sheet


# ========== 对话框类（保持不变）==========

class SheetSelectionDialog(simpledialog.Dialog):
    """工作表选择对话框"""

    def __init__(self, parent, sheet_names):
        self.sheet_names = sheet_names
        super().__init__(parent, title="选择工作表")

    def body(self, master):
        ttk.Label(master, text="请选择要处理的工作表:").grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        self.listbox = tk.Listbox(master, width=50, height=10, font=('微软雅黑', 10))
        self.listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        for sheet in self.sheet_names:
            self.listbox.insert(tk.END, sheet)

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

        self.start_combo = ttk.Combobox(master, values=self.months_list, state="readonly", width=15)
        self.end_combo = ttk.Combobox(master, values=self.months_list, state="readonly", width=15)

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


# ========== 稳岗率功能Frame ==========

class StabilityFrame(tk.Frame):
    """稳岗率计算功能 - Frame版本"""

    def __init__(self, parent, status_var=None):
        super().__init__(parent)
        self.parent = parent
        self.status_var = status_var

        # 变量
        self.input_file = None
        self.output_file = None
        self.raw_data = None
        self.months_list = []
        self.target_month = None
        self.start_month = None
        self.end_month = None
        self.selected_sheet = None
        self.copy_original_sheets = tk.BooleanVar(value=False)
        self.outputs = {}

        # 构建界面
        self._build_ui()

    def _create_tooltip(self, widget, text):
        """为控件创建悬浮提示"""
        tip = None
        def show(event=None):
            nonlocal tip
            x = widget.winfo_pointerx() + 15
            y = widget.winfo_pointery() + 15
            tip = tk.Toplevel(widget)
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{x}+{y}")
            tk.Label(
                tip, text=text, bg="#333333", fg="white",
                font=("Microsoft YaHei UI", 9), padx=10, pady=6
            ).pack()
        def hide(event=None):
            nonlocal tip
            if tip:
                tip.destroy()
                tip = None
        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

    def _bind_hover(self, btn, tooltip_text, enabled_fg="#1565C0",
                    hover_fg="#0D47A1", disabled_fg="#9E9E9E"):
        """为按钮绑定鼠标悬浮效果（颜色变化 + 手型光标 + 提示）"""
        self._create_tooltip(btn, tooltip_text)

        def on_enter(event):
            if btn['state'] == 'normal':
                btn.config(fg=hover_fg, cursor="hand2")
            else:
                btn.config(cursor="")

        def on_leave(event):
            if btn['state'] == 'normal':
                btn.config(fg=enabled_fg, cursor="")
            else:
                btn.config(fg=disabled_fg)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    def _build_ui(self):
        """构建界面组件"""
        main = ttk.Frame(self, padding=16)
        main.pack(fill="both", expand=True)

        # 标题
        title_frame = ttk.Frame(main)
        title_frame.pack(fill=tk.X, pady=(0, 18))

        tk.Label(
            title_frame,
            text="📈 稳岗率计算 - 分析工具",
            font=("Microsoft YaHei UI", 16, "bold"),
            fg="#1565C0"
        ).pack(side=tk.LEFT)

        # 按钮区域
        btn_frame = ttk.LabelFrame(main, text="操作步骤", padding=14)
        btn_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        # 按钮1：选择Excel文件（始终启用）
        self.btn_select_file = tk.Button(
            btn_frame,
            text="选择 Excel 文件",
            command=self.select_file,
            font=('微软雅黑', 11),
            fg="#1565C0",
            pady=10,
            width=35,
            takefocus=True
        )
        self.btn_select_file.pack(pady=12)
        self._bind_hover(
            self.btn_select_file,
            "第一步：点击选择要分析的 Excel 数据文件"
        )

        # 按钮2：数据处理（初始禁用）
        self.btn_process = tk.Button(
            btn_frame,
            text="数据处理",
            command=self.process_data,
            font=('微软雅黑', 11),
            fg="#9E9E9E",
            width=35,
            pady=10,
            state=tk.DISABLED,
            takefocus=True
        )
        self.btn_process.pack(pady=12)
        self._bind_hover(
            self.btn_process,
            "第二步：处理数据（需先选择文件）"
        )

        # 按钮3：选择保存位置（初始禁用）
        self.btn_save = tk.Button(
            btn_frame,
            text="选择保存位置",
            command=self.save_file,
            font=('微软雅黑', 11),
            fg="#9E9E9E",
            width=35,
            pady=10,
            state=tk.DISABLED,
            takefocus=True
        )
        self.btn_save.pack(pady=12)
        self._bind_hover(
            self.btn_save,
            "第三步：选择保存位置（需先处理数据）"
        )

        # 选项：是否复制原始工作表
        option_frame = ttk.Frame(btn_frame)
        option_frame.pack(fill=tk.X, pady=8)

        self.chk_copy_sheets = tk.Checkbutton(
            option_frame,
            text="复制原始工作表（数据量大时建议取消勾选，可显著提升速度）",
            variable=self.copy_original_sheets,
            onvalue=True,
            offvalue=False,
            font=('微软雅黑', 9)
        )
        self.chk_copy_sheets.pack(anchor=tk.W)

        # 重置按钮和打开目录按钮
        reset_frame = ttk.Frame(btn_frame)
        reset_frame.pack(fill=tk.X, pady=(10, 0))

        self.btn_reset = tk.Button(
            reset_frame,
            text="🔄 重置",
            command=self.reset_state,
            font=('微软雅黑', 9),
            fg="#1565C0",
            width=12
        )
        self.btn_reset.pack(side=tk.LEFT)
        self._bind_hover(
            self.btn_reset,
            "点击重置所有状态，重新开始分析"
        )

        # 打开文件输出目录按钮（初始禁用）
        self.btn_open_folder = tk.Button(
            reset_frame,
            text="📁 打开输出目录",
            command=self.open_output_folder,
            font=('微软雅黑', 9),
            fg="#9E9E9E",
            width=16,
            state=tk.DISABLED
        )
        self.btn_open_folder.pack(side=tk.LEFT, padx=(10, 0))
        self._bind_hover(
            self.btn_open_folder,
            "打开文件保存所在的文件夹（需先保存文件）"
        )

        # 结果说明区
        info_frame = ttk.LabelFrame(main, text="处理结果说明", padding=12)
        info_frame.pack(fill=tk.X, pady=(0, 6))

        info_text = (
            "程序将自动生成以下工作表：\n"
            "• 项目驻场\n"
            "• 项目经理\n"
            "• 稳岗率汇总表\n"
            "• 驻场&项目经理 - 奖励汇总\n"
            "• 基准数规则表"
        )
        ttk.Label(info_frame, text=info_text, justify="left", foreground='#555555').pack(anchor="w")

    def select_file(self):
        """选择Excel文件"""
        try:
            top_window = self.winfo_toplevel()
            self.input_file = choose_excel_file()

            if self.input_file is None:
                if self.status_var:
                    self.status_var.set("已取消选择")
                return

            if self.status_var:
                self.status_var.set(f"已选择: {self.input_file.name}")

            # 读取Excel文件的所有工作表
            excel_file = pd.ExcelFile(self.input_file)
            sheet_names = excel_file.sheet_names
            excel_file.close()

            # 显示工作表选择对话框
            dialog = SheetSelectionDialog(top_window, sheet_names)
            if dialog.result is None:
                if self.status_var:
                    self.status_var.set("未选择工作表")
                return

            self.selected_sheet = dialog.result

            if self.status_var:
                self.status_var.set(f"已选文件: {self.input_file.name} | 工作表: {self.selected_sheet}")

            self.btn_select_file.config(text=f"✅ 已选择文件: {self.selected_sheet}")

            # 启用数据处理按钮（恢复蓝色文字）
            self.btn_process.config(state=tk.NORMAL, fg="#1565C0")

            # 读取数据获取月份列表
            self.raw_data = read_workbook(self.input_file, self.selected_sheet)
            self.raw_data = prepare_raw(self.raw_data)
            self.months_list = extract_months(self.raw_data)

            if not self.months_list:
                messagebox.showwarning("警告", "未找到有效的月份数据！", parent=top_window)
                self.btn_process.config(state=tk.DISABLED, fg="#9E9E9E")

        except Exception as e:
            messagebox.showerror("错误", f"选择文件失败：{str(e)}", parent=self.winfo_toplevel())
            if self.status_var:
                self.status_var.set("选择文件失败")

    def process_data(self):
        """处理数据"""
        if not self.months_list:
            messagebox.showerror("错误", "没有可用的月份数据！", parent=self.winfo_toplevel())
            return

        # 显示日期范围选择对话框
        dialog = DateRangeDialog(self.winfo_toplevel(), self.months_list)
        if dialog.result is None:
            return

        self.start_month = dialog.result['start_month']
        self.end_month = dialog.result['end_month']
        self.target_month = self.end_month

        try:
            # 更新状态
            if self.status_var:
                self.status_var.set(f"正在处理数据 ({self.start_month} ~ {self.target_month})...")
            self.update_idletasks()

            # 执行数据处理
            months_filtered = [m for m in self.months_list
                               if m >= self.start_month and m <= self.end_month]

            zc_refresh = build_refresh(self.raw_data, months_filtered, "项目驻场", self.start_month, self.end_month)
            pm_refresh = build_refresh(self.raw_data, months_filtered, "项目经理", self.start_month, self.end_month)
            stability_zc = build_stability_sheet(self.raw_data, self.target_month, zc_refresh)
            stability_pm = build_stability_sheet_pm(self.raw_data, self.target_month, pm_refresh)
            stability_combined = build_stability_sheet_combined(stability_zc.copy(), stability_pm.copy())
            reward = build_reward_sheet(stability_zc, stability_pm)

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

            # 显示完成对话框
            msg = (
                f"✅ 数据处理完成！\n\n"
                f"📌 核算月份：{self.target_month}\n"
                f"📅 时间范围：{self.start_month} 至 {self.end_month}\n"
                f"📊 项目驻场基准数：{zc_base_rate:.2%}\n"
                f"📊 项目经理基准数：{pm_base_rate:.2%}"
            )

            messagebox.showinfo("处理完成", msg, parent=self.winfo_toplevel())

            if self.status_var:
                self.status_var.set(f"数据已处理完毕，请选择保存位置")

            self.btn_process.config(text=f"✅ 数据已处理 ({self.target_month})")
            self.btn_save.config(state=tk.NORMAL, fg="#1565C0")

        except Exception as e:
            messagebox.showerror("错误", f"数据处理失败：{str(e)}", parent=self.winfo_toplevel())
            if self.status_var:
                self.status_var.set("数据处理失败")

    def save_file(self):
        """保存文件"""
        try:
            self.output_file = save_output(
                self.input_file,
                self.outputs,
                copy_original_sheets=self.copy_original_sheets.get()
            )

            if self.status_var:
                self.status_var.set(f"已保存: {self.output_file.name}")

            if self.copy_original_sheets.get():
                msg = f"✅ 文件已保存至：\n{self.output_file}\n\n已包含所有原始工作表。"
            else:
                msg = f"✅ 文件已保存至：\n{self.output_file}\n\n仅包含处理结果的工作表。"

            messagebox.showinfo("保存成功", msg, parent=self.winfo_toplevel())

            # 更新按钮状态
            self.btn_save.config(text="✅ 已保存")
            
            # 启用"打开输出目录"按钮
            self.btn_open_folder.config(state=tk.NORMAL, fg="#1565C0")

        except SystemExit:
            # 用户取消了保存
            pass
        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败：{str(e)}", parent=self.winfo_toplevel())
            if self.status_var:
                self.status_var.set("保存文件失败")

    def reset_state(self):
        """重置状态，可以重新处理"""
        self.input_file = None
        self.output_file = None
        self.raw_data = None
        self.months_list = []
        self.target_month = None
        self.start_month = None
        self.end_month = None
        self.selected_sheet = None
        self.outputs = {}

        self.btn_select_file.config(text="选择 Excel 文件", fg="#1565C0")
        self.btn_process.config(text="数据处理", state=tk.DISABLED, fg="#9E9E9E")
        self.btn_save.config(text="选择保存位置", state=tk.DISABLED, fg="#9E9E9E")
        self.btn_open_folder.config(state=tk.DISABLED, fg="#9E9E9E")

        if self.status_var:
            self.status_var.set("已重置，可以重新开始")

    def open_output_folder(self):
        """打开文件输出目录"""
        import os
        import subprocess
        
        if not self.output_file or not self.output_file.exists():
            messagebox.showwarning("警告", "输出文件不存在！", parent=self.winfo_toplevel())
            return
        
        try:
            # 获取文件所在目录
            output_dir = str(self.output_file.parent)
            
            # Windows 上使用 startfile 或 explorer
            if os.name == 'nt':
                os.startfile(output_dir)
            else:
                # macOS/Linux
                subprocess.run(['open', output_dir] if os.name == 'posix' else ['xdg-open', output_dir])
            
            if self.status_var:
                self.status_var.set(f"已打开输出目录: {output_dir}")
                
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录：{str(e)}", parent=self.winfo_toplevel())
