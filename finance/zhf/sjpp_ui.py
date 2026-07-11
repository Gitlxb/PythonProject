#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
数据匹配功能 Frame
从原始表自动生成完整版表格，包含公式链和数据关联
基于 sjpp_zhf.py 改造
"""

import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 导入原始业务逻辑函数
from .sjpp_zhf import generate_full_workbook


class DataMatchFrame(tk.Frame):
    """数据匹配功能 - Frame版本"""

    def __init__(self, parent, status_var=None):
        super().__init__(parent)
        self.parent = parent
        self.status_var = status_var

        # 变量
        self.source_var = tk.StringVar()
        self.prev_var = tk.StringVar()
        self.output_var = tk.StringVar(value="等待生成...")

        # 构建界面
        self._build_ui()

    def _build_ui(self):
        """构建界面组件"""
        main = ttk.Frame(self, padding=16)
        main.pack(fill="both", expand=True)

        # 标题
        title_frame = ttk.Frame(main)
        title_frame.pack(fill=tk.X, pady=(0, 18))

        tk.Label(
            title_frame,
            text="📊 数据匹配 - 自动生成完整版",
            font=("Microsoft YaHei UI", 16, "bold"),
            fg="#2C3E50"
        ).pack(side=tk.LEFT)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main, text="文件选择", padding=14)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        # 源文件
        src_frame = ttk.Frame(file_frame)
        src_frame.pack(fill=tk.X, pady=6)

        ttk.Label(src_frame, text="源文件（当月新表）:", width=20, anchor="w").pack(side=tk.LEFT)
        ttk.Entry(src_frame, textvariable=self.source_var, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(src_frame, text="浏览...", command=self._select_source).pack(side=tk.LEFT)

        # 参考文件
        prev_frame = ttk.Frame(file_frame)
        prev_frame.pack(fill=tk.X, pady=6)

        ttk.Label(prev_frame, text="参考文件（上月旧表）:", width=20, anchor="w").pack(side=tk.LEFT)
        ttk.Entry(prev_frame, textvariable=self.prev_var, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(prev_frame, text="浏览...", command=self._select_prev).pack(side=tk.LEFT)

        # 输出位置
        out_frame = ttk.Frame(file_frame)
        out_frame.pack(fill=tk.X, pady=6)

        ttk.Label(out_frame, text="输出位置:", width=20, anchor="w").pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_var, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(out_frame, text="浏览...", command=self._select_output).pack(side=tk.LEFT)

        # 操作按钮
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X, pady=(12, 0))

        self.generate_btn = ttk.Button(
            btn_frame, text="🚀 生成完整版", command=self._generate
        )
        self.generate_btn.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(
            btn_frame, text="🔄 重置", command=self._reset
        ).pack(side=tk.LEFT)

        # 状态说明
        info_frame = ttk.LabelFrame(main, text="处理说明", padding=12)
        info_frame.pack(fill=tk.X, pady=(0, 6))

        info_text = (
            "功能说明：\n"
            "1. 选择当月新表作为源文件\n"
            "2. 选择上月旧表作为参考文件\n"
            "3. 点击「生成完整版」自动匹配数据并生成新表\n"
            "4. 生成结果将保存到指定位置"
        )
        ttk.Label(info_frame, text=info_text, justify="left", foreground='#555555').pack(anchor="w")

    def _select_source(self):
        """选择源文件"""
        path = filedialog.askopenfilename(
            title="选择源文件（当月新表）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")],
            parent=self.winfo_toplevel()
        )
        if path:
            self.source_var.set(path)
            if self.status_var:
                self.status_var.set(f"已选择源文件: {os.path.basename(path)}")

    def _select_prev(self):
        """选择参考文件"""
        path = filedialog.askopenfilename(
            title="选择参考文件（上月旧表）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")],
            parent=self.winfo_toplevel()
        )
        if path:
            self.prev_var.set(path)
            if self.status_var:
                self.status_var.set(f"已选择参考文件: {os.path.basename(path)}")

    def _select_output(self):
        """选择输出位置"""
        path = filedialog.asksaveasfilename(
            title="保存完整版表格",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            parent=self.winfo_toplevel()
        )
        if path:
            self.output_var.set(path)
            if self.status_var:
                self.status_var.set(f"输出位置: {os.path.basename(path)}")

    def _generate(self):
        """生成完整版表格"""
        source = self.source_var.get()
        prev = self.prev_var.get()
        output = self.output_var.get()

        if not source:
            messagebox.showerror("错误", "请先选择源文件（当月新表）", parent=self.winfo_toplevel())
            return

        if not prev:
            messagebox.showerror("错误", "请先选择参考文件（上月旧表）", parent=self.winfo_toplevel())
            return

        if not output or output == "等待生成...":
            messagebox.showerror("错误", "请先选择输出位置", parent=self.winfo_toplevel())
            return

        try:
            if self.status_var:
                self.status_var.set("正在生成完整版表格...")

            self.update_idletasks()

            # 调用业务逻辑
            generate_full_workbook(source, prev, output)

            if self.status_var:
                self.status_var.set(f"生成成功: {os.path.basename(output)}")

            messagebox.showinfo("完成", f"完整版表格已生成：\n{output}", parent=self.winfo_toplevel())

        except Exception as e:
            error_msg = traceback.format_exc()
            print(error_msg)
            messagebox.showerror("错误", f"生成失败：\n{str(e)}", parent=self.winfo_toplevel())
            if self.status_var:
                self.status_var.set("生成失败")

    def _reset(self):
        """重置所有选择"""
        self.source_var.set("")
        self.prev_var.set("")
        self.output_var.set("等待生成...")

        if self.status_var:
            self.status_var.set("已重置，请重新选择文件")
