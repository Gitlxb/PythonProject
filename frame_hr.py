#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
人事功能模块 - 5个功能
使用CustomTkinter实现的现代化UI
1. 拆分_拆分两次_在职离职 (吴苏霞)
2. 拆分_指定内容放一起_拆分参数 (吴苏霞)
3. 工资表匹配工具
4. 文档处理工具箱（工资表拆分 + Word文档合并，统一界面）
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import sys
import os
import traceback

# 确保能导入项目根目录的业务逻辑
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from hr.cf_duoge_wenj_zz import split_excel_by_column_zcfgzb
from hr.cf_zdnrfyq_nrcs_wsx import main_cf_zdnr_ygbg
from hr.rs_gzb_pipei_main import SalaryMatchApp

# 文档处理统一工具箱（hr/hyp/ 子包）
from hr.hyp.unified_hr_tools import UnifiedHRTools


class HRFrame(ctk.CTkFrame):
    """人事功能面板"""

    def __init__(self, parent, status_var=None):
        super().__init__(parent, fg_color="white", corner_radius=0)
        self.status_var = status_var or tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        """构建UI"""
        # 标题区
        title_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        title_frame.pack(fill="x", padx=30, pady=(25, 15))

        title_label = ctk.CTkLabel(
            title_frame,
            text="👥 人事工具",
            font=("Microsoft YaHei", 20, "bold"),
            text_color="#e67e22"
        )
        title_label.pack(anchor="w")

        desc_label = ctk.CTkLabel(
            title_frame,
            text="在职离职拆分、指定内容参数拆分等人事数据处理",
            font=("Microsoft YaHei", 12),
            text_color="#666666"
        )
        desc_label.pack(anchor="w", pady=(5, 0))

        # 功能卡片容器
        cards_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        cards_frame.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        # === 3个功能卡片 ===
        self._create_card(cards_frame,
                          title="1、拆分_拆分两次_在职离职",
                          desc="按条件拆分在职/离职人员数据（吴苏霞）",
                          icon="✂️",
                          command=self.rszy_cf_lc,
                          color="#e67e22")

        self._create_card(cards_frame,
                          title="2、拆分_指定内容放一起_拆分参数",
                          desc="将指定内容归集到一起并拆分参数（吴苏霞）",
                          icon="📌",
                          command=self.rszy_cf_cfcs_wsx,
                          color="#e67e22")

        self._create_card(cards_frame,
                          title="3、工资表匹配",
                          desc="两张工资表数据对比与高亮标记差异",
                          icon="🔍",
                          command=self.open_gzb_pipei,
                          color="#e67e22")

        self._create_card(cards_frame,
                          title="4、文档处理工具箱",
                          desc="工资表拆分（保留公式/样式）+ Word智能合并，统一界面操作",
                          icon="📦",
                          command=self.open_unified_tools,
                          color="#27ae60")

    def _create_card(self, parent, title, desc, icon, command, color):
        """创建功能卡片"""
        card = ctk.CTkFrame(
            parent,
            fg_color="#fafafa",
            corner_radius=10,
            border_width=1,
            border_color="#eeeeee"
        )
        card.pack(fill="x", pady=8)

        inner = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        inner.pack(fill="x", padx=20, pady=15)

        # 左侧图标+文字
        left = ctk.CTkFrame(inner, fg_color="transparent", corner_radius=0)
        left.pack(side="left", fill="both", expand=True)

        title_text = tk.Label(
            left,
            text=f"{icon} {title}",
            font=("Microsoft YaHei", 14, "bold"),
            fg="#333333",
            bg="#fafafa",
            anchor="w",
            justify="left"
        )
        title_text.pack(anchor="w", fill="x")

        desc_text = tk.Label(
            left,
            text=desc,
            font=("Microsoft YaHei", 11),
            fg="#888888",
            bg="#fafafa",
            anchor="w",
            justify="left",
            wraplength=500
        )
        desc_text.pack(anchor="w", fill="x", pady=(5, 0))

        # 右侧按钮
        btn = ctk.CTkButton(
            inner,
            text="▶ 运行",
            font=("Microsoft YaHei", 12),
            fg_color=color,
            hover_color="#d35400",
            text_color="white",
            corner_radius=8,
            width=100,
            height=36,
            command=command
        )
        btn.pack(side="right", padx=(15, 0))

        return card

    # ---- 功能方法 ----

    def rszy_cf_lc(self):
        """拆分_拆分两次_在职离职"""
        try:
            self.status_var.set("正在执行：拆分两次-在职离职")
            parent = self.master.winfo_toplevel()
            split_excel_by_column_zcfgzb(parent)
            self.status_var.set("完成：拆分两次-在职离职")
        except Exception as e:
            messagebox.showerror("错误", f"拆分两次出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：拆分两次失败")

    def rszy_cf_cfcs_wsx(self):
        """拆分_指定内容放一起_拆分参数"""
        try:
            self.status_var.set("正在执行：指定内容放一起-拆分参数")
            parent = self.master.winfo_toplevel()
            main_cf_zdnr_ygbg(parent)
            self.status_var.set("完成：指定内容放一起-拆分参数")
        except Exception as e:
            messagebox.showerror("错误", f"指定内容拆分出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：指定内容拆分失败")

    def open_gzb_pipei(self):
        """工资表匹配"""
        try:
            parent = self.master.winfo_toplevel()
            window = ctk.CTkToplevel(parent)
            window.title("工资表匹配工具")
            window.geometry("800x650")
            window.transient(parent)
            window.grab_set()
            app = SalaryMatchApp(window)
            self.status_var.set("已打开：工资表匹配工具")
            window.lift()
            window.focus_force()
        except Exception as e:
            messagebox.showerror("错误", f"打开工资表匹配工具失败: \n{str(e)}")
            traceback.print_exc()

    def open_unified_tools(self):
        """文档处理工具箱（工资表拆分 + Word文档合并，统一界面）"""
        try:
            self.status_var.set("正在打开：文档处理工具箱")
            parent = self.master.winfo_toplevel()
            window = tk.Toplevel(parent)
            window.title("📦 人事文档处理工具箱")
            window.geometry("1200x800")
            window.minsize(900, 650)
            window.transient(parent)
            # 不 grab_set — 让用户可以同时操作主窗口
            app = UnifiedHRTools(parent=window)
            self.status_var.set("已打开：文档处理工具箱（工资表拆分 + Word合并）")
            window.lift()
            window.focus_force()
        except Exception as e:
            messagebox.showerror("错误", f"打开文档处理工具箱失败: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：文档处理工具箱打开失败")
