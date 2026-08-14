#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
人事功能模块 - 4个功能
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
from ui_components import FunctionCard, Theme


class HRFrame(ctk.CTkFrame):
    """人事功能面板"""

    def __init__(self, parent, status_var=None):
        super().__init__(parent, fg_color="white", corner_radius=0)
        self.status_var = status_var or tk.StringVar()
        self._cards = []
        self._build_ui()

    def _build_ui(self):
        """构建UI"""
        # ---- 标题区 ----
        title_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        title_frame.pack(fill="x", padx=Theme.SPACING_XL, pady=(Theme.SPACING_LG, Theme.SPACING_MD))

        ctk.CTkLabel(
            title_frame,
            text="👥 人事工具",
            font=Theme.FONT_HEADING,
            text_color=Theme.COLOR_ACCENT_HR
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text="在职离职拆分、指定内容参数拆分等人事数据处理",
            font=Theme.FONT_BODY,
            text_color=Theme.COLOR_TEXT_SECONDARY
        ).pack(anchor="w", pady=(Theme.SPACING_XS, 0))

        # ---- 卡片网格容器 ----
        self._cards_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        self._cards_frame.pack(fill="both", expand=True,
                               padx=Theme.SPACING_XL, pady=(0, Theme.SPACING_LG))

        # ---- 4个功能卡片 ----
        self._add_card("1、拆分_拆分两次_在职离职",
                       "按条件拆分在职/离职人员数据（吴苏霞）",
                       "✂️", Theme.COLOR_ACCENT_HR, self.rszy_cf_lc)

        self._add_card("2、拆分_指定内容放一起_拆分参数",
                       "将指定内容归集到一起并拆分参数（吴苏霞）",
                       "📌", Theme.COLOR_ACCENT_HR, self.rszy_cf_cfcs_wsx)

        self._add_card("3、工资表匹配",
                       "两张工资表数据对比与高亮标记差异",
                       "🔍", Theme.COLOR_ACCENT_HR, self.open_gzb_pipei)

        self._add_card("4、文档处理工具箱",
                       "工资表拆分（保留公式/样式）+ Word智能合并，统一界面操作",
                       "📦", Theme.COLOR_ACCENT_GENERAL, self.open_unified_tools)

        # 初始布局 + resize 绑定
        self._layout_cards()
        self._cards_frame.bind("<Configure>", self._on_resize)

    # ============================================================
    # 卡片管理
    # ============================================================

    def _add_card(self, title, desc, icon, color, command):
        card = FunctionCard(
            self._cards_frame,
            title=title, description=desc, icon=icon,
            color=color, command=command
        )
        self._cards.append(card)

    def _layout_cards(self):
        width = self._cards_frame.winfo_width()
        cols = 2 if width > Theme.CARD_GRID_BREAKPOINT else 1

        for c in range(cols):
            self._cards_frame.grid_columnconfigure(c, weight=1, uniform="card_col")

        for i, card in enumerate(self._cards):
            row, col = divmod(i, cols)
            card.grid(row=row, column=col,
                      padx=Theme.SPACING_XS, pady=Theme.SPACING_XS, sticky="ew")

    def _on_resize(self, event):
        if event.widget == self._cards_frame:
            self._layout_cards()

    # ============================================================
    # 业务方法
    # ============================================================

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
            app = UnifiedHRTools(parent=window)
            self.status_var.set("已打开：文档处理工具箱（工资表拆分 + Word合并）")
            window.lift()
            window.focus_force()
        except Exception as e:
            messagebox.showerror("错误", f"打开文档处理工具箱失败: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：文档处理工具箱打开失败")
