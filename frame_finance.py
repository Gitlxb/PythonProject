#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
财务功能模块 - 9个功能
使用CustomTkinter实现的现代化UI
1. 财务_收支记账_筛选
2. 供应商_合并表格
3. 出纳_工具箱 (xzy/main_integrated 四合一工具，独立窗口)
4. 绩效考核 (zhf/main_app.py 导航版工具，独立窗口)
5. 预支平账 (Toplevel窗口)
6. 核算汇总 (Toplevel窗口)
7. 银行流水_摘要修改
8. 批量批注工具（添加批注/拆分批注，独立窗口）
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import sys
import os
import traceback

# 确保能导入项目根目录的业务逻辑
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from finance.cw_kgz_sx import select_excel_file
from finance.cw_gys_qzy import run_excel_merger_qzy
from finance.xzy.main_integrated import IntegratedApp
from finance.zhf.main_app import MainApp
from finance.pmh.yzpz_gui import ExcelProcessorGUI
from finance.cxm.hs_huizong_cxm_main import SummaryProcessorApp
from finance.cw_yhls_zyxg import process_excel as process_yhls
from finance.pl_cf_pizhu import CommentToolApp
from finance.hbgs_gui import HbgsApp
from ui_components import FunctionCard, Theme


class FinanceFrame(ctk.CTkFrame):
    """财务功能面板"""

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
            text="💰 财务工具",
            font=Theme.FONT_HEADING,
            text_color=Theme.COLOR_ACCENT_FINANCE
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text="收支记账筛选、表格合并、手续费处理、绩效考核等",
            font=Theme.FONT_BODY,
            text_color=Theme.COLOR_TEXT_SECONDARY
        ).pack(anchor="w", pady=(Theme.SPACING_XS, 0))

        # ---- 卡片网格容器 ----
        self._cards_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        self._cards_frame.pack(fill="both", expand=True,
                               padx=Theme.SPACING_XL, pady=(0, Theme.SPACING_LG))

        # ---- 9个功能卡片 ----
        self._add_card("1、财务_收支记账_筛选",
                       "开关账收支记账筛选处理",
                       "🧾", Theme.COLOR_ACCENT_FINANCE, self.run_financial)

        self._add_card("2、供应商_合并表格",
                       "将供应商相关Excel表格进行合并",
                       "📋", Theme.COLOR_ACCENT_FINANCE, self.cwzy_gys_hbbg)

        self._add_card("3、出纳_工具箱",
                       "打开工具箱（独立窗口，含代工费/比较/预支/报销）",
                       "💵", Theme.COLOR_ACCENT_FINANCE, self.open_shouxufei_integrated)

        self._add_card("4、绩效考核",
                       "打开绩效考核工具（数据匹配/稳岗率/合并表格）",
                       "📈", Theme.COLOR_ACCENT_PURPLE, self.open_cw_jxkh)

        self._add_card("5、预支平账",
                       "打开预支平账处理工具（独立窗口）",
                       "🔄", Theme.COLOR_ACCENT_PURPLE, self.open_cw_yzpz)

        self._add_card("6、核算汇总",
                       "打开核算汇总处理工具（独立窗口）",
                       "📑", Theme.COLOR_ACCENT_PURPLE, self.open_hs_hz)

        self._add_card("7、银行流水_摘要修改",
                       "处理银行流水，批量修改摘要内容（往来款/劳务费/工伤理赔等）",
                       "🏦", Theme.COLOR_ACCENT_FINANCE, self.run_yhls_process)

        self._add_card("8、批量批注工具",
                       "批量添加批注（B~K列→A列）或拆分批注（A列→右侧列）",
                       "📝", Theme.COLOR_ACCENT_TEAL, self.open_pl_pizhu_tool)

        self._add_card("9、合并个税",
                       "递归合并个税统计表并计算手续费",
                       "🧮", Theme.COLOR_ACCENT_TEAL, self.open_hbgs)

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

    def run_financial(self):
        """财务_收支记账_筛选"""
        try:
            self.status_var.set("正在执行：收支记账筛选")
            parent = self.master.winfo_toplevel()
            select_excel_file(parent)
            self.status_var.set("完成：收支记账筛选")
        except Exception as e:
            messagebox.showerror("错误", f"财务功能出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：财务功能失败")

    def cwzy_gys_hbbg(self):
        """供应商_合并表格"""
        try:
            self.status_var.set("正在执行：供应商合并表格")
            parent = self.master.winfo_toplevel()
            run_excel_merger_qzy(parent)
            self.status_var.set("完成：供应商合并表格")
        except Exception as e:
            messagebox.showerror("错误", f"供应商合并出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：供应商合并失败")

    def open_shouxufei_integrated(self):
        """出纳_手续费 - 打开xzy四合一工具"""
        try:
            parent = self.master.winfo_toplevel()
            app = IntegratedApp(parent)
            self.status_var.set("已打开：手续费处理工具（xzy四合一）")
        except Exception as e:
            messagebox.showerror("错误", f"打开手续费处理工具失败: \n{str(e)}")
            traceback.print_exc()

    def open_cw_jxkh(self):
        """绩效考核 - 打开zhf绩效考核工具"""
        try:
            parent = self.master.winfo_toplevel()
            window = ctk.CTkToplevel(parent)
            window.title("绩效考核处理工具")
            window.geometry("950x700")
            window.transient(parent)
            window.grab_set()
            app = MainApp(window)
            self.status_var.set("已打开：绩效考核工具（zhf）")
            window.lift()
            window.focus_force()
        except Exception as e:
            messagebox.showerror("错误", f"打开绩效考核工具失败: \n{str(e)}")
            traceback.print_exc()

    def open_cw_yzpz(self):
        """打开预支平账处理工具"""
        try:
            parent = self.master.winfo_toplevel()
            window = ctk.CTkToplevel(parent)
            window.title("预支平账处理工具")
            window.geometry("950x700")
            window.transient(parent)
            window.grab_set()
            app = ExcelProcessorGUI(window)
            self.status_var.set("已打开：预支平账工具")
            window.lift()
            window.focus_force()
        except Exception as e:
            messagebox.showerror("错误", f"打开预支平账工具失败: \n{str(e)}")
            traceback.print_exc()

    def open_hs_hz(self):
        """打开核算汇总处理工具"""
        try:
            parent = self.master.winfo_toplevel()
            window = ctk.CTkToplevel(parent)
            window.title("Excel汇总处理工具")
            window.geometry("1050x850")
            window.transient(parent)
            window.grab_set()
            app = SummaryProcessorApp(window)
            self.status_var.set("已打开：核算汇总工具")
            window.lift()
            window.focus_force()
        except Exception as e:
            messagebox.showerror("错误", f"打开核算工具失败: \n{str(e)}")
            traceback.print_exc()

    def run_yhls_process(self):
        """银行流水_摘要修改"""
        try:
            parent = self.master.winfo_toplevel()
            path = filedialog.askopenfilename(
                title="选择银行流水Excel文件",
                filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
                parent=parent
            )
            if not path:
                return
            self.status_var.set(f"正在处理银行流水: {os.path.basename(path)}")
            process_yhls(path)
            self.status_var.set("完成：银行流水摘要修改")
        except SystemExit:
            pass
        except Exception as e:
            messagebox.showerror("错误", f"银行流水处理出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：银行流水处理失败")

    def open_pl_pizhu_tool(self):
        """批量批注工具 - 打开合并版批注工具（添加/拆分）"""
        try:
            parent = self.master.winfo_toplevel()
            app = CommentToolApp(parent)
            self.status_var.set("已打开：批量批注工具")
        except Exception as e:
            messagebox.showerror("错误", f"打开批量批注工具失败: \n{str(e)}")
            traceback.print_exc()

    def open_hbgs(self):
        """合并个税 - 打开 hbgs 工具"""
        try:
            parent = self.master.winfo_toplevel()
            app = HbgsApp(parent)
            self.status_var.set("已打开：合并个税工具")
        except Exception as e:
            messagebox.showerror("错误", f"打开合并个税工具失败: \n{str(e)}")
            traceback.print_exc()
