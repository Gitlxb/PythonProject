#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
财务功能模块 - 8个功能
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


class FinanceFrame(ctk.CTkFrame):
    """财务功能面板"""

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
            text="💰 财务工具",
            font=("Microsoft YaHei", 20, "bold"),
            text_color="#2980b9"
        )
        title_label.pack(anchor="w")

        desc_label = ctk.CTkLabel(
            title_frame,
            text="收支记账筛选、表格合并、手续费处理、绩效考核等",
            font=("Microsoft YaHei", 12),
            text_color="#666666"
        )
        desc_label.pack(anchor="w", pady=(5, 0))

        # 功能卡片容器
        cards_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        cards_frame.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        # === 8个功能卡片 ===
        self._create_card(cards_frame,
                          title="1、财务_收支记账_筛选",
                          desc="开关账收支记账筛选处理",
                          icon="🧾",
                          command=self.run_financial,
                          color="#2980b9")

        self._create_card(cards_frame,
                          title="2、供应商_合并表格",
                          desc="将供应商相关Excel表格进行合并",
                          icon="📋",
                          command=self.cwzy_gys_hbbg,
                          color="#2980b9")

        self._create_card(cards_frame,
                          title="3、出纳_工具箱",
                          desc="打开工具箱（独立窗口，含代工费/比较/预支/报销）",
                          icon="💵",
                          command=self.open_shouxufei_integrated,
                          color="#2980b9")

        self._create_card(cards_frame,
                          title="4、绩效考核",
                          desc="打开绩效考核工具（数据匹配/稳岗率/合并表格）",
                          icon="📈",
                          command=self.open_cw_jxkh,
                          color="#8e44ad")

        self._create_card(cards_frame,
                          title="5、预支平账",
                          desc="打开预支平账处理工具（独立窗口）",
                          icon="🔄",
                          command=self.open_cw_yzpz,
                          color="#8e44ad")

        self._create_card(cards_frame,
                          title="6、核算汇总",
                          desc="打开核算汇总处理工具（独立窗口）",
                          icon="📑",
                          command=self.open_hs_hz,
                          color="#8e44ad")

        self._create_card(cards_frame,
                          title="7、银行流水_摘要修改",
                          desc="处理银行流水，批量修改摘要内容（往来款/劳务费/工伤理赔等）",
                          icon="🏦",
                          command=self.run_yhls_process,
                          color="#2980b9")

        self._create_card(cards_frame,
                          title="8、批量批注工具",
                          desc="批量添加批注（B~K列→A列）或拆分批注（A列→右侧列）",
                          icon="📝",
                          command=self.open_pl_pizhu_tool,
                          color="#16a085")

    def _create_card(self, parent, title, desc, icon, command, color):
        """创建功能卡片"""
        card = ctk.CTkFrame(
            parent,
            fg_color="#fafafa",
            corner_radius=10,
            border_width=1,
            border_color="#eeeeee"
        )
        card.pack(fill="x", pady=6)

        inner = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        inner.pack(fill="x", padx=20, pady=12)

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
            hover_color="#1a5276" if color == "#2980b9" else "#6c3483" if color == "#8e44ad" else "#117a65",
            text_color="white",
            corner_radius=8,
            width=100,
            height=36,
            command=command
        )
        btn.pack(side="right", padx=(15, 0))

        return card

    # ---- 功能方法 ----

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
