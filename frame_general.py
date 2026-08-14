#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
通用功能模块 - 5个功能
使用CustomTkinter实现的现代化UI
1. 拆分_拆成多个Excel文件
2. 拆分_拆成多个工作表
3. N层级文件夹中只获取文件
4. 表格格式转换（统一转xlsx）
5. PDF转Excel工具
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import sys
import os
import traceback

# 确保能导入项目根目录的业务逻辑
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from general.cf_duoge_wenj import split_excel_by_column as split_to_files
from general.cf_biaoge_li import split_excel_by_column as split_sheets
from general.Ncengjiwenjj_fzzt import main_zfzwj
from general.biaogegesi_zh import convert_and_copy_files as convert_format, select_directory
from general.zh_pdf_Excel import PDFToExcelConverter
from ui_components import FunctionCard, Theme


class GeneralFrame(ctk.CTkFrame):
    """通用功能面板"""

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
            text="📂 通用工具",
            font=Theme.FONT_HEADING,
            text_color=Theme.COLOR_ACCENT_GENERAL
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text="Excel 文件拆分与批量获取工具",
            font=Theme.FONT_BODY,
            text_color=Theme.COLOR_TEXT_SECONDARY
        ).pack(anchor="w", pady=(Theme.SPACING_XS, 0))

        # ---- 卡片网格容器 ----
        self._cards_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        self._cards_frame.pack(fill="both", expand=True,
                               padx=Theme.SPACING_XL, pady=(0, Theme.SPACING_LG))

        # ---- 功能卡片 ----
        self._add_card("1、拆分_拆成多个Excel文件",
                       "按列值将一个Excel拆分成多个独立文件",
                       "📄", Theme.COLOR_ACCENT_GENERAL, self.run_split_to_files)

        self._add_card("2、拆分_拆成多个工作表",
                       "按列值将数据拆分到同一文件的多个Sheet",
                       "📊", Theme.COLOR_ACCENT_GENERAL, self.run_split_to_sheets)

        self._add_card("3、N层级文件夹中只获取文件",
                       "从深层嵌套目录结构中提取所有文件",
                       "📁", Theme.COLOR_ACCENT_GENERAL, self.run_split_to_zhqwj)

        self._add_card("4、表格格式转换",
                       "将文件夹内所有Excel/CSV统一转换为xlsx格式",
                       "🔄", Theme.COLOR_ACCENT_TEAL, self.run_convert_format)

        self._add_card("5、PDF转Excel",
                       "将PDF表格提取并转换为Excel文件（对齐列版）",
                       "📄", Theme.COLOR_ACCENT_TEAL, self.open_pdf_to_excel)

        # 初始布局 + resize 绑定
        self._layout_cards()
        self._cards_frame.bind("<Configure>", self._on_resize)

    # ============================================================
    # 卡片管理
    # ============================================================

    def _add_card(self, title, desc, icon, color, command):
        """创建并登记一张功能卡片"""
        card = FunctionCard(
            self._cards_frame,
            title=title, description=desc, icon=icon,
            color=color, command=command
        )
        self._cards.append(card)

    def _layout_cards(self):
        """响应式网格：宽屏 2 列，窄屏 1 列"""
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

    def run_split_to_files(self):
        try:
            self.status_var.set("正在执行：拆分为多个Excel文件")
            parent = self.master.winfo_toplevel()
            split_to_files(parent)
            self.status_var.set("完成：拆分为多个Excel文件")
        except Exception as e:
            messagebox.showerror("错误", f"拆分为多个文件时出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：拆分为多个文件失败")

    def run_split_to_sheets(self):
        try:
            self.status_var.set("正在执行：拆分为多个工作表")
            parent = self.master.winfo_toplevel()
            split_sheets(parent)
            self.status_var.set("完成：拆分为多个工作表")
        except Exception as e:
            messagebox.showerror("错误", f"拆分为多个工作表时出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：拆分为多个工作表失败")

    def run_split_to_zhqwj(self):
        try:
            self.status_var.set("正在执行：N层级文件夹获取文件")
            parent = self.master.winfo_toplevel()
            main_zfzwj(parent)
            self.status_var.set("完成：N层级文件夹获取文件")
        except Exception as e:
            messagebox.showerror("错误", f"N层级文件夹操作出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：N层级文件夹操作失败")

    def run_convert_format(self):
        """表格格式转换"""
        try:
            self.status_var.set("请选择源目录")
            parent = self.master.winfo_toplevel()
            source_dir = select_directory("选择要读取的文件夹（源）", parent)
            if not source_dir:
                return
            target_dir = select_directory("选择要保存的文件夹（目标）", parent)
            if not target_dir:
                return
            self.status_var.set(f"正在转换: {source_dir} → {target_dir}")
            convert_format(source_dir, target_dir)
            self.status_var.set("完成：表格格式转换")
            messagebox.showinfo("完成", "文件转换和复制已完成!")
        except Exception as e:
            messagebox.showerror("错误", f"表格格式转换出错: \n{str(e)}")
            traceback.print_exc()
            self.status_var.set("错误：表格格式转换失败")

    def open_pdf_to_excel(self):
        """PDF转Excel"""
        try:
            parent = self.master.winfo_toplevel()
            app = PDFToExcelConverter(parent)
            self.status_var.set("已打开：PDF转Excel工具")
        except Exception as e:
            messagebox.showerror("错误", f"打开PDF转Excel工具失败: \n{str(e)}")
            traceback.print_exc()
