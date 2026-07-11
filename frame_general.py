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


class GeneralFrame(ctk.CTkFrame):
    """通用功能面板"""

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
            text="📂 通用工具",
            font=("Microsoft YaHei", 20, "bold"),
            text_color="#27ae60"
        )
        title_label.pack(anchor="w")

        desc_label = ctk.CTkLabel(
            title_frame,
            text="Excel 文件拆分与批量获取工具",
            font=("Microsoft YaHei", 12),
            text_color="#666666"
        )
        desc_label.pack(anchor="w", pady=(5, 0))

        # 功能卡片容器
        cards_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        cards_frame.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        # === 卡片1: 拆分成多个Excel文件 ===
        self._create_card(
            cards_frame,
            title="1、拆分_拆成多个Excel文件",
            desc="按列值将一个Excel拆分成多个独立文件",
            icon="📄",
            command=self.run_split_to_files,
            color="#27ae60"
        )

        # === 卡片2: 拆分成多个工作表 ===
        self._create_card(
            cards_frame,
            title="2、拆分_拆成多个工作表",
            desc="按列值将数据拆分到同一文件的多个Sheet",
            icon="📊",
            command=self.run_split_to_sheets,
            color="#27ae60"
        )

        # === 卡片3: N层级文件夹获取文件 ===
        self._create_card(
            cards_frame,
            title="3、N层级文件夹中只获取文件",
            desc="从深层嵌套目录结构中提取所有文件",
            icon="📁",
            command=self.run_split_to_zhqwj,
            color="#27ae60"
        )

        # === 卡片4: 表格格式转换 ===
        self._create_card(
            cards_frame,
            title="4、表格格式转换",
            desc="将文件夹内所有Excel/CSV统一转换为xlsx格式",
            icon="🔄",
            command=self.run_convert_format,
            color="#16a085"
        )

        # === 卡片5: PDF转Excel ===
        self._create_card(
            cards_frame,
            title="5、PDF转Excel",
            desc="将PDF表格提取并转换为Excel文件（对齐列版）",
            icon="📄",
            command=self.open_pdf_to_excel,
            color="#16a085"
        )

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
            hover_color="#1e8449" if color == "#27ae60" else "#117a65",
            text_color="white",
            corner_radius=8,
            width=100,
            height=36,
            command=command
        )
        btn.pack(side="right", padx=(15, 0))

        return card

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
