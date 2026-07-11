#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据去重工具 - GUI界面模块
使用 CustomTkinter 实现现代化GUI
功能：选择Excel文件 -> 选择Sheet -> 预览去重结果 -> 写入新Sheet
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

from sjqx_ghm import (
    get_sheet_names,
    get_headers,
    process_deduplication,
    write_deduplication_result,
)

# 设置CustomTkinter外观
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class DeduplicationApp:
    """数据去重工具主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("数据去重工具 - 按姓名+身份证去重保留最新日期")
        self.root.geometry("1000x750")
        self.root.minsize(900, 650)

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择Excel文件")

        self.headers = []  # 表头列表
        self.preview_data = []  # 预览数据
        self.total_count = 0  # 总数据行数
        self.dedup_count = 0  # 去重后行数

        self._create_ui()

    def _create_ui(self):
        """创建界面布局"""
        # 主容器
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ===== 标题 =====
        title_label = ctk.CTkLabel(
            main_container,
            text="📊 数据去重工具",
            font=("Microsoft YaHei", 24, "bold"),
            text_color="#2c3e50",
        )
        title_label.pack(pady=(0, 5))

        desc_label = ctk.CTkLabel(
            main_container,
            text="按「姓名+身份证」联合去重，保留「日期」最新的一整行数据",
            font=("Microsoft YaHei", 12),
            text_color="#7f8c8d",
        )
        desc_label.pack(pady=(0, 15))

        # ===== 文件选择区域 =====
        file_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(
            file_frame,
            text="Excel文件：",
            font=("Microsoft YaHei", 12),
            width=80,
        ).pack(side=tk.LEFT)

        self.file_entry = ctk.CTkEntry(
            file_frame,
            textvariable=self.file_path,
            state="readonly",
            font=("Microsoft YaHei", 11),
            width=500,
        )
        self.file_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        self.browse_btn = ctk.CTkButton(
            file_frame,
            text="浏览...",
            command=self._on_browse,
            font=("Microsoft YaHei", 12),
            width=80,
            height=32,
            corner_radius=6,
        )
        self.browse_btn.pack(side=tk.LEFT)

        # ===== Sheet选择区域 =====
        sheet_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        sheet_frame.pack(fill=tk.X, pady=(0, 15))

        ctk.CTkLabel(
            sheet_frame,
            text="选择Sheet：",
            font=("Microsoft YaHei", 12),
            width=80,
        ).pack(side=tk.LEFT)

        self.sheet_combo = ctk.CTkComboBox(
            sheet_frame,
            variable=self.sheet_name,
            state="readonly",
            font=("Microsoft YaHei", 11),
            width=300,
            height=32,
            corner_radius=6,
            values=[],
        )
        self.sheet_combo.pack(side=tk.LEFT, padx=(0, 10))

        # ===== 操作按钮区域 =====
        btn_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        btn_frame.pack(pady=(0, 15))

        self.process_btn = ctk.CTkButton(
            btn_frame,
            text="▶ 开始处理",
            command=self._on_process,
            font=("Microsoft YaHei", 14, "bold"),
            width=160,
            height=40,
            corner_radius=8,
            fg_color="#27ae60",
            hover_color="#219a52",
        )
        self.process_btn.pack(side=tk.LEFT, padx=10)

        self.export_btn = ctk.CTkButton(
            btn_frame,
            text="💾 写入新Sheet",
            command=self._on_export,
            font=("Microsoft YaHei", 14, "bold"),
            width=160,
            height=40,
            corner_radius=8,
            fg_color="#2980b9",
            hover_color="#2471a3",
            state="disabled",
        )
        self.export_btn.pack(side=tk.LEFT, padx=10)

        self.open_dir_btn = ctk.CTkButton(
            btn_frame,
            text="📂 打开输出目录",
            command=self._on_open_dir,
            font=("Microsoft YaHei", 14, "bold"),
            width=160,
            height=40,
            corner_radius=8,
            fg_color="#e67e22",
            hover_color="#d35400",
            state="disabled",
        )
        self.open_dir_btn.pack(side=tk.LEFT, padx=10)

        # ===== 统计信息区域 =====
        self.stats_frame = ctk.CTkFrame(main_container, fg_color="#ecf0f1", corner_radius=8)
        self.stats_frame.pack(fill=tk.X, pady=(0, 10))
        self.stats_frame.pack_forget()  # 初始隐藏

        self.stats_label = ctk.CTkLabel(
            self.stats_frame,
            text="",
            font=("Microsoft YaHei", 11),
            text_color="#2c3e50",
            justify="left",
        )
        self.stats_label.pack(padx=15, pady=8, anchor="w")

        # ===== 预览区域 =====
        preview_label = ctk.CTkLabel(
            main_container,
            text="📋 数据预览（去重结果）",
            font=("Microsoft YaHei", 14, "bold"),
            text_color="#2c3e50",
        )
        preview_label.pack(pady=(0, 5), anchor="w")

        # 预览表格容器（带滚动条）
        preview_container = ctk.CTkFrame(main_container, fg_color="white", corner_radius=8)
        preview_container.pack(fill=tk.BOTH, expand=True)

        # 创建Treeview（使用tkinter的Treeview，customtkinter没有表格组件）
        self.tree_frame = ctk.CTkFrame(preview_container, fg_color="white", corner_radius=0)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 垂直滚动条
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical")
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 水平滚动条
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal")
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview
        self.tree = ttk.Treeview(
            self.tree_frame,
            yscrollcommand=self.vsb.set,
            xscrollcommand=self.hsb.set,
            show="headings",
            height=15,
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # ===== 状态栏 =====
        status_bar = ctk.CTkFrame(self.root, height=28, fg_color="#2c3e50", corner_radius=0)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        status_bar.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            status_bar,
            textvariable=self.status_var,
            font=("Microsoft YaHei", 10),
            text_color="white",
            anchor="w",
        )
        self.status_label.pack(fill=tk.X, padx=15, pady=4)

    def _on_browse(self):
        """浏览并选择Excel文件"""
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.file_path.set(path)
            self._load_sheets()

    def _load_sheets(self):
        """加载Excel文件的Sheet列表"""
        path = self.file_path.get()
        if not path:
            return
        try:
            sheets = get_sheet_names(path)
            self.sheet_combo.configure(values=sheets)
            if sheets:
                self.sheet_combo.set(sheets[0])
            self.status_var.set(f"已加载文件：{os.path.basename(path)}，共 {len(sheets)} 个Sheet")
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件：\n{e}")
            self.status_var.set("文件读取失败")

    def _on_process(self):
        """开始处理按钮回调"""
        path = self.file_path.get()
        sheet = self.sheet_name.get()

        if not path:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return
        if not sheet:
            messagebox.showwarning("提示", "请选择要处理的工作表")
            return

        try:
            self.status_var.set("正在处理...")
            self.root.update_idletasks()

            # 调用核心逻辑
            self.headers, self.preview_data, self.total_count, self.dedup_count = (
                process_deduplication(path, sheet, header_row=1)
            )

            # 更新统计信息
            removed = self.total_count - self.dedup_count
            stats_text = (
                f"📊 统计信息："
                f"原始数据 {self.total_count} 行  →  "
                f"去重后 {self.dedup_count} 行  "
                f"（去除重复 {removed} 行）"
            )
            self.stats_label.configure(text=stats_text)
            self.stats_frame.pack(fill=tk.X, pady=(0, 10))

            # 显示预览
            self._show_preview()

            # 启用导出按钮
            self.export_btn.configure(state="normal")

            self.status_var.set(
                f"处理完成：{self.total_count} 行 → {self.dedup_count} 行（去除 {removed} 条重复）"
            )

        except ValueError as e:
            messagebox.showerror("处理错误", str(e))
            self.status_var.set("处理失败")
        except Exception as e:
            messagebox.showerror("处理错误", f"发生未知错误：\n{e}")
            self.status_var.set("处理失败")

    def _format_preview_value(self, val):
        """格式化预览单元格的值，将datetime转为友好格式"""
        from datetime import datetime
        if isinstance(val, datetime):
            # 格式化为 "M月D日" 中文格式
            return f"{val.month}月{val.day}日"
        return str(val) if val is not None else ""

    def _show_preview(self):
        """在Treeview中显示预览数据"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.headers or not self.preview_data:
            return

        # 设置列
        display_headers = self.headers[:20]  # 最多显示20列
        self.tree["columns"] = list(range(len(display_headers)))

        # 设置表头
        self.tree.heading("#0", text="序号")
        for idx, header in enumerate(display_headers):
            self.tree.heading(idx, text=header)
            self.tree.column(idx, width=100, minwidth=60, anchor="center")

        # 设置序号列
        self.tree.column("#0", width=50, minwidth=40, anchor="center")

        # 添加数据行
        for row_idx, row_data in enumerate(self.preview_data, start=1):
            display_values = []
            for col_idx in range(len(display_headers)):
                val = row_data[col_idx] if col_idx < len(row_data) else ""
                display_values.append(self._format_preview_value(val))
            self.tree.insert("", tk.END, text=str(row_idx), values=display_values)

    def _on_export(self):
        """写入新Sheet按钮回调"""
        path = self.file_path.get()
        sheet = self.sheet_name.get()

        if not path or not sheet:
            messagebox.showwarning("提示", "请先选择文件和Sheet")
            return

        try:
            self.status_var.set("正在写入新Sheet...")
            self.root.update_idletasks()

            new_sheet = write_deduplication_result(
                path, sheet, new_sheet_name="去重结果", header_row=1
            )

            self.status_var.set(f"写入成功！新Sheet名称：{new_sheet}")

            # 启用"打开输出目录"按钮
            self.open_dir_btn.configure(state="normal")

            messagebox.showinfo(
                "完成",
                f"去重结果已写入新Sheet：{new_sheet}\n\n"
                f"文件路径：{path}\n"
                f"去重后行数：{self.dedup_count}",
            )

        except Exception as e:
            messagebox.showerror("写入错误", f"写入失败：\n{e}")
            self.status_var.set("写入失败")

    def _on_open_dir(self):
        """打开Excel文件所在目录"""
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("提示", "没有文件路径")
            return
        dir_path = os.path.dirname(path)
        if os.path.exists(dir_path):
            os.startfile(dir_path)
        else:
            messagebox.showerror("错误", f"目录不存在：\n{dir_path}")


def main():
    """主函数"""
    root = ctk.CTk()
    app = DeduplicationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
