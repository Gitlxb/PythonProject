# -*- coding: utf-8 -*-
"""
GUI主模块 - SummaryProcessorApp 类（主入口 + Notebook框架 + 共享工具）

Tab 1 UI → hs_split_tab.py (SplitTabMixin)
Tab 2 UI → hs_check_tab.py (CheckTabMixin)
业务逻辑 → hs_gui_actions.py (GuiActionsMixin)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from typing import List, Optional

from finance.cxm.hs_huizong_cxm import SummaryProcessor
from finance.cxm.hs_check_processor import CheckProcessor
from openpyxl import load_workbook
from finance.cxm.hs_gui_actions import GuiActionsMixin


class SummaryProcessorApp(GuiActionsMixin):
    """汇总处理GUI应用程序（双Tab：拆分汇总比较 / 核对）"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel汇总处理工具")
        self.root.geometry("1100x850")
        self.root.minsize(900, 600)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # ── 处理器 ──
        self.processor = SummaryProcessor()
        self.check_processor = CheckProcessor()

        # ── Notebook / Tab 状态 ──
        self.current_tab = 'split'
        self.notebook = None
        self.tab1_frame = None
        self.tab2_frame = None

        # ── Tab 1 状态 ──
        self.selected_sheets: List[str] = []
        self.sheet_listbox: Optional[tk.Listbox] = None
        self.sheet_scrollbar: Optional[ttk.Scrollbar] = None
        self.selected_tags_frame: Optional[tk.Frame] = None
        self.selected_months_for_save = None
        self.last_summary_path = None
        self.last_check_path = None
        self.last_operation = None  # 'summary' | 'check' | None
        self.appended_types = set()  # {'summary', 'check'}

        # ── Tab 1 文件路径变量 ──
        self.file_path_var = tk.StringVar()
        self.check_file_path_var = tk.StringVar()
        self.check_sheet_var = tk.StringVar()
        self.check_file_path = None
        self.check_result_df = None
        self.target_file_path_var = tk.StringVar()
        self.target_file_path = None

        # ── Tab 2 核对状态 ──
        self.check_reconcile_path = None

        # ── 共用控件引用（由各 Tab Mixin 赋值）──
        self.tree = None
        self.status_text = None

        self._setup_treeview_style()
        self._create_ui()

    # ==================== 样式配置 ====================

    def _setup_treeview_style(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=25, font=('Microsoft YaHei UI', 9))
        style.configure("Treeview.Heading", font=('Microsoft YaHei UI', 10, 'bold'))
        style.map("Treeview",
                  background=[('selected', '#0078d7')],
                  foreground=[('selected', 'white')])

    # ==================== 主界面：Notebook 双 Tab ====================

    def _on_tab_changed(self, event=None):
        """Tab切换回调"""
        idx = self.notebook.index('current')
        self.current_tab = 'split' if idx == 0 else 'check'

    def _create_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Notebook
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)

        # Tab 1
        self.tab1_frame = ttk.Frame(self.notebook, padding="15")
        self.tab1_frame.columnconfigure(0, weight=1)
        self.tab1_frame.rowconfigure(2, weight=3)
        self.tab1_frame.rowconfigure(3, weight=1)
        self.notebook.add(self.tab1_frame, text="  拆分汇总比较  ")

        # Tab 2
        self.tab2_frame = ttk.Frame(self.notebook, padding="15")
        self.tab2_frame.columnconfigure(0, weight=1)
        self.tab2_frame.rowconfigure(3, weight=3)
        self.tab2_frame.rowconfigure(4, weight=1)
        self.notebook.add(self.tab2_frame, text="  核对  ")

        # 构建内容（由各自 Mixin 提供）
        self._create_tab1_content(self.tab1_frame)
        self._create_tab2_content(self.tab2_frame)

        self._log_status("欢迎使用Excel汇总处理工具！")
        self._log_status("请选择一个Excel文件开始处理。")

    # ==================== 共用的预览区和日志区 ====================
    # （接受 row 参数，避免硬编码导致不同 Tab 的 grid 冲突）

    def _create_preview_frame(self, parent, row=2):
        """共用数据预览区（Treeview）"""
        preview_frame = ttk.LabelFrame(parent, text="数据预览", padding="12")
        preview_frame.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 12))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(preview_frame, selectmode='browse', show='headings')
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.xview_moveto(0)

        self.tree.bind("<MouseWheel>", self._on_mousewheel)
        self.tree.bind("<Button-4>", self._on_mousewheel)
        self.tree.bind("<Button-5>", self._on_mousewheel)

    def _create_status_frame(self, parent, row=3):
        """共用状态信息区（ScrolledText）"""
        status_frame = ttk.LabelFrame(parent, text="状态信息", padding="12")
        status_frame.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)

        self.status_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, font=('Consolas', 9))
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # ==================== 共用工具方法 ====================

    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.tree.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.tree.yview_scroll(1, "units")

    def _log_status(self, message: str):
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()

    # ==================== 共用操作 ====================

    def _open_output_dir(self):
        target = (self.last_summary_path or self.last_check_path or
                  self.check_reconcile_path or self.target_file_path)
        if not target:
            messagebox.showwarning("警告", "尚未生成任何输出文件")
            return
        dir_path = os.path.dirname(target)
        if os.path.exists(dir_path):
            os.startfile(dir_path)
            self._log_status(f"已打开输出目录：{dir_path}")
        else:
            messagebox.showerror("错误", f"目录不存在：{dir_path}")

    def _clear_all(self):
        """重置 Tab 1 全部状态"""
        self.file_path_var.set("")
        self.selected_sheets = []
        if self.sheet_listbox:
            self.sheet_listbox.delete(0, tk.END)
        if self.selected_tags_frame:
            for widget in self.selected_tags_frame.winfo_children():
                widget.destroy()
        self.selected_months_for_save = None
        self.last_summary_path = None
        self.last_check_path = None
        self.last_operation = None
        self.appended_types = set()
        self.check_file_path_var.set("")
        self.check_file_path = None
        self.check_sheet_var.set("")
        if hasattr(self, 'check_sheet_combo'):
            self.check_sheet_combo['values'] = []
        self.check_result_df = None
        self.target_file_path_var.set("")
        self.target_file_path = None

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree['columns'] = ()

        self.processor = SummaryProcessor()

        self.status_text.delete('1.0', tk.END)
        self._log_status("欢迎使用Excel汇总处理工具！")
        self._log_status("请选择一个Excel文件开始处理。")


# 导入两个 Tab Mixin（追加到类上）
from finance.cxm.hs_split_tab import SplitTabMixin
from finance.cxm.hs_check_tab import CheckTabMixin

# 将 Mixin 方法混入 SummaryProcessorApp
SummaryProcessorApp.__bases__ += (SplitTabMixin, CheckTabMixin)
