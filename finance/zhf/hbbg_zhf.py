#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
合并表格 GUI — Frame 版本

作为编排层，负责：
  - 界面构建（模式切换、文件选择、选项设置、进度条、按钮）
  - 调用处理模式引擎 (hbbg_process.ProcessMergeEngine)
  - 调用通用合并引擎 (hbbg_generic.GenericMergeEngine)
  - 列移动设置 (hbbg_column_move.ColumnMovePanel)

核心逻辑已下沉到 hbbg_process.py / hbbg_generic.py 中。
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback

from .hbbg_process import ProcessMergeEngine
from .hbbg_generic import GenericMergeEngine
from .hbbg_column_move import ColumnMovePanel


class MergerFrame(tk.Frame):
    """合并表格功能 - Frame 版本（支持处理模式 + 通用合并模式）"""

    def __init__(self, parent, status_var=None):
        super().__init__(parent)
        self.parent = parent
        self.status_var = status_var
        self._mode = "process"  # "process" 或 "generic"

        # 共用变量
        self.selected_files = []
        self.process_results = []
        self.last_output_path = None  # 最后一次合并的输出文件路径

        # 处理模式变量
        self.target_col_var = tk.StringVar(value="")

        # 通用合并模式变量
        self.gap_var = tk.StringVar(value="0")
        self.header_file_var = tk.StringVar(value="")
        self.skip_header_var = tk.StringVar(value="1")

        # ---- 创建核心引擎（无 GUI 依赖）----
        self.process_engine = ProcessMergeEngine(
            target_col_letter=self.target_col_var.get(),
        )
        self.generic_engine = GenericMergeEngine(
            gap_rows=0,
        )

        # 构建界面
        self._build_ui()

    # ============================================================
    #  界面构建
    # ============================================================

    def _build_ui(self):
        """
        构建完整界面 — grid 布局，按钮始终固定在底部。

        Grid 行结构:
          Row 0: 上部区域（模式切换 + 文件选择按钮）            — weight=0
          Row 1: 文件列表区                                    — weight=1
          Row 2: 通用合并选项区（空行/表头/跳过/列移动）         — weight=0
          Row 3: 底部固定区（进度条 + 操作按钮）                — weight=0
          Row 4: 使用说明（仅处理模式）                         — weight=0
        """
        main = ttk.Frame(self, padding=16)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        # ================================================================
        #  Row 0: 上部区域（模式切换 + 文件选择按钮）
        # ================================================================
        top_area = ttk.Frame(main)
        top_area.grid(row=0, column=0, sticky="nsew")
        top_area.columnconfigure(0, weight=1)

        # ----- 模式切换区 -----
        mode_frame = ttk.Frame(top_area)
        mode_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        tk.Label(mode_frame, text="合并模式：", font=("Microsoft YaHei UI", 11, "bold")).pack(side=tk.LEFT)

        self.mode_var = tk.StringVar(value="process")
        ttk.Radiobutton(
            mode_frame, text="处理模式（填月份列）",
            variable=self.mode_var, value="process",
            command=self._on_mode_change
        ).pack(side=tk.LEFT, padx=(10, 4))

        ttk.Radiobutton(
            mode_frame, text="通用合并（原样堆叠）",
            variable=self.mode_var, value="generic",
            command=self._on_mode_change
        ).pack(side=tk.LEFT, padx=4)

        # ----- 文件选择区（只保留按钮行 + 处理模式选项）-----
        select_frame = ttk.LabelFrame(top_area, text="文件选择", padding=12)
        select_frame.grid(row=1, column=0, sticky="ew")

        btn_row = ttk.Frame(select_frame)
        btn_row.pack(fill=tk.X, pady=(0, 0))

        self.select_btn = tk.Button(
            btn_row, text="📁 选择Excel文件", command=self.select_files,
            font=('微软雅黑', 11), bg="#4CAF50", fg="white",
            width=16, cursor="hand2"
        )
        self.select_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.clear_btn = tk.Button(
            btn_row, text="🗑 清空列表", command=self.clear_files,
            font=('微软雅黑', 11), bg="#f44336", fg="white",
            width=12, cursor="hand2"
        )
        self.clear_btn.pack(side=tk.LEFT)

        self.file_count_label = ttk.Label(btn_row, text="已选择 0 个文件", font=('微软雅黑', 9), foreground='#666666')
        self.file_count_label.pack(side=tk.RIGHT, padx=10)

        # 处理模式选项（仍在 select_frame 内）
        self._process_opt_frame = ttk.Frame(select_frame)
        ttk.Label(self._process_opt_frame, text="填写月份的目标列（如 AU）：", font=('微软雅黑', 10)).pack(side=tk.LEFT, padx=(0, 6))
        self.target_col_entry = ttk.Entry(self._process_opt_frame, textvariable=self.target_col_var, font=('微软雅黑', 11), width=12)
        self.target_col_entry.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(self._process_opt_frame, text="(请输入 Excel 列字母: A, B, C ... AU, AV 等)", font=('微软雅黑', 9), foreground="#888888").pack(side=tk.LEFT)

        # ================================================================
        #  Row 2: 通用合并选项区（独立 LabelFrame，放在文件列表之后）
        # ================================================================
        generic_options_frame = ttk.LabelFrame(main, text="通用合并选项", padding=10)
        generic_options_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        generic_options_frame.columnconfigure(0, weight=1)
        self._generic_options_frame = generic_options_frame

        # 文件间空行数
        self._generic_opt_frame = ttk.Frame(generic_options_frame)
        ttk.Label(self._generic_opt_frame, text="文件间空行数：", font=('微软雅黑', 10)).pack(side=tk.LEFT, padx=(0, 6))
        self.gap_spin = tk.Spinbox(self._generic_opt_frame, textvariable=self.gap_var, from_=0, to=100, width=6, font=('微软雅黑', 10))
        self.gap_spin.pack(side=tk.LEFT)
        ttk.Label(self._generic_opt_frame, text="行（0=不插空行）", font=('微软雅黑', 9), foreground="#888888").pack(side=tk.LEFT, padx=(4, 0))

        # 保留表头文件选择
        self._generic_header_row = ttk.Frame(generic_options_frame)
        ttk.Label(self._generic_header_row, text="保留表头文件：", font=('微软雅黑', 10)).pack(side=tk.LEFT, padx=(0, 6))
        self.header_combo = ttk.Combobox(
            self._generic_header_row, textvariable=self.header_file_var,
            state="disabled", width=32, font=('微软雅黑', 10)
        )
        self.header_combo.pack(side=tk.LEFT)
        self.header_combo.bind("<<ComboboxSelected>>", self._on_header_combo_changed)
        ttk.Label(self._generic_header_row, text="（不选则所有文件均保留表头）", font=('微软雅黑', 9), foreground="#888888").pack(side=tk.LEFT, padx=(4, 0))

        # 跳过表头行数
        self._generic_skip_row = ttk.Frame(generic_options_frame)
        ttk.Label(self._generic_skip_row, text="跳过表头：", font=('微软雅黑', 10)).pack(side=tk.LEFT, padx=(0, 6))
        self.skip_spin = tk.Spinbox(
            self._generic_skip_row, textvariable=self.skip_header_var,
            from_=0, to=100, width=6, font=('微软雅黑', 10), state="disabled"
        )
        self.skip_spin.pack(side=tk.LEFT)
        ttk.Label(self._generic_skip_row, text="行（仅当选了保留表头文件时生效，0=不跳过）", font=('微软雅黑', 9), foreground="#888888").pack(side=tk.LEFT, padx=(4, 0))

        # 列移动设置
        self.col_move_panel = ColumnMovePanel(
            generic_options_frame,
            available_files_callback=lambda: [os.path.basename(f) for f in self.selected_files]
        )

        # ================================================================
        #  Row 1: 文件列表区（可伸缩，自动填充剩余空间）
        # ================================================================
        list_frame = ttk.LabelFrame(main, text="已选择的文件列表", padding=8)
        list_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        scrollbar_y = tk.Scrollbar(list_frame)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x = tk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self.file_listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            font=('微软雅黑', 10), selectmode=tk.MULTIPLE, height=5
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.config(command=self.file_listbox.yview)
        scrollbar_x.config(command=self.file_listbox.xview)

        # ================================================================
        #  Row 3: 底部固定区（进度条 + 操作按钮 — 始终可见）
        # ================================================================
        bottom_area = ttk.Frame(main)
        bottom_area.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        bottom_area.columnconfigure(0, weight=1)

        # 进度条行
        progress_frame = ttk.Frame(bottom_area)
        progress_frame.pack(fill=tk.X, pady=(0, 8))

        self.progress_status = ttk.Label(progress_frame, text="请选择需要合并的Excel文件", font=('微软雅黑', 10), foreground='#666666')
        self.progress_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress = ttk.Progressbar(progress_frame, length=200, mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=(10, 0))

        # 合并按钮 + 打开输出目录按钮（始终在底部可见）
        btn_center = ttk.Frame(bottom_area)
        btn_center.pack()

        self.merge_btn = tk.Button(
            btn_center, text="▶ 开始处理并合并", command=self._on_merge_click,
            font=('微软雅黑', 13, 'bold'), bg="#2196F3", fg="white",
            activebackground="#1976D2", width=22, height=2, cursor="hand2"
        )
        self.merge_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.open_dir_btn = tk.Button(
            btn_center, text="📂 打开输出目录", command=self._open_output_dir,
            font=('微软雅黑', 13, 'bold'), bg="#FF9800", fg="white",
            activebackground="#F57C00", width=22, height=2, cursor="hand2",
            state=tk.DISABLED
        )
        self.open_dir_btn.pack(side=tk.LEFT)

        # ===== 说明区（仅处理模式显示） =====
        self.note_frame = ttk.LabelFrame(main, text="使用说明", padding=10)
        self.note_frame.grid(row=4, column=0, sticky="ew", pady=(8, 0))

        # 初始显示处理模式
        self._on_mode_change()

        # 延迟调整父窗口大小
        self.after(50, self._auto_resize_window)

    def _on_mode_change(self):
        """模式切换：显示/隐藏对应选项，更新说明和按钮文字"""
        mode = self.mode_var.get()
        self._mode = mode

        if mode == "process":
            self._process_opt_frame.pack(fill=tk.X, pady=(8, 0))
            self._generic_options_frame.grid_remove()
            self.merge_btn.config(text="▶ 开始处理并合并")
            self.note_frame.grid()
            self._update_note_text()
            self.after(100, lambda: self.winfo_toplevel().geometry("900x850"))
        else:
            self._process_opt_frame.pack_forget()
            self._generic_options_frame.grid()
            self._generic_opt_frame.pack(fill=tk.X, pady=(6, 0))
            self._generic_header_row.pack(fill=tk.X, pady=6)
            self._generic_skip_row.pack(fill=tk.X, pady=6)
            self.col_move_panel.pack(fill=tk.X, pady=(6, 4))
            self.merge_btn.config(text="▶ 开始合并")
            self.note_frame.grid_remove()
            self.after(100, lambda: self.winfo_toplevel().geometry("900x850"))

    def _auto_resize_window(self):
        """布局完成后自动调整顶层窗口大小"""
        self.update_idletasks()
        top = self.winfo_toplevel()
        content_height = self.winfo_reqheight()
        current_w = top.winfo_width()
        current_h = top.winfo_height()
        if content_height > current_h:
            new_h = content_height + 40
            screen_h = top.winfo_screenheight()
            if new_h > screen_h - 80:
                new_h = screen_h - 80
            top.geometry(f"{current_w}x{new_h}")

    def _update_note_text(self):
        """根据当前模式更新说明文字"""
        for widget in self.note_frame.winfo_children():
            widget.destroy()

        if self._mode == "process":
            text = (
                "【处理模式】\n"
                "1）点击'选择Excel文件'，可选择多个文件\n"
                "2）填写目标列字母，程序会在该列填入对应月份\n"
                "3）点击'开始处理并合并'，程序将逐个处理并合并到汇总表\n"
                "4）处理过程中会保留原文件的格式（字体、边框、合并单元格等）"
            )
            ttk.Label(self.note_frame, text=text, justify="left", foreground='#666666').pack(anchor="w")

    # ============================================================
    #  文件选择（两种模式共用）
    # ============================================================

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择需要合并的Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")],
            parent=self.winfo_toplevel()
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            count = len(self.selected_files)
            self.file_count_label.config(text=f"已选择 {count} 个文件")
            self.progress_status.config(text=f"已选择 {count} 个文件")
            if self.status_var:
                self.status_var.set(f"合并表格 - 已选择 {count} 个文件")

            self._refresh_header_combo()
            if hasattr(self, 'header_combo') and self._mode == "generic":
                self.header_combo.config(state="readonly")
            if hasattr(self, 'col_move_panel'):
                self.col_move_panel.set_available_files(self.selected_files)

    def clear_files(self):
        self.selected_files = []
        self.file_listbox.delete(0, tk.END)
        self.process_results.clear()
        self.target_col_var.set("")
        self.file_count_label.config(text="已选择 0 个文件")
        self.progress_status.config(text="请选择需要合并的Excel文件")
        self.progress['value'] = 0

        if hasattr(self, 'header_file_var'):
            self.header_file_var.set("")
            self.header_combo.config(state="disabled")
            self.header_combo['values'] = []
            self.skip_header_var.set("1")
            self.skip_spin.config(state="disabled")

        if hasattr(self, 'col_move_panel'):
            self.col_move_panel.clear_all()

        if self.status_var:
            self.status_var.set("合并表格 - 已清空列表")

    def _refresh_header_combo(self):
        names = ["（不指定）"] + [os.path.basename(f) for f in self.selected_files]
        self.header_combo['values'] = names
        current = self.header_file_var.get()
        if not current or current not in names:
            self.header_file_var.set("（不指定）")
            self.skip_spin.config(state="disabled")

    def _on_header_combo_changed(self, _event=None):
        val = self.header_file_var.get()
        if val == "（不指定）" or val == "":
            self.skip_spin.config(state="disabled")
        else:
            self.skip_spin.config(state="normal")

    def _on_merge_click(self):
        if self._mode == "process":
            self._do_process_merge()
        else:
            self._do_generic_merge()

    # ============================================================
    #  处理模式（编排层）
    # ============================================================

    def _do_process_merge(self):
        """处理模式编排：校验 → 逐文件处理 → 合并汇总 → 清理"""
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择需要合并的Excel文件！", parent=self.winfo_toplevel())
            return

        target_col = self.target_col_var.get().strip().upper()
        if not target_col:
            messagebox.showwarning("提示", "请先填写目标列（如 AU）！", parent=self.winfo_toplevel())
            return

        self.merge_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.process_results.clear()

        # 更新引擎的目标列
        self.process_engine.target_col_letter = target_col

        total_files = len(self.selected_files)
        success_count = 0
        error_count = 0
        skip_count = 0
        processed_paths = []
        temp_files_to_clean = []

        try:
            for i, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                self.progress_status.config(text=f"[{i+1}/{total_files}] 正在处理: {filename}")
                self.progress['value'] = (i / total_files) * 80
                self.update_idletasks()
                if self.status_var:
                    self.status_var.set(f"正在处理 [{i+1}/{total_files}]: {filename}")

                result = self.process_engine.process_single_file(file_path)
                self.process_results.append(result)

                if result["status"] == "success":
                    success_count += 1
                    processed_paths.append(result["temp_path"])
                    temp_files_to_clean.append(result["temp_path"])
                elif result["status"] == "skip":
                    skip_count += 1
                    processed_paths.append(result["file"])
                else:
                    error_count += 1
                    processed_paths.append(result["file"])

            if success_count == 0:
                messagebox.showwarning("警告",
                    f"没有成功处理的文件！\n\n跳过: {skip_count}\n错误: {error_count}",
                    parent=self.winfo_toplevel())
                self.process_engine.cleanup(temp_files_to_clean)
                return

            self.progress_status.config(text="请选择保存位置...")
            self.update_idletasks()

            output_path = filedialog.asksaveasfilename(
                title="选择汇总表保存位置", defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")], initialfile="汇总表.xlsx",
                parent=self.winfo_toplevel()
            )
            if not output_path:
                messagebox.showinfo("提示", "已取消保存", parent=self.winfo_toplevel())
                self.process_engine.cleanup(temp_files_to_clean)
                return

            self.progress_status.config(text="正在合并文件...")
            self.progress['value'] = 85
            self.update_idletasks()
            if self.status_var:
                self.status_var.set("正在合并文件...")

            self.process_engine.merge_processed_files(processed_paths, output_path)
            self.process_engine.cleanup(temp_files_to_clean)

            self.progress['value'] = 100
            final_msg = f"完成! 成功:{success_count} 跳过:{skip_count} 错误:{error_count}"
            self.progress_status.config(text=final_msg)
            if self.status_var:
                self.status_var.set(final_msg)

            self.last_output_path = output_path
            self.open_dir_btn.config(state=tk.NORMAL)

            summary = ProcessMergeEngine.build_result_summary(
                success_count, skip_count, error_count, output_path, self.process_results
            )
            messagebox.showinfo("处理完成", summary, parent=self.winfo_toplevel())

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：\n{str(e)}", parent=self.winfo_toplevel())
            if self.status_var:
                self.status_var.set("处理失败")
        finally:
            self.merge_btn.config(state=tk.NORMAL)

    # ============================================================
    #  通用合并模式（编排层）
    # ============================================================

    def _do_generic_merge(self):
        """通用合并编排：校验 → 解析参数 → 委托引擎合并"""
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择需要合并的Excel文件！", parent=self.winfo_toplevel())
            return

        output_path = filedialog.asksaveasfilename(
            title="选择合并后文件保存位置", defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")], initialfile="合并结果.xlsx",
            parent=self.winfo_toplevel()
        )
        if not output_path:
            return

        self.merge_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0

        try:
            # ---- 参数解析 ----
            try:
                gap_rows = int(self.gap_var.get())
            except ValueError:
                gap_rows = 0

            try:
                skip_rows = int(self.skip_header_var.get())
            except ValueError:
                skip_rows = 1

            # 保留表头文件
            header_basename = self.header_file_var.get()
            if header_basename in ("（不指定）", ""):
                keep_header_file = None
            else:
                keep_header_file = None
                for f in self.selected_files:
                    if os.path.basename(f) == header_basename:
                        keep_header_file = f
                        break

            # 列移动规则
            col_move_rules = []
            if hasattr(self, 'col_move_panel'):
                col_move_rules = self.col_move_panel.get_rules()
                if col_move_rules:
                    print(f"\n[列移动] 共 {len(col_move_rules)} 条规则:")
                    for r in col_move_rules:
                        print(f"  {r}")

            # 文件顺序：保留表头文件排第一
            if keep_header_file:
                ordered_files = [keep_header_file] + [
                    f for f in self.selected_files if f != keep_header_file
                ]
            else:
                ordered_files = list(self.selected_files)

            # ---- 更新引擎参数 + 回调 ----
            self.generic_engine.gap_rows = gap_rows

            def progress_cb(pct, text):
                self.progress['value'] = pct
                self.progress_status.config(text=text)
                self.update_idletasks()

            def status_cb(text):
                if self.status_var:
                    self.status_var.set(text)

            self.generic_engine._progress_cb = progress_cb
            self.generic_engine._status_cb = status_cb

            # ---- 委托引擎执行 ----
            total = len(ordered_files)
            self.generic_engine.do_merge(
                ordered_files, output_path,
                keep_header_file=keep_header_file,
                skip_rows=skip_rows,
                col_move_rules=col_move_rules,
            )

            self.progress['value'] = 100
            self.progress_status.config(text=f"合并完成！共 {total} 个文件")

            self.last_output_path = output_path
            self.open_dir_btn.config(state=tk.NORMAL)

            messagebox.showinfo("完成",
                f"合并完成！\n\n共处理 {total} 个文件\n保存位置:\n{output_path}",
                parent=self.winfo_toplevel())
            if self.status_var:
                self.status_var.set(f"通用合并 - 完成，共 {total} 个文件")

        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"合并过程中出现错误：\n{str(e)}", parent=self.winfo_toplevel())
        finally:
            self.merge_btn.config(state=tk.NORMAL)

    # ============================================================
    #  工具方法
    # ============================================================

    def _open_output_dir(self):
        """打开输出文件所在目录"""
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showwarning("提示", "输出文件不存在或路径无效！", parent=self.winfo_toplevel())
            return
        output_dir = os.path.dirname(self.last_output_path)
        try:
            os.startfile(output_dir)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开输出目录：\n{e}", parent=self.winfo_toplevel())
