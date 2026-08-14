# -*- coding: utf-8 -*-
"""hbgs_gui.py - 合并个税工具 GUI (CustomTkinter)

入口: HbgsApp(parent) -> 弹出 CTkToplevel 窗口
核心逻辑见 finance/hbgs_core.py (run_merge_collect / write_merge_output / load_fee_rules)
"""

import os
import sys
import threading
from datetime import datetime
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox

# 确保能导入项目根目录模块
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from ui_components import Theme, darken_color
from finance.hbgs_core import run_merge_collect, write_merge_output, load_fee_rules


class HbgsApp:
    """合并个税工具窗口"""

    def __init__(self, parent):
        self.parent = parent
        self.window = ctk.CTkToplevel(parent)
        self.window.title("合并个税工具")
        self.window.geometry("800x610")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(fg_color=Theme.COLOR_BG_CONTENT)

        self.scan_var = tk.StringVar()
        self.rules_var = tk.StringVar()
        self.cancel_event = threading.Event()
        self.running = False
        self._last_result = None  # 内存中的汇总结果, 供保存对话框使用

        self._build_ui()
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)
        self.window.lift()
        self.window.focus_force()

    # ============================================================
    # UI 构建
    # ============================================================

    def _build_ui(self):
        scroll = ctk.CTkFrame(self.window, fg_color=Theme.COLOR_BG_CONTENT)
        scroll.pack(fill="both", expand=True, padx=Theme.SPACING_XL, pady=Theme.SPACING_MD)

        ctk.CTkLabel(scroll, text="🧮 合并个税工具",
                     font=Theme.FONT_HEADING,
                     text_color=Theme.COLOR_ACCENT_FINANCE).pack(anchor="w", pady=(0, Theme.SPACING_XS))
        ctk.CTkLabel(scroll, text="递归合并个税统计表, 按规则计算手续费, 完成后自行选择保存位置",
                     font=Theme.FONT_BODY,
                     text_color=Theme.COLOR_TEXT_SECONDARY).pack(anchor="w", pady=(0, Theme.SPACING_MD))

        # 扫描目录
        self._add_path_row(scroll, "扫描目录:", self.scan_var,
                           lambda: self._browse_dir(self.scan_var), "浏览")
        # 规则表
        self._add_path_row(scroll, "手续费规则表:", self.rules_var,
                           self._browse_rules, "选择")

        # 按钮行
        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, Theme.SPACING_MD))
        self._btn_start = ctk.CTkButton(btn_row, text="▶ 开始合并", command=self._start,
                                        fg_color=Theme.COLOR_ACCENT_FINANCE,
                                        hover_color=darken_color(Theme.COLOR_ACCENT_FINANCE, 0.25),
                                        font=Theme.FONT_BOLD, height=40, width=160)
        self._btn_start.pack(side="left", padx=(0, Theme.SPACING_SM))
        self._btn_cancel = ctk.CTkButton(btn_row, text="■ 取消", command=self._cancel,
                                         fg_color=Theme.COLOR_ACCENT_HR,
                                         hover_color=darken_color(Theme.COLOR_ACCENT_HR, 0.25),
                                         font=Theme.FONT_BOLD, height=40, width=120,
                                         state="disabled")
        self._btn_cancel.pack(side="left")

        # 进度
        self._progress = ctk.CTkProgressBar(scroll, height=12)
        self._progress.pack(fill="x", pady=(0, Theme.SPACING_XS))
        self._progress.set(0)
        self._progress_label = ctk.CTkLabel(scroll, text="就绪",
                                            font=Theme.FONT_CAPTION,
                                            text_color=Theme.COLOR_TEXT_CAPTION)
        self._progress_label.pack(anchor="w", pady=(0, Theme.SPACING_MD))

        # 日志
        ctk.CTkLabel(scroll, text="处理日志:",
                     font=Theme.FONT_SUBHEADING,
                     text_color=Theme.COLOR_TEXT_PRIMARY).pack(anchor="w")
        self._log = ctk.CTkTextbox(scroll, height=240, font=Theme.FONT_SMALL,
                                   fg_color=Theme.COLOR_BG_CARD,
                                   border_color=Theme.COLOR_BORDER, border_width=1)
        self._log.pack(fill="both", expand=False, pady=(Theme.SPACING_XS, 0))

    def _add_path_row(self, parent, label, var, command, btn_text):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=(0, Theme.SPACING_SM))
        ctk.CTkLabel(frame, text=label, font=Theme.FONT_BODY,
                     text_color=Theme.COLOR_TEXT_PRIMARY, width=150, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(frame, textvariable=var, font=Theme.FONT_BODY,
                             state="readonly")
        entry.pack(side="left", fill="x", expand=True, padx=(0, Theme.SPACING_SM))
        ctk.CTkButton(frame, text=btn_text, command=command,
                      fg_color=Theme.COLOR_NAV_BUTTON,
                      hover_color=Theme.COLOR_NAV_BUTTON_ACTIVE,
                      font=Theme.FONT_BOLD, width=80).pack(side="right")

    # ============================================================
    # 交互
    # ============================================================

    def _browse_dir(self, var):
        p = filedialog.askdirectory(title="选择文件夹", parent=self.window)
        if p:
            var.set(p)

    def _browse_rules(self):
        p = filedialog.askopenfilename(title="选择手续费规则表",
                                       filetypes=[("Excel文件", "*.xlsx *.xls"),
                                                  ("所有文件", "*.*")],
                                       parent=self.window)
        if p:
            self.rules_var.set(p)
            try:
                rules = load_fee_rules(p, on_log=self._append_log)
                self._append_log(f"规则表已加载, 共 {len(rules)} 条规则")
            except Exception as e:
                self._append_log(f"⚠ 读取规则表出错: {e}")

    def _start(self):
        if self.running:
            return
        scan = self.scan_var.get()
        rules = self.rules_var.get()
        if not scan or not os.path.isdir(scan):
            messagebox.showwarning("提示", "请先选择扫描目录")
            return
        if not rules or not os.path.isfile(rules):
            messagebox.showwarning("提示", "请先选择手续费规则表")
            return

        self.running = True
        self.cancel_event.clear()
        self._btn_start.configure(state="disabled")
        self._btn_cancel.configure(state="normal")
        self._log.delete("0.0", "end")
        self._append_log("开始合并...")
        self._progress.set(0)
        t = threading.Thread(target=self._worker, args=(scan, rules),
                             daemon=True)
        t.start()

    def _worker(self, scan, rules):
        run_merge_collect(scan_root=scan, rules_path=rules,
                          cancel_event=self.cancel_event,
                          on_log=self._safe_log,
                          on_progress=self._safe_progress,
                          on_done=self._safe_done)

    def _safe_log(self, msg):
        self.window.after(0, lambda: self._append_log(msg))

    def _safe_progress(self, cur, total, msg):
        self.window.after(0, lambda: self._set_progress(cur, total, msg))

    def _safe_done(self, result):
        self.window.after(0, lambda: self._on_done(result))

    def _append_log(self, msg):
        self._log.insert("end", msg + "\n")
        self._log.see("end")

    def _set_progress(self, cur, total, msg):
        if total > 0:
            self._progress.set(cur / total)
        self._progress_label.configure(text=f"{msg}  ({cur}/{total})")

    def _cancel(self):
        self.cancel_event.set()
        self._append_log("正在取消...")

    def _on_done(self, result):
        self.running = False
        self._btn_start.configure(state="normal")
        self._btn_cancel.configure(state="disabled")
        self._progress.set(1)
        self._last_result = result  # 保留内存数据, 取消保存后仍可重新运行

        if result.get("rows"):
            # 弹出保存对话框, 默认文件名: 个税汇总YYYYMMDD.xlsx
            default_name = f"个税汇总{datetime.now():%Y%m%d}.xlsx"
            save_path = filedialog.asksaveasfilename(
                title="保存汇总结果",
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
                parent=self.window)
            if save_path:
                try:
                    write_merge_output(save_path, result["header"],
                                       result["rows"], result["target_id_col"])
                    # 提示汇总 (对应 VBA 末尾 MsgBox 的 ⚠ 文字)
                    extra = ""
                    ek = result.get("empty_keyword_ops")
                    mo = result.get("missing_ops")
                    if ek or mo:
                        parts = []
                        if ek:
                            parts.append(f"规则表有 {len(ek)} 个运营项目未填写扣费关键词")
                        if mo:
                            parts.append(f"源文件有 {len(mo)} 个运营项目缺规则")
                        extra = "\n⚠ " + "；".join(parts) + "（详见日志）"
                    messagebox.showinfo(
                        "完成",
                        f"合并完成！\n扫描 {result['total_files']} 个表格, "
                        f"有效 {result['total_rows']} 条数据, "
                        f"跳过 {result['skipped']} 个文件。{extra}\n已保存至: {save_path}")
                except Exception as e:
                    messagebox.showerror("保存失败", f"写出文件时出错:\n{str(e)}")
            else:
                self._append_log("⚠ 已取消保存（结果仍保留在内存, 可重新运行合并）")
        else:
            messagebox.showwarning("未完成", "未能生成汇总数据, 请查看日志。")

    def _on_close(self):
        if self.running:
            self.cancel_event.set()
        self.window.destroy()
