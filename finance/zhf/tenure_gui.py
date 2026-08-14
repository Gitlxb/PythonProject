# -*- coding: utf-8 -*-
"""
员工在职月份计算 —— GUI 版

职责：
    本文件只负责界面交互：选择输入文件、配置参数、启动后台计算、
    处理完成后弹窗让用户选择输出位置，并打开结果所在目录。

    所有业务计算逻辑均委托给 tenure_months_calculator.process()。

依赖：
    tkinter（Python 标准库）；pandas/openpyxl（由核心模块使用）

运行方式：
    内嵌于绩效考核处理工具（finance/zhf/main_app.py）的导航页面中
"""

import datetime
import os
import queue
import shutil
import sys
import tempfile
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

from tkcalendar import DateEntry

from . import tenure_months_calculator as tmc


DONE_SIGNAL = object()


class StatMonthPicker(ttk.Frame):
    """截止年月份选择器：使用 tkcalendar.DateEntry，显示完整日期，程序取年月。"""

    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self._entry = DateEntry(self, width=14, date_pattern="yyyy-mm-dd", locale="zh_CN")
        self._entry.pack(side=tk.LEFT)

    def set_date(self, date: datetime.date):
        self._entry.set_date(date)

    def get_date(self) -> datetime.date:
        return self._entry.get_date()

    def get_ym(self) -> str:
        return self.get_date().strftime("%Y-%m")


class AsOfDatePicker(ttk.Frame):
    """在职截止日选择器：使用 tkcalendar.DateEntry 标准日历下拉。"""

    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self._entry = DateEntry(self, width=14, date_pattern="yyyy-mm-dd", locale="zh_CN")
        self._entry.pack(side=tk.LEFT)

    def set_date(self, date: datetime.date):
        self._entry.set_date(date)

    def get_date(self) -> datetime.date:
        return self._entry.get_date()


class StdoutRedirector:
    """把 print 输出重定向到队列，供主线程刷新日志框。"""

    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, text: str):
        if text:
            self.q.put(text)

    def flush(self):
        pass


class App:
    """员工在职月份计算 —— 界面逻辑（嵌入绩效考核工具作为导航子页面）。"""

    def __init__(self, parent, status_var=None):
        self.container = parent  # 父容器

        self.log_q = queue.Queue()
        self.running = False
        self.merge_var = tk.BooleanVar(value=True)  # 是否合并总在职天数列
        self.status_var = status_var if status_var is not None else tk.StringVar(value="就绪")
        self._temp_output: str = ""          # 后台计算生成的临时结果文件
        self._last_output: str = ""          # 用户最终选择的保存路径

        self._build_ui()

    # ----------------------------- 界面构建 -----------------------------
    def _build_ui(self):
        # —— 参数区 ——
        frm = ttk.Frame(self.container, padding=10)
        frm.pack(fill=tk.X)
        # 固定列宽，避免窗口拉伸时输入框和按钮间距变大
        frm.columnconfigure(0, weight=0)
        frm.columnconfigure(1, weight=0)
        frm.columnconfigure(2, weight=0)

        # Excel 文件行：输入框在 column 1，浏览按钮在 column 2（与下方说明文字同列左对齐）
        ttk.Label(frm, text="Excel 文件：").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.file_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.file_var, width=55).grid(row=0, column=1, sticky=tk.W+tk.E, padx=4)
        ttk.Button(frm, text="浏览...", command=self._browse_input).grid(row=0, column=2, sticky=tk.W, padx=4)

        # 截止年月份：tkcalendar 下拉日历，显示完整日期，程序取年月
        ttk.Label(frm, text="截止年月份：").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.stat_picker = StatMonthPicker(frm)
        self.stat_picker.grid(row=1, column=1, sticky=tk.W, padx=4)
        # 读取默认统计月
        default_stat = tmc.CONFIG.get("stat_month", "2026-07")
        try:
            default_stat_date = datetime.datetime.strptime(default_stat, "%Y-%m").date()
            self.stat_picker.set_date(default_stat_date)
        except ValueError:
            pass
        ttk.Label(frm, text="选择任意日期，程序自动取该年月作为统计截止月份",
                  anchor=tk.W).grid(row=1, column=2, sticky=tk.W, padx=4, pady=4)

        # 在职截止日：tkcalendar 标准下拉日历
        ttk.Label(frm, text="在职截止日：").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.asof_picker = AsOfDatePicker(frm)
        self.asof_picker.grid(row=2, column=1, sticky=tk.W, padx=4)
        self.asof_picker.set_date(datetime.date.today())
        ttk.Label(frm, text="默认今天（决定在职天数截到哪天）",
                  anchor=tk.W).grid(row=2, column=2, sticky=tk.W, padx=4, pady=4)

        # 合并选项
        self.merge_chk = ttk.Checkbutton(
            frm, text="合并"'总在职天数'"列（数据量大时取消可加速）",
            variable=self.merge_var)
        self.merge_chk.grid(row=3, column=1, sticky=tk.W, padx=4, pady=4)

        # —— 操作区 ——
        btn_frm = ttk.Frame(self.container, padding=(10, 0))
        btn_frm.pack(fill=tk.X, pady=6)

        self.start_btn = ttk.Button(btn_frm, text="开始处理", command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=4)

        ttk.Button(btn_frm, text="打开输出目录", command=self._open_output).pack(side=tk.LEFT, padx=4)

        ttk.Label(btn_frm, textvariable=self.status_var).pack(side=tk.LEFT, padx=12)

        # —— 日志区 ——
        ttk.Label(self.container, text="日志：").pack(anchor=tk.W, padx=10)
        self.log = scrolledtext.ScrolledText(self.container, height=26, state=tk.DISABLED,
                                             font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

    # ----------------------------- 交互逻辑 -----------------------------
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择员工信息 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if path:
            self.file_var.set(path)

    def _open_output(self):
        d = self._last_output or self._temp_output
        if not d:
            messagebox.showinfo("提示", "尚未生成结果文件。")
            return
        d = os.path.dirname(os.path.abspath(d))
        if not os.path.isdir(d):
            messagebox.showerror("错误", f"目录不存在：\n{d}")
            return
        try:
            os.startfile(d)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录：{e}")

    def _start(self):
        if self.running:
            return

        path = self.file_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return
        if not os.path.isfile(path):
            messagebox.showerror("错误", f"文件不存在：\n{path}")
            return

        # 参数校验与取值
        stat_month = self.stat_picker.get_ym()
        if not stat_month:
            messagebox.showwarning("提示", "请选择截止年月份")
            return

        asof_date = self.asof_picker.get_date()
        raw_asof = asof_date.strftime("%Y-%m-%d") if asof_date else None

        merge_total_days = self.merge_var.get()

        self._temp_output = ""
        self._last_output = ""
        self._append_log("")
        self._append_log(f"开始处理：{path}")
        self._append_log(f"截止年月份：{stat_month}  |  在职截止日：{raw_asof or '今天'}")
        self._append_log(f"合并总在职天数列：{'是' if merge_total_days else '否'}\n")

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.status_var.set("处理中...")

        threading.Thread(target=self._worker, args=(path, stat_month, raw_asof, merge_total_days), daemon=True).start()
        self.container.after(80, self._poll)

    def _worker(self, input_file: str, stat_month: str, raw_asof: str, merge_total_days: bool):
        old_stdout = sys.stdout
        sys.stdout = StdoutRedirector(self.log_q)

        # 生成临时结果文件，等主线程弹窗让用户选择最终保存位置
        fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="tenure_result_")
        os.close(fd)
        self._temp_output = temp_path

        try:
            tmc.process(
                input_file=input_file,
                output_file=temp_path,
                stat_month=stat_month,
                as_of_date=raw_asof if raw_asof else None,
                merge_total_days=merge_total_days,
            )
            self.log_q.put("\n✅ 计算完成，等待选择保存位置...\n")
        except Exception:
            # 计算失败时清理临时文件
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
            self._temp_output = ""
            self.log_q.put("\n❌ 处理出错：\n" + traceback.format_exc())
        finally:
            sys.stdout = old_stdout
        self.log_q.put(DONE_SIGNAL)

    def _poll(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                if msg is DONE_SIGNAL:
                    self._on_done()
                    return
                self._append_log(msg)
        except queue.Empty:
            pass
        if self.running:
            self.container.after(80, self._poll)

    def _on_done(self):
        self.running = False
        self.start_btn.config(state=tk.NORMAL)

        if not self._temp_output or not os.path.exists(self._temp_output):
            self.status_var.set("失败")
            return

        # 弹窗让用户选择最终保存位置
        default_name = "员工在职月份_结果.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="选择结果保存位置",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])

        if not save_path:
            # 用户取消：清理临时文件
            self._clean_temp()
            self.status_var.set("已取消保存")
            self._append_log("已取消保存，临时结果文件已清理。\n")
            return

        try:
            shutil.copy2(self._temp_output, save_path)
            self._last_output = save_path
            self.status_var.set("完成")
            self._append_log(f"✅ 结果已保存：{save_path}\n")
        except Exception as e:
            self.status_var.set("保存失败")
            self._append_log(f"❌ 保存结果失败：{e}\n临时结果文件：{self._temp_output}\n")
        finally:
            # 清理临时文件
            try:
                if os.path.exists(self._temp_output):
                    os.remove(self._temp_output)
            except Exception:
                pass
            self._temp_output = ""

    def _clean_temp(self):
        """删除后台计算生成的临时结果文件（若存在）。"""
        try:
            if self._temp_output and os.path.exists(self._temp_output):
                os.remove(self._temp_output)
        except Exception:
            pass
        self._temp_output = ""

    def _append_log(self, text: str):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, text)
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
