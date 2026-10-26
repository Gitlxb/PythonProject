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

from tkcalendar import Calendar

from . import tenure_months_calculator as tmc


DONE_SIGNAL = object()


class DatePicker(ttk.Frame):
    """自定义日期选择器：Entry + 下拉按钮，点击弹出 Calendar 日历。"""

    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self._date = None
        self._popup = None
        self._calendar = None

        self._var = tk.StringVar(value="")
        self._entry = ttk.Entry(self, textvariable=self._var, width=14, state="readonly")
        self._entry.pack(side=tk.LEFT)

        self._btn = ttk.Button(self, text="▼", width=3, command=self._toggle_popup)
        self._btn.pack(side=tk.LEFT, padx=(2, 0))

    def set_date(self, date: datetime.date):
        self._date = date
        self._var.set(date.strftime("%Y-%m-%d"))

    def get_date(self) -> datetime.date:
        return self._date

    def _toggle_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._close_popup()
        else:
            self._open_popup()

    def _open_popup(self):
        self._close_popup()

        self._popup = tk.Toplevel(self)
        self._popup.title("选择日期")
        self._popup.resizable(False, False)
        self._popup.overrideredirect(False)
        self._popup.transient(self.winfo_toplevel())

        self._calendar = Calendar(
            self._popup,
            selectmode="day",
            locale="zh_CN",
            date_pattern="yyyy-mm-dd",
        )
        self._calendar.pack(padx=10, pady=10)

        if self._date:
            self._calendar.selection_set(self._date)

        btn_frame = ttk.Frame(self._popup)
        btn_frame.pack(pady=(0, 8))
        ttk.Button(btn_frame, text="确定", command=self._on_select).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="取消", command=self._close_popup).pack(side=tk.LEFT, padx=4)

        self._calendar.bind("<<CalendarSelected>>", lambda e: self._on_select())
        self._popup.bind("<Escape>", lambda e: self._close_popup())

        self._popup.update_idletasks()
        self._popup.grab_set()

    def _close_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.grab_release()
            self._popup.destroy()
        self._popup = None
        self._calendar = None

    def _on_select(self):
        if self._calendar:
            sel = self._calendar.selection_get()
            if sel:
                self._date = sel if isinstance(sel, datetime.date) else sel.date()
                self._var.set(self._date.strftime("%Y-%m-%d"))
        self._close_popup()


class StatMonthPicker(DatePicker):
    """起始日期选择器。"""
    pass


class AsOfDatePicker(DatePicker):
    """截止日期选择器。"""
    pass


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

        # 起始日期：tkcalendar 下拉日历，返回完整日期
        ttk.Label(frm, text="起始日期：").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.stat_picker = StatMonthPicker(frm)
        self.stat_picker.grid(row=1, column=1, sticky=tk.W, padx=4)
        # 读取默认起始日期
        default_start = tmc.CONFIG.get("start_date", "2026-08-01")
        try:
            default_start_date = datetime.datetime.strptime(default_start, "%Y-%m-%d").date()
            self.stat_picker.set_date(default_start_date)
        except ValueError:
            pass

        # 截止日期：tkcalendar 标准下拉日历
        ttk.Label(frm, text="截止日期：").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.asof_picker = AsOfDatePicker(frm)
        self.asof_picker.grid(row=2, column=1, sticky=tk.W, padx=4)
        default_end = tmc.CONFIG.get("end_date")
        if default_end:
            try:
                default_end_date = datetime.datetime.strptime(default_end, "%Y-%m-%d").date()
                self.asof_picker.set_date(default_end_date)
            except ValueError:
                self.asof_picker.set_date(datetime.date.today())
        else:
            self.asof_picker.set_date(datetime.date.today())

        # 合并选项
        self.merge_chk = ttk.Checkbutton(
            frm, text="合并"'员工总在职天数'"列（数据量大时取消可加速）",
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
        start_date = self.stat_picker.get_date()
        if not start_date:
            messagebox.showwarning("提示", "请选择起始日期")
            return

        end_date = self.asof_picker.get_date()
        if not end_date:
            messagebox.showwarning("提示", "请选择截止日期")
            return

        if start_date > end_date:
            messagebox.showwarning("提示", "起始日期不能晚于截止日期")
            return

        raw_start = start_date.strftime("%Y-%m-%d")
        raw_end = end_date.strftime("%Y-%m-%d")

        merge_total_days = self.merge_var.get()

        self._temp_output = ""
        self._last_output = ""
        self._append_log("")
        self._append_log(f"开始处理：{path}")
        self._append_log(f"起始日期：{raw_start}  |  截止日期：{raw_end}")
        self._append_log(f"合并总在职天数列：{'是' if merge_total_days else '否'}\n")

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.status_var.set("处理中...")

        threading.Thread(target=self._worker, args=(path, raw_start, raw_end, merge_total_days), daemon=True).start()
        self.container.after(80, self._poll)

    def _worker(self, input_file: str, raw_start: str, raw_end: str, merge_total_days: bool):
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
                start_date=raw_start,
                end_date=raw_end,
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
