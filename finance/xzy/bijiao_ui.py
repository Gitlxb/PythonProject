# -- coding: utf-8 --
# @File : bijiao_ui.py
# @Description: 数据比较工具 - UI组件（日历弹窗和主界面Frame）

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime, date, timedelta
import calendar

from .bijiao_core import WorkerAdvanceComparatorCore
from .bijiao_utils import get_sheet_names
from .bijiao_frame import format_number, save_to_excel


# ==================== 日历弹窗组件 ====================
class CalendarPopup:
    """日历弹窗"""

    def __init__(self, parent, callback):
        self.parent = parent
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("选择日期")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        self.today = date.today()
        self.year = self.today.year
        self.month = self.today.month

        self._build_ui()
        self._draw_calendar()
        self.window.update_idletasks()
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() + 50
        self.window.geometry(f"+{x}+{y}")

    def _build_ui(self):
        nav_frame = tk.Frame(self.window)
        nav_frame.pack(pady=5)

        tk.Button(nav_frame, text="<<", command=lambda: self._change_year(-1),
                  font=("微软雅黑", 9), width=3).pack(side="left", padx=1)
        tk.Button(nav_frame, text="<", command=lambda: self._change_month(-1),
                  font=("微软雅黑", 9), width=3).pack(side="left", padx=1)

        self.month_label = tk.Label(nav_frame, text="", font=("微软雅黑", 11, "bold"), width=12)
        self.month_label.pack(side="left", padx=10)

        tk.Button(nav_frame, text=">", command=lambda: self._change_month(1),
                  font=("微软雅黑", 9), width=3).pack(side="left", padx=1)
        tk.Button(nav_frame, text=">>", command=lambda: self._change_year(1),
                  font=("微软雅黑", 9), width=3).pack(side="left", padx=1)

        weekday_frame = tk.Frame(self.window)
        weekday_frame.pack()
        weekdays = ['日', '一', '二', '三', '四', '五', '六']
        for i, wd in enumerate(weekdays):
            color = "#e74c3c" if i == 0 or i == 6 else "#2c3e50"
            tk.Label(weekday_frame, text=wd, font=("微软雅黑", 9, "bold"),
                     width=4, fg=color).pack(side="left", padx=1)

        self.cal_frame = tk.Frame(self.window)
        self.cal_frame.pack(pady=3)

        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="今天", command=self._select_today,
                  font=("微软雅黑", 9), bg="#3498db", fg="white", width=8).pack(side="left", padx=5)
        tk.Button(btn_frame, text="关闭", command=self.window.destroy,
                  font=("微软雅黑", 9), width=8).pack(side="left", padx=5)

    def _change_month(self, delta):
        self.month += delta
        if self.month > 12:
            self.month = 1; self.year += 1
        elif self.month < 1:
            self.month = 12; self.year -= 1
        self._draw_calendar()

    def _change_year(self, delta):
        self.year += delta
        self._draw_calendar()

    def _draw_calendar(self):
        for w in self.cal_frame.winfo_children():
            w.destroy()
        self.month_label.config(text=f"{self.year}年{self.month:02d}月")

        cal = calendar.monthcalendar(self.year, self.month)
        for week_idx, week in enumerate(cal):
            for day_idx, day in enumerate(week):
                if day == 0:
                    tk.Label(self.cal_frame, text="", width=4, font=("微软雅黑", 9)).grid(
                        row=week_idx, column=day_idx, padx=1, pady=1)
                else:
                    is_today = (self.year == self.today.year and
                                self.month == self.today.month and day == self.today.day)
                    bg = "#3498db" if is_today else "#f0f0f0"
                    fg = "white" if is_today else "#2c3e50"
                    btn = tk.Button(
                        self.cal_frame, text=str(day), width=4, height=1,
                        font=("微软雅黑", 9), bg=bg, fg=fg,
                        command=lambda d=day: self._select_day(d),
                        relief="flat", cursor="hand2")
                    btn.grid(row=week_idx, column=day_idx, padx=1, pady=1)

    def _select_day(self, day):
        selected = f"{self.year}-{self.month:02d}-{day:02d}"
        self.callback(selected)
        self.window.destroy()

    def _select_today(self):
        callback = self.callback
        self.callback(self.today.strftime("%Y-%m-%d"))
        self.window.destroy()


# ==================== 主界面Frame组件 ====================
class BijiaoFrame(ttk.Frame):
    """数据比较功能 - 作为Frame嵌入主界面"""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.core = WorkerAdvanceComparatorCore()
        self.info_labels = {}
        self._date_range_refs = {}
        self.setup_ui()

    def _get_toplevel(self):
        return self.winfo_toplevel()

    # ==================== UI构建 ====================

    def setup_ui(self):
        # 标题栏
        title_frame = tk.Frame(self, bg="#2c3e50")
        title_frame.pack(fill="x")
        tk.Label(
            title_frame, text="工人预支申请表 vs 预支明细表 对比工具",
            font=("微软雅黑", 16, "bold"), fg="white", bg="#2c3e50",
            pady=12
        ).pack()

        # Notebook 分步容器
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # 步骤1：文件选择
        self.step1_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step1_frame, text="  ① 选择文件  ")
        self.setup_step1_ui()

        # 步骤2：筛选日期
        self.step2_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step2_frame, text="  ② 筛选日期  ")
        self.setup_step2_ui()

        # 步骤3：结果预览
        self.step3_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step3_frame, text="  ③ 结果预览  ")
        self.setup_step3_ui()

        # 状态栏
        status_frame = tk.Frame(self, bg="#e9ecef")
        status_frame.pack(fill="x", side="bottom")

        self.status_label = tk.Label(
            status_frame, text="请先选择两个Excel文件", bg="#e9ecef",
            font=("微软雅黑", 10), anchor="w", padx=10, pady=5
        )
        self.status_label.pack(fill="x")

        # 默认禁用步骤2、3
        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")

    # ==================== 步骤1：文件选择 ====================

    def setup_step1_ui(self):
        canvas = tk.Canvas(self.step1_frame, bg="#f8f9fa", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.step1_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f8f9fa")
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 文件1
        file1_box = tk.LabelFrame(
            scrollable_frame, text=" 第一个表格：工人预支申请表 ",
            font=("微软雅黑", 12, "bold"), bg="#f8f9fa", padx=15, pady=10
        )
        file1_box.pack(fill="x", padx=20, pady=(20, 10))

        tk.Label(
            file1_box,
            text="选择包含工人预支申请数据的Excel文件。\n列位置固定为：F-预支金额, I-姓名, N-提交时间, S-身份证号（表头在第1行）",
            font=("微软雅黑", 9), bg="#f8f9fa", fg="#6c757d",
            wraplength=700, justify="left"
        ).pack(anchor="w", pady=(0, 8))

        btn_frame1 = tk.Frame(file1_box, bg="#f8f9fa")
        btn_frame1.pack(fill="x")
        self.select_btn1 = tk.Button(
            btn_frame1, text="📂 选择工人预支申请表",
            command=self.select_first_file, font=("微软雅黑", 11),
            bg="#3498db", fg="white", padx=15, pady=5, cursor="hand2"
        )
        self.select_btn1.pack(side="left")
        self.file1_label = tk.Label(btn_frame1, text="未选择文件", bg="#f8f9fa",
                                    font=("微软雅黑", 10), fg="#6c757d", anchor="w")
        self.file1_label.pack(side="left", padx=(15, 0), fill="x", expand=True)
        self.sheet1_label = tk.Label(file1_box, text="", bg="#f8f9fa",
                                     font=("微软雅黑", 10), fg="#28a745", anchor="w")
        self.sheet1_label.pack(anchor="w", pady=(5, 0))

        # 分隔线
        tk.Frame(scrollable_frame, height=2, bg="#dee2e6").pack(fill="x", padx=20, pady=10)

        # 文件2
        file2_box = tk.LabelFrame(
            scrollable_frame, text=" 第二个表格：预支明细表 ",
            font=("微软雅黑", 12, "bold"), bg="#f8f9fa", padx=15, pady=10
        )
        file2_box.pack(fill="x", padx=20, pady=(10, 10))

        tk.Label(
            file2_box,
            text="选择包含预支明细数据的Excel文件（第二行为表头），需包含：预支日期、单位、姓名、预支数额",
            font=("微软雅黑", 9), bg="#f8f9fa", fg="#6c757d",
            wraplength=700, justify="left"
        ).pack(anchor="w", pady=(0, 8))

        btn_frame2 = tk.Frame(file2_box, bg="#f8f9fa")
        btn_frame2.pack(fill="x")
        self.select_btn2 = tk.Button(
            btn_frame2, text="📂 选择预支明细表",
            command=self.select_second_file, font=("微软雅黑", 11),
            bg="#e67e22", fg="white", padx=15, pady=5, cursor="hand2"
        )
        self.select_btn2.pack(side="left")
        self.file2_label = tk.Label(btn_frame2, text="未选择文件", bg="#f8f9fa",
                                    font=("微软雅黑", 10), fg="#6c757d", anchor="w")
        self.file2_label.pack(side="left", padx=(15, 0), fill="x", expand=True)
        self.sheet2_label = tk.Label(file2_box, text="", bg="#f8f9fa",
                                     font=("微软雅黑", 10), fg="#28a745", anchor="w")
        self.sheet2_label.pack(anchor="w", pady=(5, 0))

        # 进入下一步按钮
        tk.Frame(scrollable_frame, height=10, bg="#f8f9fa").pack()
        self.next_step_btn = tk.Button(
            scrollable_frame, text="▶ 进入下一步 - 设置日期范围",
            command=self.go_to_step2, font=("微软雅黑", 12, "bold"),
            bg="#28a745", fg="white", padx=20, pady=8, cursor="hand2", state="disabled"
        )
        self.next_step_btn.pack(pady=15)

    # ==================== 步骤2：筛选日期 ====================

    def setup_step2_ui(self):
        info_frame = tk.Frame(self.step2_frame, bg="#f8f9fa")
        info_frame.pack(fill="x", padx=30, pady=(25, 5))
        tk.Label(info_frame, text="筛选日期范围", font=("微软雅黑", 15, "bold"),
                 bg="#f8f9fa", fg="#2c3e50").pack(anchor="w")

        file_info_frame = tk.LabelFrame(
            self.step2_frame, text=" 文件信息 ",
            font=("微软雅黑", 11, "bold"), bg="#f8f9fa", padx=15, pady=8
        )
        file_info_frame.pack(padx=30, pady=(10, 5), fill="x")
        self.info_labels["table1"] = tk.Label(file_info_frame, text="未选择文件",
                                              bg="#f8f9fa", font=("微软雅黑", 10), anchor="w")
        self.info_labels["table1"].pack(anchor="w", pady=2)
        self.info_labels["table2"] = tk.Label(file_info_frame, text="未选择文件",
                                              bg="#f8f9fa", font=("微软雅黑", 10), anchor="w")
        self.info_labels["table2"].pack(anchor="w", pady=2)

        date_info_frame = tk.LabelFrame(
            self.step2_frame, text=" 数据中的日期范围 ",
            font=("微软雅黑", 11, "bold"), bg="#f8f9fa", padx=15, pady=8
        )
        date_info_frame.pack(padx=30, pady=(5, 5), fill="x")

        self.date_range_label1 = tk.Label(date_info_frame, text="工人预支申请表：正在读取...",
                                          bg="#f8f9fa", font=("微软雅黑", 10), fg="#e67e22", anchor="w")
        self.date_range_label1.pack(anchor="w", pady=2)
        self.date_range_label2 = tk.Label(date_info_frame, text="预支明细表：正在读取...",
                                          bg="#f8f9fa", font=("微软雅黑", 10), fg="#e67e22", anchor="w")
        self.date_range_label2.pack(anchor="w", pady=2)

        sel_frame = tk.LabelFrame(
            self.step2_frame, text=" 选择筛选时间范围 ",
            font=("微软雅黑", 11, "bold"), bg="#f8f9fa", padx=30, pady=15
        )
        sel_frame.pack(padx=30, pady=(5, 10), fill="x")

        self._step2_entry_start = self._create_date_picker(sel_frame, "开始日期：", "2026-01-01")
        self._step2_entry_end = self._create_date_picker(sel_frame, "结束日期：", datetime.now().strftime("%Y-%m-%d"))

        btn_frame = tk.Frame(self.step2_frame, bg="#f8f9fa")
        btn_frame.pack(pady=10)
        self.process_btn = tk.Button(
            btn_frame, text="▶ 开始对比  → 查看结果",
            command=self._process_both_tables, font=("微软雅黑", 13, "bold"),
            bg="#e74c3c", fg="white", padx=25, pady=8, cursor="hand2"
        )
        self.process_btn.pack()

        self.progress_bar = ttk.Progressbar(self.step2_frame, mode="indeterminate", length=400)
        self.progress_bar.pack(pady=5)

    def _load_both_date_ranges(self):
        min1, max1 = "无有效日期", "无有效日期"
        min2, max2 = "无有效日期", "无有效日期"

        try:
            min1, max1 = self.core.get_date_range_table1()
            if min1 == "无有效日期":
                self.date_range_label1.config(text=f"⚠ 工人预支申请表：未识别到有效日期，请检查日期列格式", fg="#e74c3c")
            else:
                self.date_range_label1.config(text=f"✅ 工人预支申请表：{min1} ~ {max1}", fg="#28a745")
        except Exception as e:
            self.date_range_label1.config(text=f"⚠ 工人预支申请表读取失败: {str(e)}", fg="#e74c3c")

        try:
            min2, max2 = self.core.get_date_range_table2()
            if min2 == "无有效日期":
                self.date_range_label2.config(text=f"⚠ 预支明细表：未识别到有效日期，请检查日期列格式", fg="#e74c3c")
            else:
                self.date_range_label2.config(text=f"✅ 预支明细表：{min2} ~ {max2}", fg="#28a745")
        except Exception as e:
            self.date_range_label2.config(text=f"⚠ 预支明细表读取失败: {str(e)}", fg="#e74c3c")

        def _parse(d):
            try:
                return datetime.strptime(d, "%Y-%m-%d")
            except Exception:
                return None

        dt1_min, dt1_max = _parse(min1), _parse(max1)
        dt2_min, dt2_max = _parse(min2), _parse(max2)

        default_start, default_end = None, None
        if dt1_min and dt2_min:
            default_start = max(dt1_min, dt2_min)
        elif dt1_min:
            default_start = dt1_min
        elif dt2_min:
            default_start = dt2_min

        if dt1_max and dt2_max:
            default_end = min(dt1_max, dt2_max)
        elif dt1_max:
            default_end = dt1_max
        elif dt2_max:
            default_end = dt2_max

        if default_start and default_end and default_start <= default_end:
            self._step2_entry_start.delete(0, tk.END)
            self._step2_entry_start.insert(0, default_start.strftime("%Y-%m-%d"))
            self._step2_entry_end.delete(0, tk.END)
            self._step2_entry_end.insert(0, default_end.strftime("%Y-%m-%d"))

    def _process_both_tables(self):
        start_str = self._step2_entry_start.get().strip()
        end_str = self._step2_entry_end.get().strip()
        if not start_str or not end_str:
            messagebox.showwarning("警告", "请输入开始日期和结束日期")
            return

        try:
            self.core.set_date_range_both(start_str, end_str)
        except Exception:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return

        self.progress_bar.start()
        self.process_btn.config(state="disabled", text="处理中...")
        self.status_label.config(text="正在处理数据，请稍候...")
        try:
            self._get_toplevel().update()
        except Exception:
            pass

        try:
            self.core.process_first_table()
            self.core.process_second_table()
            self.progress_bar.stop()
            self.process_btn.config(state="normal", text="✓ 处理完成")
            self.status_label.config(text="处理完成")
            self.notebook.tab(2, state="normal")
            self.notebook.select(2)
            self.display_results()
        except Exception as e:
            self.progress_bar.stop()
            self.process_btn.config(state="normal", text="▶ 重新对比")
            import traceback; traceback.print_exc()
            messagebox.showerror("处理错误", f"数据处理过程中出现错误:\n{str(e)}")

    def _create_date_picker(self, parent, label_text, default_value):
        frame = tk.Frame(parent, bg="#f8f9fa")
        frame.pack(fill="x", pady=8)
        tk.Label(frame, text=label_text, font=("微软雅黑", 11),
                 bg="#f8f9fa", width=10, anchor="e").pack(side=tk.LEFT)

        entry = tk.Entry(frame, font=("微软雅黑", 11), width=20, relief="solid", bd=1)
        entry.pack(side=tk.LEFT, padx=(10, 5))
        entry.insert(0, default_value)

        tk.Button(frame, text="📅", font=("微软雅黑", 10),
                  command=lambda: CalendarPopup(self._get_toplevel(), lambda d: self._set_date(entry, d)),
                  cursor="hand2", relief="solid", bd=1, width=3).pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(frame, text="格式：YYYY-MM-DD", font=("微软雅黑", 9),
                 bg="#f8f9fa", fg="#6c757d").pack(side=tk.LEFT)
        return entry

    @staticmethod
    def _set_date(entry, date_str):
        entry.delete(0, tk.END)
        entry.insert(0, date_str)

    # ==================== 步骤3：结果预览 ====================

    def setup_step3_ui(self):
        toolbar = tk.Frame(self.step3_frame, bg="#f8f9fa")
        toolbar.pack(fill="x", pady=(10, 5), padx=10)

        self.save_btn = tk.Button(
            toolbar, text="💾 保存结果到Excel",
            command=self.save_results, font=("微软雅黑", 11, "bold"),
            bg="#27ae60", fg="white", padx=15, pady=5, cursor="hand2", state="disabled"
        )
        self.save_btn.pack(side="left", padx=5)

        self.back_btn = tk.Button(
            toolbar, text="◀ 返回修改",
            command=lambda: self.notebook.select(1), font=("微软雅黑", 10),
            bg="#6c757d", fg="white", padx=10, pady=5, cursor="hand2"
        )
        self.back_btn.pack(side="left", padx=5)

        self.result_info_label = tk.Label(toolbar, text="", bg="#f8f9fa",
                                          font=("微软雅黑", 10), fg="#28a745", anchor="w")
        self.result_info_label.pack(side="left", padx=(15, 0), fill="x", expand=True)

        result_frame = tk.Frame(self.step3_frame, bg="#f8f9fa")
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ('姓名', '预支金额', '身份证号', '比较结果', '差额', '姓名(右)', '预支数额(右)', '备注', '运营中心')
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=20)

        col_widths = [100, 100, 180, 90, 80, 100, 100, 160, 160]
        col_labels = ['姓名', '预支金额', '身份证号', '比较结果', '差额', '姓名', '预支数额', '备注', '运营中心']
        for col, width, label in zip(columns, col_widths, col_labels):
            self.result_tree.heading(col, text=label)
            self.result_tree.column(col, width=width, anchor="center")

        v_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.result_tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        self.result_tree.bind("<Double-1>", self.on_result_double_click)

        self.empty_label = tk.Label(
            self.step3_frame, text="请先完成数据处理",
            font=("微软雅黑", 12), fg="#adb5bd", bg="#f8f9fa"
        )
        self.empty_label.pack(pady=30)

    # ==================== 文件选择事件 ====================

    def select_first_file(self):
        file_path = filedialog.askopenfilename(
            title="选择工人预支申请表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        file_name = os.path.basename(file_path)
        self.file1_label.config(text=file_name, fg="#2c3e50")
        self.select_sheet_for_file(file_path, 1)

    def select_second_file(self):
        file_path = filedialog.askopenfilename(
            title="选择预支明细表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        file_name = os.path.basename(file_path)
        self.file2_label.config(text=file_name, fg="#2c3e50")
        self.select_sheet_for_file(file_path, 2)

    def select_sheet_for_file(self, file_path, file_index):
        try:
            sheet_names = get_sheet_names(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件的工作表: {str(e)}")
            return

        if not sheet_names:
            messagebox.showerror("错误", "文件中没有工作表")
            return

        if len(sheet_names) == 1:
            selected_sheet = sheet_names[0]
        else:
            selected_sheet = self.show_sheet_selector(sheet_names, file_index)
            if not selected_sheet:
                return

        if file_index == 1:
            self.core.set_file1(file_path, selected_sheet)
            self.core.set_table1_fixed_layout(
                header_row=0, name_col=8, amount_col=5,
                id_col=18, date_col=13, unit_col=3
            )
            self.sheet1_label.config(text=f"✓ 已选择工作表：{selected_sheet}（固定列: D-运营中心, F-预支, I-姓名, N-日期, S-身份证）")
        else:
            self.core.set_file2(file_path, selected_sheet)
            self.core.set_table2_fixed_layout(header_row=1, name_col=4, amount_col=5, date_col=1)
            self.sheet2_label.config(text=f"✓ 已选择工作表：{selected_sheet}（固定列: B-日期, E-姓名, F-金额, 第2行表头）")

        self.check_files_ready()

    def show_sheet_selector(self, sheet_names, file_index):
        selector = tk.Toplevel(self._get_toplevel())
        selector.title(f"选择工作表 - 表格{file_index}")
        selector.geometry("400x350")
        selector.transient(self._get_toplevel())
        selector.grab_set()

        title_text = f"请选择{'工人预支申请表' if file_index == 1 else '预支明细表'}的工作表"
        tk.Label(selector, text=title_text, font=("微软雅黑", 12, "bold")).pack(pady=15)

        list_frame = tk.Frame(selector)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        listbox = tk.Listbox(list_frame, font=("微软雅黑", 11),
                             yscrollcommand=scrollbar.set, selectmode="single")
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        for name in sheet_names:
            listbox.insert("end", name)
        if sheet_names:
            listbox.selection_set(0)

        result = [None]

        def confirm():
            selection = listbox.curselection()
            if selection:
                result[0] = listbox.get(selection[0])
                selector.destroy()
            else:
                messagebox.showwarning("警告", "请选择一个工作表")

        def cancel():
            selector.destroy()

        btn_frame = tk.Frame(selector)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="确认选择", command=confirm,
                  font=("微软雅黑", 10), bg="#3498db", fg="white", padx=15, pady=3).pack(side="left", padx=5)
        tk.Button(btn_frame, text="取消", command=cancel,
                  font=("微软雅黑", 10), padx=15, pady=3).pack(side="left", padx=5)

        self._get_toplevel().wait_window(selector)
        return result[0]

    def check_files_ready(self):
        if (self.core.file1_path and self.core.sheet1_name and
                self.core.file2_path and self.core.sheet2_name):
            self.next_step_btn.config(state="normal")
            self.status_label.config(text="两个文件已选择完毕，请点击'进入下一步'继续")
            self.update_file_info()

    def update_file_info(self):
        if self.core.file1_path and self.core.sheet1_name and "table1" in self.info_labels:
            self.info_labels["table1"].config(
                text=f"工人预支申请表：{os.path.basename(self.core.file1_path)}  →  工作表：{self.core.sheet1_name}"
            )
        if self.core.file2_path and self.core.sheet2_name and "table2" in self.info_labels:
            self.info_labels["table2"].config(
                text=f"预支明细表：{os.path.basename(self.core.file2_path)}  →  工作表：{self.core.sheet2_name}"
            )

    def go_to_step2(self):
        self.update_file_info()
        self.notebook.tab(1, state="normal")
        self.notebook.select(1)
        self.status_label.config(text="请设置统一的筛选日期范围")
        self.process_btn.config(state="normal", text="▶ 开始对比  → 查看结果")
        self._load_both_date_ranges()

    # ==================== 结果展示与保存 ====================

    def display_results(self):
        try:
            self.core.merge_results()
        except Exception as e:
            messagebox.showerror("合并错误", f"对比合并时出错:\n{str(e)}")
            return

        if self.core.final_df is None or self.core.final_df.empty:
            messagebox.showinfo("提示", "没有可显示的结果数据")
            return

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.result_info_label.config(text=f"共 {len(self.core.final_df)} 条对比记录")

        self.result_tree.tag_configure('false_row', foreground='red')
        self.result_tree.tag_configure('true_row', foreground='green')

        for _, row in self.core.final_df.iterrows():
            values = (
                str(row['姓名']), format_number(row['预支金额']),
                str(row['身份证号']), str(row['比较结果']),
                format_number(row['差额']), str(row['姓名(右)']),
                format_number(row['预支数额(右)']), str(row.get('备注', '')),
                str(row.get('运营中心', ''))
            )
            item_id = self.result_tree.insert("", "end", values=values)
            if row['比较结果'] == 'FALSE':
                self.result_tree.item(item_id, tags=('false_row',))
            else:
                self.result_tree.item(item_id, tags=('true_row',))

        self.empty_label.pack_forget()
        self.save_btn.config(state="normal")
        self.app.status_label.config(text="数据比较 - 结果已生成")

    def on_result_double_click(self, event):
        selection = self.result_tree.selection()
        if not selection:
            return
        item = self.result_tree.item(selection[0])
        values = item['values']
        detail_msg = (
            f"姓名(左)：{values[0]}\n预支金额：{values[1]}\n身份证号：{values[2]}\n"
            f"比较结果：{values[3]}\n差额：{values[4]}\n姓名：{values[5]}\n预支数额：{values[6]}"
        )
        messagebox.showinfo("详细信息", detail_msg)

    def save_results(self):
        if self.core.final_df is None or self.core.final_df.empty:
            messagebox.showwarning("警告", "没有数据可以保存")
            return

        save_path = filedialog.asksaveasfilename(
            title="保存对比结果", defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile="预支对比结果.xlsx"
        )
        if not save_path:
            return

        try:
            self.status_label.config(text="正在保存文件...")
            try:
                self._get_toplevel().update()
            except Exception:
                pass

            save_to_excel(self.core.final_df, save_path)
            self.status_label.config(text=f"结果已保存到: {save_path}")
            messagebox.showinfo("保存成功", f"对比结果已成功保存到:\n{save_path}")
            self.app.status_label.config(text="数据比较 - 已保存")
        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("保存失败", f"保存文件时出错:\n{str(e)}")

