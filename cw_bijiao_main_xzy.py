# -- coding: utf-8 --
# @Time : 2026-05-08
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cw_bijiao_main_xzy.py
# @Software: PyCharm
# @Description: 工人预支申请表 vs 预支明细表 对比工具 - GUI界面模块

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
from datetime import datetime, date, timedelta
import calendar

from cw_bijiao_xzy import WorkerAdvanceComparatorCore, get_sheet_names


# ==================== 日历弹窗组件 ====================
class CalendarPopup:
    """日历弹窗 - 替代手动输入日期"""

    def __init__(self, parent, callback):
        self.parent = parent
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("选择日期")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        # 当前显示的年份/月份
        self.today = date.today()
        self.year = self.today.year
        self.month = self.today.month

        self._build_ui()
        self._draw_calendar()
        # 居中
        self.window.update_idletasks()
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() + 50
        self.window.geometry(f"+{x}+{y}")

    def _build_ui(self):
        # 导航栏
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

        # 星期标题
        weekday_frame = tk.Frame(self.window)
        weekday_frame.pack()
        weekdays = ['日', '一', '二', '三', '四', '五', '六']
        for i, wd in enumerate(weekdays):
            color = "#e74c3c" if i == 0 or i == 6 else "#2c3e50"
            tk.Label(weekday_frame, text=wd, font=("微软雅黑", 9, "bold"),
                     width=4, fg=color).pack(side="left", padx=1)

        # 日期网格容器
        self.cal_frame = tk.Frame(self.window)
        self.cal_frame.pack(pady=3)

        # 底部按钮
        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="今天", command=self._select_today,
                  font=("微软雅黑", 9), bg="#3498db", fg="white", width=8).pack(side="left", padx=5)
        tk.Button(btn_frame, text="关闭", command=self.window.destroy,
                  font=("微软雅黑", 9), width=8).pack(side="left", padx=5)

    def _change_month(self, delta):
        self.month += delta
        if self.month > 12:
            self.month = 1
            self.year += 1
        elif self.month < 1:
            self.month = 12
            self.year -= 1
        self._draw_calendar()

    def _change_year(self, delta):
        self.year += delta
        self._draw_calendar()

    def _draw_calendar(self):
        # 清空
        for w in self.cal_frame.winfo_children():
            w.destroy()

        self.month_label.config(text=f"{self.year}年{self.month:02d}月")

        cal = calendar.monthcalendar(self.year, self.month)
        for week_idx, week in enumerate(cal):
            for day_idx, day in enumerate(week):
                if day == 0:
                    # 空白占位
                    lbl = tk.Label(self.cal_frame, text="", width=4, font=("微软雅黑", 9))
                    lbl.grid(row=week_idx, column=day_idx, padx=1, pady=1)
                else:
                    is_today = (self.year == self.today.year and
                                self.month == self.today.month and
                                day == self.today.day)
                    bg = "#3498db" if is_today else "#f0f0f0"
                    fg = "white" if is_today else "#2c3e50"
                    btn = tk.Button(
                        self.cal_frame, text=str(day), width=4, height=1,
                        font=("微软雅黑", 9), bg=bg, fg=fg,
                        command=lambda d=day: self._select_day(d),
                        relief="flat", cursor="hand2"
                    )
                    btn.grid(row=week_idx, column=day_idx, padx=1, pady=1)

    def _select_day(self, day):
        selected = f"{self.year}-{self.month:02d}-{day:02d}"
        self.callback(selected)
        self.window.destroy()

    def _select_today(self):
        self.callback(self.today.strftime("%Y-%m-%d"))
        self.window.destroy()


class WorkerAdvanceComparatorGUI:
    """工人预支对比工具 - GUI界面类"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("工人预支对比工具")
        self.root.geometry("1000x750")
        self.center_window(self.root, 1000, 750)

        # 核心功能实例
        self.core = WorkerAdvanceComparatorCore()

        # 各步骤中的标签引用（由_build_table_filter_ui填充）
        self.info_labels = {}
        self._date_range_refs = {}

        # 设置UI
        self.setup_ui()

    def center_window(self, window, width, height):
        """设置窗口居中"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        """设置用户界面"""
        # ==================== 标题 ====================
        title_frame = tk.Frame(self.root, bg="#2c3e50")
        title_frame.pack(fill="x")
        tk.Label(
            title_frame, text="工人预支申请表 vs 预支明细表 对比工具",
            font=("微软雅黑", 16, "bold"), fg="white", bg="#2c3e50",
            pady=12
        ).pack()

        # ==================== 主内容区域（使用Notebook分步） ====================
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # ---------- 步骤1：文件选择 ----------
        self.step1_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step1_frame, text="  ① 选择文件  ")
        self.setup_step1_ui()

        # ---------- 步骤2：筛选表1 ----------
        self.step2_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step2_frame, text="  ② 筛选表1  ")
        self.setup_step2_ui()

        # ---------- 步骤3：筛选表2 ----------
        self.step3_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step3_frame, text="  ③ 筛选表2  ")
        self.setup_step3_ui()

        # ---------- 步骤4：结果预览 ----------
        self.step4_frame = tk.Frame(self.notebook, bg="#f8f9fa")
        self.notebook.add(self.step4_frame, text="  ④ 结果预览  ")
        self.setup_step4_ui()

        # ==================== 底部状态栏 ====================
        status_frame = tk.Frame(self.root, bg="#e9ecef")
        status_frame.pack(fill="x", side="bottom")

        self.status_label = tk.Label(
            status_frame, text="请先选择两个Excel文件", bg="#e9ecef",
            font=("微软雅黑", 10), anchor="w", padx=10, pady=5
        )
        self.status_label.pack(fill="x")

        # 默认禁用步骤2、3、4
        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")
        self.notebook.tab(3, state="disabled")

    # ==================== 步骤1：文件选择 ====================
    def setup_step1_ui(self):
        """步骤1界面：文件选择"""
        canvas = tk.Canvas(self.step1_frame, bg="#f8f9fa", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.step1_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f8f9fa")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ---------- 第一个文件选择 ----------
        file1_box = tk.LabelFrame(
            scrollable_frame, text=" 第一个表格：工人预支申请表 ",
            font=("微软雅黑", 12, "bold"), bg="#f8f9fa",
            padx=15, pady=10
        )
        file1_box.pack(fill="x", padx=20, pady=(20, 10))

        desc_label1 = tk.Label(
            file1_box,
            text="选择包含工人预支申请数据的Excel文件，需包含：姓名、预支金额、身份证号、提交时间",
            font=("微软雅黑", 9), bg="#f8f9fa", fg="#6c757d", wraplength=700, justify="left"
        )
        desc_label1.pack(anchor="w", pady=(0, 8))

        btn_frame1 = tk.Frame(file1_box, bg="#f8f9fa")
        btn_frame1.pack(fill="x")

        self.select_btn1 = tk.Button(
            btn_frame1, text="📂 选择工人预支申请表",
            command=self.select_first_file,
            font=("微软雅黑", 11), bg="#3498db", fg="white",
            padx=15, pady=5, cursor="hand2"
        )
        self.select_btn1.pack(side="left")

        self.file1_label = tk.Label(
            btn_frame1, text="未选择文件", bg="#f8f9fa",
            font=("微软雅黑", 10), fg="#6c757d", anchor="w"
        )
        self.file1_label.pack(side="left", padx=(15, 0), fill="x", expand=True)

        self.sheet1_label = tk.Label(
            file1_box, text="", bg="#f8f9fa",
            font=("微软雅黑", 10), fg="#28a745", anchor="w"
        )
        self.sheet1_label.pack(anchor="w", pady=(5, 0))

        # ---------- 分隔 ----------
        tk.Frame(scrollable_frame, height=2, bg="#dee2e6").pack(fill="x", padx=20, pady=10)

        # ---------- 第二个文件选择 ----------
        file2_box = tk.LabelFrame(
            scrollable_frame, text=" 第二个表格：预支明细表 ",
            font=("微软雅黑", 12, "bold"), bg="#f8f9fa",
            padx=15, pady=10
        )
        file2_box.pack(fill="x", padx=20, pady=(10, 10))

        desc_label2 = tk.Label(
            file2_box,
            text="选择包含预支明细数据的Excel文件（第二行为表头），需包含：姓名、预支数额、预支日期",
            font=("微软雅黑", 9), bg="#f8f9fa", fg="#6c757d", wraplength=700, justify="left"
        )
        desc_label2.pack(anchor="w", pady=(0, 8))

        btn_frame2 = tk.Frame(file2_box, bg="#f8f9fa")
        btn_frame2.pack(fill="x")

        self.select_btn2 = tk.Button(
            btn_frame2, text="📂 选择预支明细表",
            command=self.select_second_file,
            font=("微软雅黑", 11), bg="#e67e22", fg="white",
            padx=15, pady=5, cursor="hand2"
        )
        self.select_btn2.pack(side="left")

        self.file2_label = tk.Label(
            btn_frame2, text="未选择文件", bg="#f8f9fa",
            font=("微软雅黑", 10), fg="#6c757d", anchor="w"
        )
        self.file2_label.pack(side="left", padx=(15, 0), fill="x", expand=True)

        self.sheet2_label = tk.Label(
            file2_box, text="", bg="#f8f9fa",
            font=("微软雅黑", 10), fg="#28a745", anchor="w"
        )
        self.sheet2_label.pack(anchor="w", pady=(5, 0))

        # ---------- 进入下一步按钮 ----------
        tk.Frame(scrollable_frame, height=10, bg="#f8f9fa").pack()
        self.next_step_btn = tk.Button(
            scrollable_frame, text="▶ 进入下一步 - 设置日期范围",
            command=self.go_to_step2,
            font=("微软雅黑", 12, "bold"), bg="#28a745", fg="white",
            padx=20, pady=8, cursor="hand2", state="disabled"
        )
        self.next_step_btn.pack(pady=15)

    # ==================== 步骤2：筛选表1（工人预支申请表） ====================
    def _build_table_filter_ui(self, parent_frame, table_label, table_info_key,
                                date_column_name, date_range_method, process_method,
                                set_date_method, next_step_index, next_step_text):
        """
        通用的单表日期筛选UI构建器
        parent_frame: 父Frame
        table_label: 表格名称标签
        table_info_key: 用于info_text的标识
        date_column_name: 日期列名
        date_range_method: core.get_date_range_table1/2
        process_method: core.process_first_table/process_second_table
        set_date_method: core.set_date_range1/2
        next_step_index: 处理后跳转到的tab索引
        next_step_text: 处理按钮文本
        """
        # 标题
        info_frame = tk.Frame(parent_frame, bg="#f8f9fa")
        info_frame.pack(fill="x", padx=30, pady=(25, 5))

        tk.Label(
            info_frame,
            text=f"筛选{table_label}",
            font=("微软雅黑", 15, "bold"), bg="#f8f9fa", fg="#2c3e50"
        ).pack(anchor="w")

        # 文件信息
        file_info_frame = tk.LabelFrame(
            parent_frame, text=" 文件信息 ",
            font=("微软雅黑", 11, "bold"), bg="#f8f9fa",
            padx=15, pady=8
        )
        file_info_frame.pack(padx=30, pady=(10, 5), fill="x")
        self.info_labels[table_info_key] = tk.Label(
            file_info_frame, text="未选择文件", bg="#f8f9fa",
            font=("微软雅黑", 10), anchor="w"
        )
        self.info_labels[table_info_key].pack(anchor="w", pady=2)

        # 日期范围信息
        date_info_frame = tk.LabelFrame(
            parent_frame, text=" 数据中的日期范围 ",
            font=("微软雅黑", 11, "bold"), bg="#f8f9fa",
            padx=15, pady=8
        )
        date_info_frame.pack(padx=30, pady=(5, 5), fill="x")

        date_range_label = tk.Label(
            date_info_frame, text="正在读取数据...", bg="#f8f9fa",
            font=("微软雅黑", 10), fg="#e67e22", anchor="w"
        )
        date_range_label.pack(anchor="w", pady=2)

        # 选择日期范围
        sel_frame = tk.LabelFrame(
            parent_frame, text=" 选择筛选时间范围 ",
            font=("微软雅黑", 11, "bold"), bg="#f8f9fa",
            padx=30, pady=15
        )
        sel_frame.pack(padx=30, pady=(5, 10), fill="x")

        entry_start = self._create_date_picker(sel_frame, "开始日期：", "2026-01-01")
        entry_end = self._create_date_picker(sel_frame, "结束日期：", datetime.now().strftime("%Y-%m-%d"))

        # 处理按钮
        btn_frame = tk.Frame(parent_frame, bg="#f8f9fa")
        btn_frame.pack(pady=10)

        process_btn = tk.Button(
            btn_frame, text=next_step_text,
            command=lambda: self._process_single_table(
                date_range_method, process_method, set_date_method,
                entry_start, entry_end, next_step_index, process_btn, progress_bar
            ),
            font=("微软雅黑", 13, "bold"), bg="#e74c3c", fg="white",
            padx=25, pady=8, cursor="hand2"
        )
        process_btn.pack()

        progress_bar = ttk.Progressbar(parent_frame, mode="indeterminate", length=400)
        progress_bar.pack(pady=5)

        # 存储引用，供后续进入该步骤时触发加载
        self._date_range_refs[table_info_key] = {
            'method': date_range_method,
            'label': date_range_label,
            'entry_start': entry_start,
            'entry_end': entry_end,
        }

    def _load_date_range_info(self, date_range_method, label, entry_start, entry_end):
        """加载表格的可用日期范围并更新UI"""
        try:
            min_date, max_date = date_range_method()
            if min_date == "无有效日期":
                label.config(text="⚠ 未识别到有效日期，请检查日期列格式", fg="#e74c3c")
            else:
                label.config(
                    text=f"✅ 数据范围：{min_date}  ~  {max_date}",
                    fg="#28a745"
                )
                entry_start.delete(0, tk.END)
                entry_start.insert(0, min_date)
                entry_end.delete(0, tk.END)
                entry_end.insert(0, max_date)
        except Exception as e:
            label.config(text=f"⚠ 读取失败: {str(e)}", fg="#e74c3c")

    def _process_single_table(self, date_range_method, process_method,
                               set_date_method, entry_start, entry_end,
                               next_index, btn, progress_bar):
        """处理单个表格的筛选"""
        start_str = entry_start.get().strip()
        end_str = entry_end.get().strip()
        if not start_str or not end_str:
            messagebox.showwarning("警告", "请输入开始日期和结束日期")
            return

        try:
            set_date_method(start_str, end_str)
        except Exception:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return

        progress_bar.start()
        btn.config(state="disabled", text="处理中...")
        self.status_label.config(text="正在处理数据，请稍候...")
        self.root.update()

        try:
            process_method()
            progress_bar.stop()
            btn.config(state="normal", text="✓ 处理完成")
            self.status_label.config(text="处理完成")
            self.notebook.select(next_index)
            self.notebook.tab(next_index, state="normal")
            # 如果跳到步骤3（筛选表2），自动加载日期范围
            if next_index == 2 and "table2" in self._date_range_refs:
                ref = self._date_range_refs["table2"]
                self._load_date_range_info(ref['method'], ref['label'],
                                           ref['entry_start'], ref['entry_end'])
            # 如果跳到步骤4（结果预览），自动合并并显示
            if next_index == 3:
                self.display_results()
        except Exception as e:
            progress_bar.stop()
            btn.config(state="normal", text="▶ 重新处理")
            import traceback
            traceback.print_exc()
            messagebox.showerror("处理错误", f"数据处理过程中出现错误:\n{str(e)}")

    def setup_step2_ui(self):
        """步骤2界面：筛选工人预支申请表"""
        self._build_table_filter_ui(
            parent_frame=self.step2_frame,
            table_label="工人预支申请表",
            table_info_key="table1",
            date_column_name="提交时间",
            date_range_method=self.core.get_date_range_table1,
            process_method=self.core.process_first_table,
            set_date_method=self.core.set_date_range1,
            next_step_index=2,
            next_step_text="▶ 处理表1  →"
        )

    # ==================== 步骤3：筛选表2（预支明细表） ====================
    def setup_step3_ui(self):
        """步骤3界面：筛选预支明细表"""
        self._build_table_filter_ui(
            parent_frame=self.step3_frame,
            table_label="预支明细表",
            table_info_key="table2",
            date_column_name="预支日期",
            date_range_method=self.core.get_date_range_table2,
            process_method=self.core.process_second_table,
            set_date_method=self.core.set_date_range2,
            next_step_index=3,
            next_step_text="▶ 处理表2  → 查看结果"
        )

    def _create_date_picker(self, parent, label_text, default_value):
        """创建一个带日历按钮的日期选择行"""
        frame = tk.Frame(parent, bg="#f8f9fa")
        frame.pack(fill="x", pady=8)

        tk.Label(
            frame, text=label_text, font=("微软雅黑", 11),
            bg="#f8f9fa", width=10, anchor="e"
        ).pack(side="left")

        entry = tk.Entry(
            frame, font=("微软雅黑", 11), width=20,
            relief="solid", bd=1
        )
        entry.pack(side="left", padx=(10, 5))
        entry.insert(0, default_value)

        tk.Button(
            frame, text="📅", font=("微软雅黑", 10),
            command=lambda: CalendarPopup(self.root, lambda d: self._set_date(entry, d)),
            cursor="hand2", relief="solid", bd=1, width=3
        ).pack(side="left", padx=(0, 5))

        tk.Label(
            frame, text="格式：YYYY-MM-DD", font=("微软雅黑", 9),
            bg="#f8f9fa", fg="#6c757d"
        ).pack(side="left")

        return entry

    @staticmethod
    def _set_date(entry, date_str):
        """设置日期到 Entry"""
        entry.delete(0, tk.END)
        entry.insert(0, date_str)

    # ==================== 步骤4：结果预览 ====================
    def setup_step4_ui(self):
        """步骤4界面：结果预览与保存"""
        # 工具栏
        toolbar = tk.Frame(self.step4_frame, bg="#f8f9fa")
        toolbar.pack(fill="x", pady=(10, 5), padx=10)

        self.save_btn = tk.Button(
            toolbar, text="💾 保存结果到Excel",
            command=self.save_results,
            font=("微软雅黑", 11, "bold"), bg="#27ae60", fg="white",
            padx=15, pady=5, cursor="hand2", state="disabled"
        )
        self.save_btn.pack(side="left", padx=5)

        self.back_btn = tk.Button(
            toolbar, text="◀ 返回修改",
            command=lambda: self.notebook.select(2),
            font=("微软雅黑", 10), bg="#6c757d", fg="white",
            padx=10, pady=5, cursor="hand2"
        )
        self.back_btn.pack(side="left", padx=5)

        self.result_info_label = tk.Label(
            toolbar, text="", bg="#f8f9fa",
            font=("微软雅黑", 10), fg="#28a745", anchor="w"
        )
        self.result_info_label.pack(side="left", padx=(15, 0), fill="x", expand=True)

        # ---------- 结果表格 ----------
        result_frame = tk.Frame(self.step4_frame, bg="#f8f9fa")
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ('姓名', '预支金额', '身份证号', '比较结果', '差额', '姓名(右)', '预支数额(右)')
        self.result_tree = ttk.Treeview(
            result_frame, columns=columns, show='headings', height=20
        )

        col_widths = [100, 100, 180, 90, 80, 100, 100]
        for col, width in zip(columns, col_widths):
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=width, anchor="center")

        v_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set
        )

        self.result_tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        self.result_tree.bind("<Double-1>", self.on_result_double_click)

        self.empty_label = tk.Label(
            self.step4_frame, text="请先完成数据处理",
            font=("微软雅黑", 12), fg="#adb5bd", bg="#f8f9fa"
        )
        self.empty_label.pack(pady=30)

    # ==================== 文件选择事件 ====================
    def select_first_file(self):
        """选择第一个表格（工人预支申请表）"""
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
        """选择第二个表格（预支明细表）"""
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
        """为指定文件选择工作表"""
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
            self.sheet1_label.config(text=f"✓ 已选择工作表：{selected_sheet}")
        else:
            self.core.set_file2(file_path, selected_sheet)
            self.sheet2_label.config(text=f"✓ 已选择工作表：{selected_sheet}")

        self.check_files_ready()

    def show_sheet_selector(self, sheet_names, file_index):
        """显示工作表选择器弹窗"""
        selector = tk.Toplevel(self.root)
        selector.title(f"选择工作表 - 表格{file_index}")
        selector.geometry("400x350")
        selector.transient(self.root)
        selector.grab_set()
        self.center_window(selector, 400, 350)

        title_text = f"请选择{'工人预支申请表' if file_index == 1 else '预支明细表'}的工作表"
        tk.Label(
            selector, text=title_text,
            font=("微软雅黑", 12, "bold")
        ).pack(pady=15)

        list_frame = tk.Frame(selector)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        listbox = tk.Listbox(
            list_frame, font=("微软雅黑", 11),
            yscrollcommand=scrollbar.set, selectmode="single"
        )
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
        tk.Button(
            btn_frame, text="确认选择", command=confirm,
            font=("微软雅黑", 10), bg="#3498db", fg="white", padx=15, pady=3
        ).pack(side="left", padx=5)
        tk.Button(
            btn_frame, text="取消", command=cancel,
            font=("微软雅黑", 10), padx=15, pady=3
        ).pack(side="left", padx=5)

        self.root.wait_window(selector)
        return result[0]

    def check_files_ready(self):
        """检查两个文件是否都已选择完毕"""
        if self.core.file1_path and self.core.sheet1_name and self.core.file2_path and self.core.sheet2_name:
            self.next_step_btn.config(state="normal")
            self.status_label.config(text="两个文件已选择完毕，请点击'进入下一步'继续")
            self.update_file_info()

    def update_file_info(self):
        """更新文件信息显示"""
        if self.core.file1_path and self.core.sheet1_name and "table1" in self.info_labels:
            self.info_labels["table1"].config(
                text=f"工人预支申请表：{os.path.basename(self.core.file1_path)}  →  工作表：{self.core.sheet1_name}"
            )
        if self.core.file2_path and self.core.sheet2_name and "table2" in self.info_labels:
            self.info_labels["table2"].config(
                text=f"预支明细表：{os.path.basename(self.core.file2_path)}  →  工作表：{self.core.sheet2_name}"
            )

    def go_to_step2(self):
        """进入步骤2：筛选表1"""
        self.update_file_info()
        self.notebook.select(1)
        self.notebook.tab(1, state="normal")
        self.status_label.config(text="请为工人预支申请表设置筛选日期范围")
        # 进入步骤后加载日期范围信息
        if "table1" in self._date_range_refs:
            ref = self._date_range_refs["table1"]
            self._load_date_range_info(ref['method'], ref['label'],
                                       ref['entry_start'], ref['entry_end'])

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

        for _, row in self.core.final_df.iterrows():
            values = (
                str(row['姓名']),
                self.core.format_number(row['预支金额']),
                str(row['身份证号']),
                str(row['比较结果']),
                self.core.format_number(row['差额']),
                str(row['姓名(右)']),
                self.core.format_number(row['预支数额(右)'])
            )
            item_id = self.result_tree.insert("", "end", values=values)

            if row['比较结果'] == 'FALSE':
                self.result_tree.tag_configure('false_row', foreground='red')
                self.result_tree.item(item_id, tags=('false_row',))
            else:
                self.result_tree.tag_configure('true_row', foreground='green')
                self.result_tree.item(item_id, tags=('true_row',))

        self.empty_label.pack_forget()
        self.save_btn.config(state="normal")

    def on_result_double_click(self, event):
        """双击结果行查看详情"""
        selection = self.result_tree.selection()
        if not selection:
            return

        item = self.result_tree.item(selection[0])
        values = item['values']

        detail_msg = (
            f"姓名(左)：{values[0]}\n"
            f"预支金额：{values[1]}\n"
            f"身份证号：{values[2]}\n"
            f"比较结果：{values[3]}\n"
            f"差额：{values[4]}\n"
            f"姓名(右)：{values[5]}\n"
            f"预支数额(右)：{values[6]}"
        )
        messagebox.showinfo("详细信息", detail_msg)

    # ==================== 保存结果 ====================
    def save_results(self):
        """保存结果到Excel文件"""
        if self.core.final_df is None or self.core.final_df.empty:
            messagebox.showwarning("警告", "没有数据可以保存")
            return

        save_path = filedialog.asksaveasfilename(
            title="保存对比结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile="预支对比结果.xlsx"
        )

        if not save_path:
            return

        try:
            self.status_label.config(text="正在保存文件...")
            self.root.update()

            self.core.save_to_excel(self.core.final_df, save_path)

            self.status_label.config(text=f"结果已保存到: {save_path}")
            messagebox.showinfo("保存成功", f"对比结果已成功保存到:\n{save_path}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("保存失败", f"保存文件时出错:\n{str(e)}")

    # ==================== 运行 ====================
    def run(self):
        """运行应用程序"""
        self.root.mainloop()


def main():
    """主函数"""
    app = WorkerAdvanceComparatorGUI()
    app.run()


if __name__ == "__main__":
    main()
