import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from finance.pmh.yzpz_core import excel_processor
import os


class ExcelProcessorGUI:
    """预支平账处理工具 GUI（完整版）"""

    def __init__(self, root):
        self.root = root
        self.root.title("预支平账处理工具")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)

        # 初始化变量
        self.advance_file_path = None
        self.sheet_names = []
        self.check_file_path = None
        self.check_sheet_name = None
        self.check_workbook = None
        self.selected_unit = None
        self.selected_month = None
        self.available_units = []
        self.last_save_path = None

        # 预支明细表
        self.advance_sheet_var = tk.StringVar()
        self.advance_sheet_combo = None
        self.advance_sheets = []

        # 工资表
        self.salary_file_path = None
        self.salary_sheet_var = tk.StringVar()
        self.salary_sheet_combo = None
        self.salary_sheets = []
        self.salary_workbook = None

        self.create_widgets()
        self.center_window()
        self.root.update()

    def center_window(self):
        """将窗口设置在屏幕中央"""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 900
        window_height = 650
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def create_widgets(self):
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(2, weight=1)

        # ==================== 文件操作区域 ====================
        file_frame = ttk.LabelFrame(main_container, text="文件操作")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        file_frame.columnconfigure(1, weight=1)
        file_frame.columnconfigure(4, weight=1)

        # 预支明细表行
        tk.Label(file_frame, text="预支明细表：").grid(row=0, column=0, sticky=tk.W, padx=8, pady=6)
        self.advance_entry = tk.Entry(file_frame, state="readonly", readonlybackground="white")
        self.advance_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)
        self.advance_entry_var = tk.StringVar(value="尚未选择文件")
        self.advance_entry.config(textvariable=self.advance_entry_var)
        ttk.Button(file_frame, text="浏览", command=self.browse_advance_file).grid(row=0, column=2, padx=(0, 8), pady=6)
        tk.Label(file_frame, text="工作簿：").grid(row=0, column=3, sticky=tk.W, padx=(8, 4), pady=6)
        self.advance_sheet_combo = ttk.Combobox(file_frame, textvariable=self.advance_sheet_var, state="readonly")
        self.advance_sheet_combo.grid(row=0, column=4, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)
        self.advance_sheet_combo.bind("<<ComboboxSelected>>", self.on_advance_sheet_selected)

        # 工资表行
        tk.Label(file_frame, text="工资表：").grid(row=1, column=0, sticky=tk.W, padx=8, pady=6)
        self.salary_entry = tk.Entry(file_frame, state="readonly", readonlybackground="white")
        self.salary_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)
        self.salary_entry_var = tk.StringVar(value="尚未选择文件")
        self.salary_entry.config(textvariable=self.salary_entry_var)
        ttk.Button(file_frame, text="浏览", command=self.browse_salary_file).grid(row=1, column=2, padx=(0, 8), pady=6)
        tk.Label(file_frame, text="工作簿：").grid(row=1, column=3, sticky=tk.W, padx=(8, 4), pady=6)
        self.salary_sheet_combo = ttk.Combobox(file_frame, textvariable=self.salary_sheet_var, state="readonly")
        self.salary_sheet_combo.grid(row=1, column=4, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)
        self.salary_sheet_combo.bind("<<ComboboxSelected>>", self.on_salary_sheet_selected)

        # 第三行：单位 + 平账时间
        tk.Label(file_frame, text="单位：").grid(row=2, column=0, sticky=tk.W, padx=8, pady=6)
        self.unit_var = tk.StringVar()
        self.unit_combo = ttk.Combobox(file_frame, textvariable=self.unit_var, state="disabled")
        self.unit_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)
        self.unit_combo.bind("<<ComboboxSelected>>", self.on_unit_selected)
        self.unit_combo.bind("<KeyRelease>", self._filter_units)
        self.unit_combo.bind("<FocusOut>", self.on_unit_focus_out)

        tk.Label(file_frame, text="平账时间：").grid(row=2, column=3, sticky=tk.W, padx=(8, 4), pady=6)
        self.month_var = tk.StringVar()
        self.month_combo = ttk.Combobox(file_frame, textvariable=self.month_var, state="disabled", width=10)
        self.month_combo.grid(row=2, column=4, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)
        self.month_combo["values"] = [f"{i}月" for i in range(1, 13)]
        self.month_combo.bind("<<ComboboxSelected>>", self.on_month_selected)
        self.month_combo.bind("<FocusOut>", self.on_month_focus_out)

        # ==================== 功能操作区域 ====================
        self.func_frame = ttk.LabelFrame(main_container, text="功能操作")
        self.func_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 8))

        self.func_btn_container = tk.Frame(self.func_frame)
        self.func_btn_container.pack(pady=10)

        self.btn_check_balance = ttk.Button(self.func_btn_container, text="核对 - 平账", command=self.check_and_balance)
        self.btn_check_balance.pack(side=tk.LEFT, padx=8)
        self.btn_save = ttk.Button(self.func_btn_container, text="选择保存位置", command=self.select_save_location)
        self.btn_save.pack(side=tk.LEFT, padx=8)
        self.btn_open_dir = ttk.Button(self.func_btn_container, text="打开保存目录", command=self.open_save_directory)
        self.btn_open_dir.pack(side=tk.LEFT, padx=8)
        self.btn_reset = ttk.Button(self.func_btn_container, text="重置", command=self.reset_all)
        self.btn_reset.pack(side=tk.LEFT, padx=8)

        # ==================== 状态信息区域 ====================
        status_frame = ttk.LabelFrame(main_container, text="状态信息")
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)

        text_container = tk.Frame(status_frame)
        text_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=8, pady=6)
        text_container.columnconfigure(0, weight=1)
        text_container.rowconfigure(0, weight=1)

        self.status_text = tk.Text(
            text_container,
            height=10,
            wrap=tk.WORD,
            state=tk.DISABLED,
        )
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        status_scrollbar = ttk.Scrollbar(text_container, orient="vertical", command=self.status_text.yview)
        status_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.status_text.config(yscrollcommand=status_scrollbar.set)

        self.status_text.tag_config("success", foreground="green")
        self.status_text.tag_config("error", foreground="red")
        self.status_text.tag_config("warning", foreground="orange")
        self.status_text.tag_config("header", foreground="blue", font=("", 10, "bold"))
        self.status_text.tag_config("normal", foreground="black")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    # ==================== 日志与对话框 ====================

    def log_status(self, message):
        """记录状态信息，带颜色区分"""
        self.status_text.config(state=tk.NORMAL)

        if message.startswith("✓") or message.startswith("✔"):
            self.status_text.insert(tk.END, message + "\n", "success")
        elif message.startswith("✗") or message.startswith("✘") or message.startswith("错误"):
            self.status_text.insert(tk.END, message + "\n", "error")
        elif message.startswith("⚠") or message.startswith("警告"):
            self.status_text.insert(tk.END, message + "\n", "warning")
        elif "开始" in message and "=====" in message:
            self.status_text.insert(tk.END, message + "\n", "header")
        else:
            self.status_text.insert(tk.END, message + "\n", "normal")

        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update()

    def center_dialog(self, dialog, width, height):
        """将对话框居中显示"""
        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()

        x = root_x + (root_width - width) // 2
        y = root_y + (root_height - height) // 2

        dialog.geometry(f"{width}x{height}+{x}+{y}")

    # ==================== 第三行逻辑 ====================

    def _load_units(self):
        """从预支明细表已选工作簿的第2行提取单位列表，加载到单位下拉框"""
        if not self.advance_file_path or not self.advance_sheet_var.get():
            return

        try:
            wb = load_workbook(self.advance_file_path, data_only=True)
            sheet = wb[self.advance_sheet_var.get()]

            # 查找单位列（在预支明细表工作簿的第2行查找标题）
            unit_col = None
            for col in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=2, column=col).value
                if cell_value and str(cell_value).strip() == "单位":
                    unit_col = col
                    break

            if unit_col:
                units = set()
                for row in range(3, sheet.max_row + 1):
                    val = sheet.cell(row=row, column=unit_col).value
                    if val and str(val).strip():
                        units.add(str(val).strip())
                self.available_units = sorted(units)
                self.unit_combo["values"] = self.available_units
                self.unit_combo.config(state="normal")
                self.log_status(f"✓ 已加载单位列表，共 {len(self.available_units)} 个单位")
            else:
                self.available_units = []
                self.unit_combo["values"] = []
                self.unit_combo.config(state="disabled")
                self.log_status("✗ 未找到'单位'列，请检查预支明细表第2行是否有'单位'列")

            wb.close()
        except Exception as e:
            self.log_status(f"✗ 提取单位列表失败：{str(e)}")
            self.available_units = []
            self.unit_combo["values"] = []
            self.unit_combo.config(state="disabled")

    def _enable_month(self):
        """启用平账时间下拉框（在工资表选择完成后调用）"""
        self.month_combo.config(state="normal")
        self.log_status("✓ 平账时间下拉框已启用")

    def on_unit_selected(self, event):
        """单位下拉框选择事件"""
        self.selected_unit = self.unit_var.get()
        self.log_status(f"已选择单位：{self.selected_unit}")

    def on_month_selected(self, event):
        """平账时间下拉框选择事件"""
        self.selected_month = self.month_var.get()
        self.log_status(f"已选择平账时间：{self.selected_month}")

    def _filter_units(self, event):
        """单位输入框模糊过滤"""
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab'):
            return

        typed = self.unit_var.get()
        if not typed:
            self.unit_combo["values"] = self.available_units
            return

        filtered = [u for u in self.available_units if typed.lower() in u.lower()]
        self.unit_combo["values"] = filtered

        if filtered:
            self.unit_combo.event_generate('<Down>')

    def on_unit_focus_out(self, event):
        """单位输入框失去焦点时同步选择值"""
        self.selected_unit = self.unit_var.get().strip()

    def on_month_focus_out(self, event):
        """平账时间输入框失去焦点时同步选择值"""
        self.selected_month = self.month_var.get().strip()

    # ==================== 文件操作 ====================

    def browse_advance_file(self):
        """选择预支明细表（主文件）"""
        file_path = filedialog.askopenfilename(
            title="选择预支明细表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        self.advance_file_path = file_path
        self.advance_entry_var.set(os.path.basename(file_path))

        success, message, sheet_names = excel_processor.select_excel_file(file_path)
        if success:
            self.sheet_names = sheet_names
            self.advance_sheets = sheet_names
            self.advance_sheet_combo["values"] = sheet_names
            self.advance_sheet_var.set("")
            self.log_status(f"✓ {message}")
        else:
            self.log_status(f"✗ {message}")
            self.advance_sheets = []
            self.advance_sheet_combo["values"] = []
            self.advance_sheet_var.set("")

        # 预支明细表选择完成后，立即加载单位列表
        self._load_units()

    def browse_salary_file(self):
        """选择工资表（外部数据源）"""
        file_path = filedialog.askopenfilename(
            title="选择工资表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        if self.salary_workbook:
            try:
                self.salary_workbook.close()
            except:
                pass

        self.salary_file_path = file_path
        self.salary_entry_var.set(os.path.basename(file_path))

        try:
            self.salary_workbook = load_workbook(file_path, data_only=True)
            self.salary_sheets = self.salary_workbook.sheetnames
            self.salary_sheet_combo["values"] = self.salary_sheets
            self.salary_sheet_var.set("")
            self.log_status(f"✓ 成功读取工资表文件：{os.path.basename(file_path)}")
        except Exception as e:
            self.log_status(f"✗ 读取工资表文件失败：{str(e)}")
            if self.salary_workbook:
                try:
                    self.salary_workbook.close()
                except:
                    pass
            self.salary_workbook = None
            self.salary_sheets = []
            self.salary_sheet_combo["values"] = []
            self.salary_sheet_var.set("")
            # 工资表读取失败，仅禁用平账时间下拉框，不影响单位下拉框
            self.month_combo.config(state="disabled")
            self.selected_month = None
            self.month_var.set("")
            return

        # 工资表选择完成后，启用平账时间下拉框
        self._enable_month()

    def on_advance_sheet_selected(self, event):
        """预支明细表工作簿选择事件"""
        sheet_name = self.advance_sheet_var.get()
        if sheet_name:
            self.log_status(f"已选择预支明细表工作簿：{sheet_name}")
            # 预支明细表工作簿切换后，重新提取单位列表
            self._load_units()

    def on_salary_sheet_selected(self, event):
        """工资表工作簿选择事件"""
        sheet_name = self.salary_sheet_var.get()
        if sheet_name and self.salary_workbook:
            self.check_sheet_name = sheet_name
            self.check_workbook = self.salary_workbook[sheet_name]
            self.log_status(f"已选择工资表工作簿：{sheet_name}")
            # 工资表工作簿选择后，启用平账时间下拉框
            self._enable_month()
