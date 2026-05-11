import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cw_yzpz_pmh as excel_processor
import os


class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("预支平账处理工具")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        # 初始化变量
        self.current_file_path = None
        self.sheet_names = []
        self.check_file_path = None
        self.check_sheet_name = None
        self.check_workbook = None

        # 创建GUI组件
        self.create_widgets()

        # 窗口居中
        self.center_window()

    def center_window(self):
        """将窗口设置在屏幕中央"""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 800
        window_height = 600
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def create_widgets(self):
        # 主容器
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(2, weight=1)

        # ==================== 文件操作区域 ====================
        file_frame = ttk.LabelFrame(main_container, text="文件操作")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        file_frame.columnconfigure(1, weight=1)

        tk.Label(file_frame, text="当前文件：").grid(row=0, column=0, sticky=tk.W, padx=8, pady=6)
        self.file_path_var = tk.StringVar(value="尚未选择文件")
        self.file_path_label = tk.Label(file_frame, textvariable=self.file_path_var, anchor=tk.W)
        self.file_path_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 8), pady=6)

        btn_frame = tk.Frame(file_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=8, pady=(0, 6))

        ttk.Button(btn_frame, text="选择 Excel 文件", command=self.select_excel_file).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="选择保存位置", command=self.select_save_location).pack(side=tk.LEFT)

        # ==================== 功能操作区域 ====================
        func_frame = ttk.LabelFrame(main_container, text="功能操作")
        func_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        func_frame.columnconfigure(0, weight=1)
        func_frame.columnconfigure(1, weight=1)

        # 左侧：核对
        check_card = tk.Frame(func_frame)
        check_card.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=6, pady=6)
        check_card.columnconfigure(0, weight=1)

        tk.Label(check_card, text="核对", font=("", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 4))
        tk.Label(check_card, text="从外部核对文件提取姓名与扣款数据，\n匹配到当前工作簿的对应 Sheet 中。", justify=tk.LEFT).grid(row=1, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Button(check_card, text="开始核对", command=self.check_data).grid(row=2, column=0, sticky=tk.W)

        # 右侧：平账
        balance_card = tk.Frame(func_frame)
        balance_card.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=6, pady=6)
        balance_card.columnconfigure(0, weight=1)

        tk.Label(balance_card, text="平账", font=("", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 4))
        tk.Label(balance_card, text="根据已核对的扣款数据，\n在工资表或预支表中进行平账处理。", justify=tk.LEFT).grid(row=1, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Button(balance_card, text="开始平账", command=self.balance_accounts).grid(row=2, column=0, sticky=tk.W)

        # 操作提示
        hint_frame = tk.Frame(func_frame)
        hint_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=6, pady=(0, 6))
        tk.Label(hint_frame, text="操作顺序：选择文件 → 核对 → 平账 → 保存。核对和平账支持多轮次，每轮自动使用新的 Sheet 组。").pack(anchor=tk.W)

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

        # 配置状态文本颜色标签
        self.status_text.tag_config("success", foreground="green")
        self.status_text.tag_config("error", foreground="red")
        self.status_text.tag_config("warning", foreground="orange")
        self.status_text.tag_config("header", foreground="blue", font=("", 10, "bold"))
        self.status_text.tag_config("normal", foreground="black")

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

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

    def select_excel_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )

        if file_path:
            self.current_file_path = file_path
            self.file_path_var.set(f"{os.path.basename(file_path)}")

            success, message, sheet_names = excel_processor.excel_processor.select_excel_file(file_path)

            if success:
                self.sheet_names = sheet_names
                self.log_status(f"✓ {message}")
                self.log_status(f"提示：请按顺序执行 核对 -> 平账 -> 保存")
            else:
                self.log_status(f"✗ {message}")

    def select_save_location(self):
        """选择保存位置并保存文件"""
        if not self.current_file_path:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("选择保存方式")
        dialog.geometry("320x140")
        dialog.transient(self.root)
        dialog.grab_set()

        self.center_dialog(dialog, 320, 140)

        tk.Label(dialog, text="请选择保存方式：").pack(pady=(16, 12))

        save_path = [None]

        def save_to_original():
            save_path[0] = self.current_file_path
            dialog.destroy()

        def save_as_new():
            path = filedialog.asksaveasfilename(
                title="选择保存位置",
                defaultextension=".xlsx",
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            if path:
                save_path[0] = path
            dialog.destroy()

        def cancel_save():
            save_path[0] = None
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=8)

        ttk.Button(btn_frame, text="保存到原文件", command=save_to_original).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="另存为新文件", command=save_as_new).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=cancel_save).pack(side=tk.LEFT, padx=5)

        self.root.wait_window(dialog)

        if not save_path[0]:
            return

        success, message = excel_processor.excel_processor.save_workbook(save_path[0])
        if success:
            self.log_status(f"✓ {message}")
            messagebox.showinfo("成功", f"文件已保存到:\n{save_path[0]}")
        else:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", message)

    def check_data(self):
        """核对功能 - 支持多轮核对（每轮使用不同的Sheet组）"""
        if not self.current_file_path:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return

        # 第一步：选择核对文件和工作表（两个步骤共用）
        if not self.check_file_path:
            self.check_file_path = filedialog.askopenfilename(
                title="选择核对 Excel 文件",
                filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
            )

            if not self.check_file_path:
                return

            self.log_status(f"已选择核对文件：{os.path.basename(self.check_file_path)}")

            self.check_sheet_name, self.check_workbook = self.select_external_sheet(
                self.check_file_path,
                "选择核对工作表",
                "请选择包含姓名和需扣回/预支扣款/预支扣回数据的工作表："
            )

            if not self.check_sheet_name or not self.check_workbook:
                self.check_file_path = None
                return

            self.log_status(f"已选择工作表：{self.check_sheet_name}")
        else:
            self.log_status(f"使用已选择的核对文件：{os.path.basename(self.check_file_path)}")

        current_round = excel_processor.excel_processor.get_current_check_round()
        names_num, data_num, result_num = excel_processor.excel_processor._get_round_sheet_numbers()

        self.log_status(f"\n===== 第 {current_round} 轮核对开始 =====")
        self.log_status(f"本轮将使用：Sheet{names_num}(姓名)、Sheet{data_num}(数据)、Sheet{result_num}(结果)")

        required_sheets = [f'Sheet{names_num}', f'Sheet{data_num}', f'Sheet{result_num}']

        missing_sheets = [s for s in required_sheets if s not in excel_processor.excel_processor.sheet_names]
        if missing_sheets:
            success, message = excel_processor.excel_processor._create_check_sheets()
            if not success:
                self.log_status(f"✗ {message}")
                messagebox.showerror("错误", f"创建工作表失败：\n{message}")
                return
            excel_processor.excel_processor.sheet_names = excel_processor.excel_processor.workbook.sheetnames
            self.log_status(f"✓ 已创建缺失的工作表：{', '.join(missing_sheets)}")
        else:
            self.log_status(f"✓ 已确认 Sheet{names_num}、Sheet{data_num}、Sheet{result_num} 存在")

        success, message = excel_processor.excel_processor.copy_names_from_workbook(
            self.check_workbook,
            self.check_sheet_name
        )

        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", f"复制失败：\n{message}")
            return

        self.log_status(f"✓ 姓名已复制到 Sheet{names_num} 工作表中")

        merged_sheet = self.get_selected_sheet("选择匹配姓名工作表", "请选择需要匹配姓名的工作表：")
        if not merged_sheet:
            return

        success, message = excel_processor.excel_processor.apply_vlookup_formula(merged_sheet)
        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", message)
            return

        if message.startswith("UNITS:"):
            units_str = message.replace("UNITS:", "")
            available_units = units_str.split("|")

            selected_units = self.select_units(available_units)
            if selected_units is None:
                return

            success, message = excel_processor.excel_processor.apply_vlookup_formula(merged_sheet, selected_units)
            if not success:
                self.log_status(f"✗ {message}")
                messagebox.showerror("错误", message)
                return

        self.log_status(f"✓ {message}")

        self.log_status(f"开始处理需扣回/预支扣款/预支扣回数据...")

        selected_column = self.select_deduction_column()

        success, message = excel_processor.excel_processor.process_deduction_data(
            self.check_workbook,
            self.check_sheet_name,
            selected_column
        )

        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", f"处理需扣回/预支扣款/预支扣回数据失败：\n{message}")
            return

        self.log_status(f"✓ {message}")
        self.log_status(f"===== 第 {current_round} 轮核对完成 =====\n")

        messagebox.showinfo("成功",
            f"第 {current_round} 轮核对完成！\n{message}\n\n"
            f"数据已保存到：\n"
            f"  Sheet{names_num}: 姓名数据\n"
            f"  Sheet{data_num}: 核对数据\n"
            f"  Sheet{result_num}: 平账结果（待平账后生成）\n\n"
            f"您可以：\n"
            f"1. 点击【平账】完成本轮平账\n"
            f"2. 再次点击【核对】进行下一轮核对（将使用 Sheet{names_num+3}/{data_num+3}/{result_num+3}）")

    def select_units(self, available_units):
        """选择单位（支持多选）"""
        dialog = tk.Toplevel(self.root)
        dialog.title("选择单位")
        dialog.geometry("380x480")
        dialog.transient(self.root)
        dialog.grab_set()

        self.center_dialog(dialog, 380, 480)

        tk.Label(dialog, text="请选择要筛选的单位（可多选）：").pack(pady=(12, 8))

        btn_frame_top = tk.Frame(dialog)
        btn_frame_top.pack(pady=4)

        unit_vars = {}

        def select_all():
            for var in unit_vars.values():
                var.set(True)

        def deselect_all():
            for var in unit_vars.values():
                var.set(False)

        ttk.Button(btn_frame_top, text="全选", command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame_top, text="全不选", command=deselect_all).pack(side=tk.LEFT, padx=5)

        list_container = tk.Frame(dialog)
        list_container.pack(pady=8, padx=20, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        def on_dialog_close():
            try:
                canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass
            dialog.destroy()

        for unit in available_units:
            var = tk.BooleanVar(value=True)
            unit_vars[unit] = var
            chk = ttk.Checkbutton(scroll_frame, text=unit, variable=var)
            chk.pack(anchor=tk.W, pady=2, padx=5)

        selected_units = [None]

        def on_confirm():
            selected = [unit for unit, var in unit_vars.items() if var.get()]
            if not selected:
                messagebox.showwarning("警告", "请至少选择一个单位")
                return
            selected_units[0] = selected
            on_dialog_close()

        def on_cancel():
            selected_units[0] = None
            on_dialog_close()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=10)

        self.root.wait_window(dialog)
        return selected_units[0]

    def select_deduction_column(self):
        """选择需扣回数据类型（当有多个列有数据时）"""
        available_columns = []

        name_col = None
        deduction_col = None
        advance_deduction_col = None
        advance_return_col = None

        for col in range(1, self.check_workbook.max_column + 1):
            cell_value = self.check_workbook.cell(row=2, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                if cell_str == "姓名":
                    name_col = col
                elif "需扣回" in cell_str:
                    deduction_col = col
                elif "预支扣款" in cell_str:
                    advance_deduction_col = col
                elif "预支扣回" in cell_str:
                    advance_return_col = col

        if deduction_col:
            for row in range(3, self.check_workbook.max_row + 1):
                val = self.check_workbook.cell(row=row, column=deduction_col).value
                if val is not None and str(val).strip() != '':
                    available_columns.append('需扣回')
                    break

        if advance_deduction_col:
            for row in range(3, self.check_workbook.max_row + 1):
                val = self.check_workbook.cell(row=row, column=advance_deduction_col).value
                if val is not None and str(val).strip() != '':
                    available_columns.append('预支扣款')
                    break

        if advance_return_col:
            for row in range(3, self.check_workbook.max_row + 1):
                val = self.check_workbook.cell(row=row, column=advance_return_col).value
                if val is not None and str(val).strip() != '':
                    available_columns.append('预支扣回')
                    break

        if len(available_columns) <= 1:
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title("选择数据列")
        dialog.geometry("320x200")
        dialog.transient(self.root)
        dialog.grab_set()

        self.center_dialog(dialog, 320, 200)

        tk.Label(dialog, text="检测到多个数据列，请选择使用哪一列：").pack(pady=(16, 10))

        selected = [None]

        def on_select(col_type):
            selected[0] = col_type
            dialog.destroy()

        for col_type in available_columns:
            btn = ttk.Button(dialog, text=col_type, command=lambda c=col_type: on_select(c))
            btn.pack(pady=4, fill=tk.X, padx=30)

        self.root.wait_window(dialog)
        return selected[0]

    def select_external_sheet(self, file_path, title, prompt):
        """选择外部文件的某个工作表"""
        temp_processor = excel_processor.ExcelProcessor()
        try:
            success, message, sheet_names = temp_processor.select_excel_file(file_path)

            if not success:
                self.log_status(f"✗ {message}")
                messagebox.showerror("错误", f"读取文件失败：\n{message}")
                return None, None

            self.log_status(f"✓ {message}")

            dialog = tk.Toplevel(self.root)
            dialog.title(title)
            dialog.geometry("440x360")
            dialog.transient(self.root)
            dialog.grab_set()

            self.center_dialog(dialog, 440, 360)

            tk.Label(dialog, text=prompt).pack(pady=(12, 8))

            list_container = tk.Frame(dialog)
            list_container.pack(pady=8, padx=20, fill=tk.BOTH, expand=True)

            listbox = tk.Listbox(list_container, height=10)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            listbox.config(yscrollcommand=scrollbar.set)

            for sheet in sheet_names:
                listbox.insert(tk.END, sheet)

            selected_sheet = [None]
            selected_workbook = [None]

            def on_select():
                selection = listbox.curselection()
                if selection:
                    selected_sheet[0] = sheet_names[selection[0]]
                    selected_workbook[0] = temp_processor.workbook[selected_sheet[0]]
                    dialog.destroy()

            def on_cancel():
                selected_sheet[0] = None
                dialog.destroy()

            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=10)

            ttk.Button(btn_frame, text="确定", command=on_select).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

            self.root.wait_window(dialog)

            if not selected_sheet[0]:
                return None, None

            return selected_sheet[0], selected_workbook[0]
        finally:
            temp_processor.close_workbook()

    def balance_accounts(self):
        """平账功能 - 平账后自动递增轮次，为下一轮核对做准备"""
        if not self.current_file_path:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return

        salary_sheet = self.get_selected_sheet("请选择工作表", "请选择需要平账的工作表：")
        if not salary_sheet:
            return

        month = self.select_month()
        if not month:
            return

        current_round = excel_processor.excel_processor.get_current_check_round()
        names_num, data_num, result_num = excel_processor.excel_processor._get_round_sheet_numbers()

        self.log_status(f"\n===== 第 {current_round} 轮平账开始 =====")
        self.log_status(f"使用数据：Sheet{data_num} -> 结果保存到 Sheet{result_num}")
        self.log_status(f"平账月份：{month}")

        success, message = excel_processor.excel_processor.balance_accounts(salary_sheet, month)

        if success:
            self.log_status(f"✓ {message}")

            messagebox.showinfo("成功",
                f"第 {current_round} 轮平账完成！\n{message}\n\n"
                f"平账结果已保存到：\n"
                f"  Sheet{result_num}\n\n"
                f"轮次已自动递增，下次核对将使用新的Sheet组。")

            next_round = excel_processor.excel_processor.increment_round()
            self.log_status(f"✓ 轮次已递增：{current_round} -> {next_round}")
            self.log_status(f"===== 第 {current_round} 轮平账完成 =====\n")
        else:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", message)

    def select_month(self):
        """选择平账月份"""
        dialog = tk.Toplevel(self.root)
        dialog.title("选择平账月份")
        dialog.geometry("300x180")
        dialog.transient(self.root)
        dialog.grab_set()

        self.center_dialog(dialog, 300, 180)

        tk.Label(dialog, text="请选择平账月份：").pack(pady=(16, 10))

        month_var = tk.StringVar()
        month_spinbox = ttk.Spinbox(dialog, from_=1, to=12, width=12, textvariable=month_var)
        month_spinbox.pack(pady=4)
        month_spinbox.set(1)

        selected_month = [None]

        def on_confirm():
            try:
                month_num = int(month_spinbox.get())
                if 1 <= month_num <= 12:
                    selected_month[0] = f"{month_num}月"
                else:
                    messagebox.showwarning("警告", "请输入1-12之间的月份")
                    return
            except ValueError:
                messagebox.showwarning("警告", "请输入有效的月份数字")
                return
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=14)

        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=10)

        self.root.wait_window(dialog)
        return selected_month[0]

    def get_selected_sheet(self, title, prompt):
        """获取用户选择的工作表"""
        if not self.sheet_names:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("440x360")
        dialog.transient(self.root)
        dialog.grab_set()

        self.center_dialog(dialog, 440, 360)

        tk.Label(dialog, text=prompt).pack(pady=(12, 8))

        list_container = tk.Frame(dialog)
        list_container.pack(pady=8, padx=20, fill=tk.BOTH, expand=True)

        listbox = tk.Listbox(list_container, height=10)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)

        for sheet in self.sheet_names:
            listbox.insert(tk.END, sheet)

        selected_sheet = [None]

        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_sheet[0] = self.sheet_names[selection[0]]
                dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="确定", command=on_select).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        self.root.wait_window(dialog)
        return selected_sheet[0]

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


def main():
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
