import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cw_yzpz_pmh as excel_processor
import os

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 数据处理工具")
        self.root.geometry("800x600")
            
        # 设置窗口在屏幕中央
        self.center_window()

        # 初始化变量
        self.current_file_path = None
        self.sheet_names = []
        # 核对步骤的共享变量
        self.check_file_path = None
        self.check_sheet_name = None
        self.check_workbook = None

        # 创建GUI组件
        self.create_widgets()

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
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件操作", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # 文件路径显示
        self.file_path_var = tk.StringVar()
        self.file_path_label = ttk.Label(file_frame, textvariable=self.file_path_var, wraplength=400)
        self.file_path_label.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))

        # 选择文件按钮
        select_file_btn = ttk.Button(file_frame, text="选择Excel文件", command=self.select_excel_file)
        select_file_btn.grid(row=1, column=0, padx=(0, 10))

        # 保存位置按钮
        save_location_btn = ttk.Button(file_frame, text="选择保存位置", command=self.select_save_location)
        save_location_btn.grid(row=1, column=1)

        # 功能按钮区域
        func_frame = ttk.LabelFrame(main_frame, text="功能操作", padding="10")
        func_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 核对按钮
        check_btn = ttk.Button(func_frame, text="核对", command=self.check_data)
        check_btn.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        # 平账按钮
        balance_btn = ttk.Button(func_frame, text="平账", command=self.balance_accounts)
        balance_btn.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="10")
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))

        self.status_text = tk.Text(status_frame, height=8, width=70)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        status_scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        status_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.status_text.config(yscrollcommand=status_scrollbar.set)

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)

    def log_status(self, message):
        """记录状态信息"""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()

    def select_excel_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )

        if file_path:
            self.current_file_path = file_path
            self.file_path_var.set(f"当前文件: {os.path.basename(file_path)}")

            # 读取文件
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
    
        # 直接显示两个按钮供用户选择
        dialog = tk.Toplevel(self.root)
        dialog.title("选择保存方式")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        self.center_dialog(dialog, 300, 150)
        
        ttk.Label(dialog, text="请选择保存方式：").pack(pady=15)
        
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
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="保存到原文件", command=save_to_original).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="另存为新文件", command=save_as_new).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=cancel_save).pack(side=tk.LEFT, padx=5)
        
        self.root.wait_window(dialog)
        
        if not save_path[0]:
            return
    
        # 保存文件
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
        # 注意：平账后不会自动重置，可以继续使用同一文件进行下一轮核对
        if not self.check_file_path:
            # 选择核对 Excel 文件
            self.check_file_path = filedialog.askopenfilename(
                title="选择核对 Excel 文件",
                filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
            )
            
            if not self.check_file_path:
                return
            
            self.log_status(f"已选择核对文件：{os.path.basename(self.check_file_path)}")
            
            # 选择包含数据的工作表
            self.check_sheet_name, self.check_workbook = self.select_external_sheet(
                self.check_file_path,
                "选择核对工作表",
                "请选择包含姓名和需扣回/预支扣款/预支扣回数据的工作表："
            )
            
            if not self.check_sheet_name or not self.check_workbook:
                self.check_file_path = None  # 取消则清空
                return
            
            self.log_status(f"已选择工作表：{self.check_sheet_name}")
        else:
            self.log_status(f"使用已选择的核对文件：{os.path.basename(self.check_file_path)}")
        
        # 获取当前轮次信息
        current_round = excel_processor.excel_processor.get_current_check_round()
        names_num, data_num, result_num = excel_processor.excel_processor._get_round_sheet_numbers()
        
        self.log_status(f"\n===== 第 {current_round} 轮核对开始 =====")
        self.log_status(f"本轮将使用：Sheet{names_num}(姓名)、Sheet{data_num}(数据)、Sheet{result_num}(结果)")
        
        # 第二步：处理姓名匹配
        # 1. 确保当前轮次所需的Sheet存在
        required_sheets = [f'Sheet{names_num}', f'Sheet{data_num}', f'Sheet{result_num}']
        
        missing_sheets = [s for s in required_sheets if s not in excel_processor.excel_processor.sheet_names]
        if missing_sheets:
            success, message = excel_processor.excel_processor._create_check_sheets()
            if not success:
                self.log_status(f"✗ {message}")
                messagebox.showerror("错误", f"创建工作表失败：\n{message}")
                return
            # 更新sheet_names列表
            excel_processor.excel_processor.sheet_names = excel_processor.excel_processor.workbook.sheetnames
            self.log_status(f"✓ 已创建缺失的工作表：{', '.join(missing_sheets)}")
        else:
            self.log_status(f"✓ 已确认 Sheet{names_num}、Sheet{data_num}、Sheet{result_num} 存在")
        
        # 2. 从核对文件复制姓名列到主文件的 Sheet{names_num}
        success, message = excel_processor.excel_processor.copy_names_from_workbook(
            self.check_workbook, 
            self.check_sheet_name
        )
        
        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", f"复制失败：\n{message}")
            # 注意：这里不重置 check_file_path，允许用户重新操作
            return
        
        self.log_status(f"✓ 姓名已复制到 Sheet{names_num} 工作表中")
        
        # 3. 在合并好的表格中应用姓名匹配并筛选到 Sheet{data_num}
        merged_sheet = self.get_selected_sheet("选择匹配姓名工作表", "请选择需要匹配姓名的工作表：")
        if not merged_sheet:
            return
        
        # 第一次调用：检测有哪些单位
        success, message = excel_processor.excel_processor.apply_vlookup_formula(merged_sheet)
        if not success:
            self.log_status(f"✗ {message}")
            messagebox.showerror("错误", message)
            return
        
        # 检查是否需要选择单位
        if message.startswith("UNITS:"):
            units_str = message.replace("UNITS:", "")
            available_units = units_str.split("|")
            
            # 弹窗让用户选择单位
            selected_units = self.select_units(available_units)
            if selected_units is None:  # 用户取消
                return
            
            # 第二次调用：使用选择的单位进行筛选
            success, message = excel_processor.excel_processor.apply_vlookup_formula(merged_sheet, selected_units)
            if not success:
                self.log_status(f"✗ {message}")
                messagebox.showerror("错误", message)
                return
        
        self.log_status(f"✓ {message}")
        
        # 第三步：处理需扣回/预支扣款/预支扣回数据（直接使用已选择的工作表）
        self.log_status(f"开始处理需扣回/预支扣款/预支扣回数据...")
        
        # 检测有哪些列有数据，让用户选择
        selected_column = self.select_deduction_column()
        
        # 处理需扣回/预支扣款/预支扣回数据并复制到 Sheet{data_num}
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
        dialog.geometry("350x450")
        dialog.transient(self.root)
        dialog.grab_set()
        
        self.center_dialog(dialog, 350, 450)
        
        ttk.Label(dialog, text="请选择要筛选的单位（可多选）：").pack(pady=10)
        
        # 全选/全不选按钮
        btn_frame_top = ttk.Frame(dialog)
        btn_frame_top.pack(pady=5)
        
        unit_vars = {}
        
        def select_all():
            for var in unit_vars.values():
                var.set(True)
        
        def deselect_all():
            for var in unit_vars.values():
                var.set(False)
        
        ttk.Button(btn_frame_top, text="全选", command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame_top, text="全不选", command=deselect_all).pack(side=tk.LEFT, padx=5)
        
        # 创建带滚动条的列表
        list_container = ttk.Frame(dialog)
        list_container.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(list_container, borderwidth=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 绑定鼠标滚轮
        def on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass  # 窗口已销毁，忽略错误
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # 对话框关闭时解绑事件
        def on_dialog_close():
            try:
                canvas.unbind_all("<MouseWheel>")
            except:
                pass
            dialog.destroy()
        
        # 创建复选框
        for unit in available_units:
            var = tk.BooleanVar(value=True)  # 默认全选
            unit_vars[unit] = var
            chk = ttk.Checkbutton(scroll_frame, text=unit, variable=var)
            chk.pack(anchor=tk.W, pady=2, padx=5)
        
        selected_units = [None]
        
        def on_confirm():
            # 收集所有选中的单位
            selected = [unit for unit, var in unit_vars.items() if var.get()]
            if not selected:
                messagebox.showwarning("警告", "请至少选择一个单位")
                return
            selected_units[0] = selected
            on_dialog_close()
        
        def on_cancel():
            selected_units[0] = None
            on_dialog_close()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        self.root.wait_window(dialog)
        return selected_units[0]
    
    def select_deduction_column(self):
        """选择需扣回数据类型（当有多个列有数据时）"""
        # 先检测哪些列有数据
        available_columns = []
        
        # 查找各列
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
        
        # 检查哪些列有有效数据
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
        
        # 如果只有一列或没有，直接返回None（自动选择）
        if len(available_columns) <= 1:
            return None
        
        # 多列时弹窗让用户选择
        dialog = tk.Toplevel(self.root)
        dialog.title("选择数据列")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        self.center_dialog(dialog, 300, 200)
        
        ttk.Label(dialog, text="检测到多个数据列，请选择使用哪一列：").pack(pady=10)
        
        selected = [None]
        
        def on_select(col_type):
            selected[0] = col_type
            dialog.destroy()
        
        for col_type in available_columns:
            ttk.Button(dialog, text=col_type, command=lambda c=col_type: on_select(c)).pack(pady=5, fill=tk.X, padx=20)
        
        self.root.wait_window(dialog)
        return selected[0]
    
    def select_external_sheet(self, file_path, title, prompt):
        """选择外部文件的某个工作表"""
        # 读取核对文件的工作表
        temp_processor = excel_processor.ExcelProcessor()
        try:
            success, message, sheet_names = temp_processor.select_excel_file(file_path)
            
            if not success:
                self.log_status(f"✗ {message}")
                messagebox.showerror("错误", f"读取文件失败：\n{message}")
                return None, None
            
            self.log_status(f"✓ {message}")
            
            # 让用户选择工作表
            dialog = tk.Toplevel(self.root)
            dialog.title(title)
            dialog.geometry("400x300")
            dialog.transient(self.root)
            dialog.grab_set()
            
            self.center_dialog(dialog, 400, 300)
            
            ttk.Label(dialog, text=prompt).pack(pady=10)
            
            listbox = tk.Listbox(dialog, height=10)
            listbox.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
            
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
            
            btn_frame = ttk.Frame(dialog)
            btn_frame.pack(pady=10)
            
            ttk.Button(btn_frame, text="确定", command=on_select).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)
            
            self.root.wait_window(dialog)
            
            if not selected_sheet[0]:
                return None, None
            
            return selected_sheet[0], selected_workbook[0]
        finally:
            # 确保释放资源
            temp_processor.close_workbook()

    def balance_accounts(self):
        """平账功能 - 平账后自动递增轮次，为下一轮核对做准备"""
        if not self.current_file_path:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return
    
        # 选择需扣回工资表或预支表明细表
        salary_sheet = self.get_selected_sheet("请选择工作表", "请选择需要平账的工作表：")
        if not salary_sheet:
            return
    
        # 选择月份
        month = self.select_month()
        if not month:
            return
    
        # 获取当前轮次信息（平账前）
        current_round = excel_processor.excel_processor.get_current_check_round()
        names_num, data_num, result_num = excel_processor.excel_processor._get_round_sheet_numbers()
    
        self.log_status(f"\n===== 第 {current_round} 轮平账开始 =====")
        self.log_status(f"使用数据：Sheet{data_num} -> 结果保存到 Sheet{result_num}")
        self.log_status(f"平账月份：{month}")
    
        # 执行平账
        success, message = excel_processor.excel_processor.balance_accounts(salary_sheet, month)
    
        if success:
            self.log_status(f"✓ {message}")
            
            messagebox.showinfo("成功", 
                f"第 {current_round} 轮平账完成！\n{message}\n\n"
                f"平账结果已保存到：\n"
                f"  Sheet{result_num}\n\n"
                f"轮次已自动递增，下次核对将使用新的Sheet组。")
            
            # 平账完成后递增轮次，为下一轮做准备
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
        dialog.geometry("250x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        self.center_dialog(dialog, 250, 150)
        
        ttk.Label(dialog, text="请选择平账月份：").pack(pady=10)
        
        month_var = tk.StringVar()
        month_spinbox = ttk.Spinbox(dialog, from_=1, to=12, width=10, textvariable=month_var)
        month_spinbox.pack(pady=5)
        
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
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        self.root.wait_window(dialog)
        return selected_month[0]
    
    def get_selected_sheet(self, title, prompt):
        """获取用户选择的工作表"""
        if not self.sheet_names:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return None

        # 创建选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        # 设置对话框位置在主窗口中央
        self.center_dialog(dialog, 400, 300)

        # 对话框内容
        ttk.Label(dialog, text=prompt).pack(pady=10)

        listbox = tk.Listbox(dialog, height=10)
        listbox.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

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

        btn_frame = ttk.Frame(dialog)
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