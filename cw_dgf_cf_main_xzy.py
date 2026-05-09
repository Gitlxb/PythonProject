"""
GUI 界面模块：tkinter
与核心逻辑完全分离，仅负责交互与调用
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from cw_dgf_cf_xzy import get_sheet_names, process_first_excel, process_second_excel


class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("代工费拆分工具")
        self.root.geometry("650x520")
        self.root.resizable(False, False)
        self.summary_data = None
        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="代工费拆分工具", font=("Microsoft YaHei", 16, "bold")).pack(pady=(0, 15))

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 代工费表格
        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X, pady=5)
        ttk.Label(row1, text="代工费表格：", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT)
        self.lbl_file1 = ttk.Label(row1, text="未选择", font=("Microsoft YaHei", 9), foreground="gray")
        self.lbl_file1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self._file1_name = None
        self.btn_file1 = ttk.Button(row1, text="选择", command=self.select_file1, width=8)
        self.btn_file1.pack(side=tk.RIGHT)

        # 发放记录表格
        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=5)
        ttk.Label(row2, text="发放记录表格：", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT)
        self.lbl_file2 = ttk.Label(row2, text="未选择", font=("Microsoft YaHei", 9), foreground="gray")
        self.lbl_file2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self._file2_name = None
        self.btn_file2 = ttk.Button(row2, text="选择", command=self.select_file2, width=8)
        self.btn_file2.pack(side=tk.RIGHT)

        # 开始处理按钮
        self.btn_start = ttk.Button(main_frame, text="开始处理", command=self.start_process, width=18)
        self.btn_start.pack(pady=10)
        self.btn_start.config(state=tk.DISABLED)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = tk.Text(log_frame, height=14, wrap=tk.WORD, font=("Consolas", 10))
        vsb = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=vsb.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 状态变量
        self.file1 = None
        self.file2 = None
        self.sheet1 = None
        self.sheet2 = None

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _select_single_file(self, title):
        """选择单个 Excel 文件"""
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )

    def select_file1(self):
        """选择代工费表格（第1个文件）"""
        filepath = self._select_single_file("选择代工费表格（含'介绍人'、'合计金额'）")
        if not filepath:
            return
        self.file1 = filepath
        self._file1_name = filepath.split("/")[-1].split("\\")[-1]
        self.lbl_file1.config(text=self._file1_name, foreground="black")
        self.log(f"已选择代工费表格：{filepath}")

        # 立即选择工作簿
        self.sheet1 = self.ask_select_sheet(filepath, "选择代工费表格的工作簿")
        if self.sheet1:
            self.log(f"  工作簿：{self.sheet1}\n")
            self._check_ready()
        else:
            self.file1 = None
            self.lbl_file1.config(text="未选择", foreground="gray")

    def select_file2(self):
        """选择发放记录表格（第2个文件）"""
        filepath = self._select_single_file("选择发放记录表格（含'供应商名称'、'打款金额'）")
        if not filepath:
            return
        self.file2 = filepath
        self._file2_name = filepath.split("/")[-1].split("\\")[-1]
        self.lbl_file2.config(text=self._file2_name, foreground="black")
        self.log(f"已选择发放记录表格：{filepath}")

        # 立即选择工作簿
        self.sheet2 = self.ask_select_sheet(filepath, "选择发放记录表格的工作簿")
        if self.sheet2:
            self.log(f"  工作簿：{self.sheet2}\n")
            self._check_ready()
        else:
            self.file2 = None
            self.lbl_file2.config(text="未选择", foreground="gray")

    def _check_ready(self):
        """检查两个文件是否都已选择，启用开始按钮"""
        if self.file1 and self.file2 and self.sheet1 and self.sheet2:
            self.btn_start.config(state=tk.NORMAL)
        else:
            self.btn_start.config(state=tk.DISABLED)

    def ask_select_sheet(self, filepath, title):
        try:
            sheets = get_sheet_names(filepath)
        except Exception as e:
            messagebox.showerror("错误", str(e))
            return None

        if len(sheets) == 1:
            self.log(f"  自动选定唯一工作簿：{sheets[0]}")
            return sheets[0]

        # 多工作簿时弹出选择窗口
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("360x200")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        ttk.Label(dialog, text="检测到多个工作簿，请选择一个：", font=("Microsoft YaHei", 10)).pack(pady=20)

        var = tk.StringVar(value=sheets[0])
        combo = ttk.Combobox(dialog, values=sheets, textvariable=var, state="readonly", width=30, font=("Microsoft YaHei", 10))
        combo.pack(pady=5)

        result = [None]

        def on_ok():
            result[0] = var.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_ok, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT, padx=10)

        self.root.wait_window(dialog)
        return result[0]

    def start_process(self):
        self.btn_start.config(state=tk.DISABLED)
        self.log_text.delete(1.0, tk.END)

        try:
            # 检查是否都已选择
            if not self.file1 or not self.file2:
                messagebox.showwarning("提示", "请先选择两个 Excel 文件")
                return
            if not self.sheet1 or not self.sheet2:
                messagebox.showwarning("提示", "请选择工作簿")
                return

            self.log(f"代工费表格：{self.file1}")
            self.log(f"  工作簿：{self.sheet1}")
            self.log(f"发放记录表格：{self.file2}")
            self.log(f"  工作簿：{self.sheet2}\n")

            # ========== 处理代工费表格 ==========
            self.log("【步骤1】汇总'介绍人'和'合计金额'...")
            self.summary_data = process_first_excel(self.file1, self.sheet1)
            self.log(f"  汇总完成，共 {len(self.summary_data)} 条记录：")
            for name, amount in self.summary_data:
                self.log(f"    - {name}: {amount}")
            self.log("  代工费表格已保存（汇总结果已写入原文件下方）\n")

            # ========== 处理发放记录表格 ==========
            self.log("【步骤2】追加数据并填充公式...")
            process_second_excel(self.file2, self.summary_data, self.sheet2)
            self.log("  数据追加完成，VLOOKUP 及拆分处理完毕！")
            self.log("  发放记录表格已保存。\n")
            self.log("全部处理完毕！")
            messagebox.showinfo("完成", "所有操作已成功完成！\n可继续选择新文件进行处理。")

        except Exception as e:
            messagebox.showerror("处理错误", str(e))
            self.log(f"  [错误] {e}")
        finally:
            self.btn_start.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
