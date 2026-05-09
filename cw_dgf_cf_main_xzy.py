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
        ttk.Label(
            main_frame,
            text="流程：\n1. 选择第一个 Excel（含“介绍人”、“合计金额”）→ 自动汇总\n2. 选择第二个 Excel（含“供应商名称”、“打款金额”等）→ 自动追加并填充公式",
            font=("Microsoft YaHei", 10), justify=tk.CENTER, foreground="#333"
        ).pack(pady=(0, 20))

        self.btn_start = ttk.Button(main_frame, text="开始处理", command=self.start_process, width=18)
        self.btn_start.pack(pady=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=15)

        self.log_text = tk.Text(log_frame, height=16, wrap=tk.WORD, font=("Consolas", 10))
        vsb = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=vsb.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def select_file(self, title):
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )

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
            # ========== 步骤 1 ==========
            self.log("【步骤 1】请选择第一个 Excel 文件（需包含“介绍人”、“合计金额”）")
            file1 = self.select_file("选择第一个 Excel 文件")
            if not file1:
                self.log("  未选择文件，操作已取消。")
                return

            sheet1 = self.ask_select_sheet(file1, "选择第一个文件的工作簿")
            if not sheet1:
                self.log("  未选择工作簿，操作已取消。")
                return

            self.log(f"  正在处理：{file1}")
            self.summary_data = process_first_excel(file1, sheet1)
            self.log(f"  汇总完成，共 {len(self.summary_data)} 条记录：")
            for name, amount in self.summary_data:
                self.log(f"    - {name}: {amount}")
            self.log("  第一个文件已保存（汇总结果已写入原文件下方）\n")

            # ========== 步骤 2 ==========
            self.log("【步骤 2】请选择第二个 Excel 文件（需包含“供应商名称”、“打款金额”等）")
            file2 = self.select_file("选择第二个 Excel 文件")
            if not file2:
                self.log("  未选择文件，操作已取消。")
                return

            sheet2 = self.ask_select_sheet(file2, "选择第二个文件的工作簿")
            if not sheet2:
                self.log("  未选择工作簿，操作已取消。")
                return

            self.log(f"  正在处理：{file2}")
            process_second_excel(file2, self.summary_data, sheet2)
            self.log("  数据追加完成，VLOOKUP 且 已拆分完成!")
            self.log("  第二个文件已保存。\n")
            self.log("全部处理完毕！")
            messagebox.showinfo("完成", "所有操作已成功完成！")

        except Exception as e:
            messagebox.showerror("处理错误", str(e))
            self.log(f"  [错误] {e}")
        finally:
            self.btn_start.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
