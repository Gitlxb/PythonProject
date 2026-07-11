# -- coding: utf-8 --
# @File : dgf_cf_frame.py
# @Description: 代工费拆分工具 - Frame嵌入版
# @Note: 核心逻辑完全复用 dgf_cf_xzy.py，不修改原文件

import tkinter as tk
from tkinter import ttk, filedialog

import openpyxl

from .dgf_cf_xzy import (
    get_sheet_names, process_first_excel,
    write_split_rows,
    init_b_state, parse_remark1_from_filename,
    fill_split_rows,
    get_company_prefix, build_appended_records
)


class DgfCfFrame(ttk.Frame):
    """代工费拆分功能 - 作为Frame嵌入主界面"""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app  # 引用主应用，用于更新状态栏
        self.summary_data = None
        self._build_ui()

    # ==================== UI构建 ====================

    def _build_ui(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        ttk.Label(
            main_frame, text="代工费拆分工具_批量处理",
            font=("Microsoft YaHei", 16, "bold")
        ).pack(pady=(0, 15))

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=12)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 文件1：代工费表格（支持多选）
        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X, pady=5)
        ttk.Label(row1, text="代工费表格：", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT)
        self.lbl_file1 = ttk.Label(row1, text="未选择", font=("Microsoft YaHei", 9), foreground="gray")
        self.lbl_file1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.btn_file1 = ttk.Button(row1, text="选择", command=self.select_file1, width=8)
        self.btn_file1.pack(side=tk.RIGHT)

        # 文件列表显示区域
        self.file1_list_frame = ttk.Frame(file_frame)
        self.file1_list_frame.pack(fill=tk.X, pady=5, padx=(40, 0))
        self.file1_listbox = tk.Listbox(self.file1_list_frame, height=4, font=("Microsoft YaHei", 9))
        self.file1_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scrollbar = ttk.Scrollbar(self.file1_list_frame, orient=tk.VERTICAL, command=self.file1_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file1_listbox.config(yscrollcommand=scrollbar.set)

        # 文件2：发放记录表格
        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=5)
        ttk.Label(row2, text="发放记录表格：", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT)
        self.lbl_file2 = ttk.Label(row2, text="未选择", font=("Microsoft YaHei", 9), foreground="gray")
        self.lbl_file2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.btn_file2 = ttk.Button(row2, text="选择", command=self.select_file2, width=8)
        self.btn_file2.pack(side=tk.RIGHT)

        # 处理按钮
        self.btn_start = ttk.Button(
            main_frame, text="开始处理",
            command=self.start_process, width=18
        )
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
        self.file1_list = []  # 代工费表格列表（多文件）
        self.sheet1_list = []  # 对应的工作簿列表
        self.file2 = None
        self.sheet2 = None

    # ==================== 工具方法 ====================

    def _get_toplevel(self):
        """获取顶层窗口，用于弹窗"""
        return self.winfo_toplevel()

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        try:
            self._get_toplevel().update_idletasks()
        except Exception:
            pass
        print(msg)

    def _select_multiple_files(self, title):
        return filedialog.askopenfilenames(
            title=title,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )

    def show_custom_dialog(self, title, message, msg_type="info"):
        from tkinter import messagebox
        if msg_type == "error":
            messagebox.showerror(title, message)
        elif msg_type == "warning":
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)

    # ==================== 文件选择 ====================

    def select_file1(self):
        filepaths = self._select_multiple_files("选择代工费表格（含'介绍人'、'金额'，可多选）")
        if not filepaths:
            return
        
        self.file1_list = list(filepaths)
        self.sheet1_list = []
        self.file1_listbox.delete(0, tk.END)
        
        first_sheet = None  # 记录第一个文件选择的工作簿名
        
        # 为每个文件选择工作簿
        for i, filepath in enumerate(self.file1_list):
            filename = filepath.split("/")[-1].split("\\")[-1]
            self.file1_listbox.insert(tk.END, f"[{i+1}] {filename}")
            self.log(f"已选择代工费表格：{filepath}")
            
            # 第一个文件：弹窗选择工作簿
            if i == 0:
                sheet = self.ask_select_sheet(filepath, f"选择工作簿（后续文件将自动使用此工作簿）")
                if sheet:
                    first_sheet = sheet
                    self.sheet1_list.append(sheet)
                    self.log(f"  工作簿：{sheet}\n")
                else:
                    self.file1_list.pop(i)
                    self.file1_listbox.delete(i)
                    self.log(f"  [警告] 取消选择，已移除\n")
                    continue
            else:
                # 后续文件：尝试自动使用第一个文件的工作簿名
                try:
                    sheets = get_sheet_names(filepath)
                    if first_sheet in sheets:
                        self.sheet1_list.append(first_sheet)
                        self.log(f"  自动使用工作簿：{first_sheet}\n")
                    else:
                        # 工作簿不存在，弹窗选择
                        self.log(f"  [提示] 文件中不存在'{first_sheet}'，请手动选择")
                        sheet = self.ask_select_sheet(filepath, f"选择文件{i+1}的工作簿")
                        if sheet:
                            self.sheet1_list.append(sheet)
                            self.log(f"  工作簿：{sheet}\n")
                        else:
                            self.file1_list.pop(i)
                            self.file1_listbox.delete(i)
                            self.log(f"  [警告] 取消选择，已移除\n")
                            continue
                except Exception as e:
                    self.log(f"  [错误] {e}\n")
                    self.file1_list.pop(i)
                    self.file1_listbox.delete(i)
                    continue
        
        if self.file1_list:
            self.lbl_file1.config(text=f"已选择 {len(self.file1_list)} 个文件", foreground="black")
        else:
            self.lbl_file1.config(text="未选择", foreground="gray")
        
        self._check_ready()

    def select_file2(self):
        filepath = self._select_multiple_files("选择发放记录表格（含'供应商名称'、'打款金额'）")
        if not filepath:
            return
        # 只取第一个文件
        filepath = filepath[0]
        self.file2 = filepath
        self.lbl_file2.config(text=filepath.split("/")[-1].split("\\")[-1], foreground="black")
        self.log(f"已选择发放记录表格：{filepath}")

        self.sheet2 = self.ask_select_sheet(filepath, "选择发放记录表格的工作簿")
        if self.sheet2:
            self.log(f"  工作簿：{self.sheet2}\n")
            self._check_ready()
        else:
            self.file2 = None
            self.lbl_file2.config(text="未选择", foreground="gray")

    def _check_ready(self):
        if self.file1_list and self.sheet1_list and self.file2 and self.sheet2:
            self.btn_start.config(state=tk.NORMAL)
        else:
            self.btn_start.config(state=tk.DISABLED)

    def ask_select_sheet(self, filepath, title):
        try:
            sheets = get_sheet_names(filepath)
        except Exception as e:
            self.show_custom_dialog("错误", str(e), "error")
            return None

        if len(sheets) == 1:
            self.log(f"  自动选定唯一工作簿：{sheets[0]}")
            return sheets[0]

        dialog = tk.Toplevel(self._get_toplevel())
        dialog.title(title)
        dialog.geometry("360x200")
        dialog.transient(self._get_toplevel())
        dialog.grab_set()
        dialog.resizable(False, False)

        ttk.Label(dialog, text="检测到多个工作簿，请选择一个：",
                  font=("Microsoft YaHei", 10)).pack(pady=20)

        var = tk.StringVar(value=sheets[0])
        combo = ttk.Combobox(dialog, values=sheets, textvariable=var,
                             state="readonly", width=30, font=("Microsoft YaHei", 10))
        combo.pack(pady=5)

        result = [None]

        def on_ok():
            result[0] = combo.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_ok, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT, padx=10)

        self._get_toplevel().wait_window(dialog)
        return result[0]

    # ==================== 核心处理流程 ====================

    def start_process(self):
        self.btn_start.config(state=tk.DISABLED)
        self.log_text.delete(1.0, tk.END)

        wb = None
        try:
            if not self.file1_list or not self.file2:
                self.show_custom_dialog("提示", "请先选择 Excel 文件", "warning")
                return
            if not self.sheet1_list or not self.sheet2:
                self.show_custom_dialog("提示", "请选择工作簿", "warning")
                return

            self.log(f"代工费表格：共 {len(self.file1_list)} 个文件")
            for i, filepath in enumerate(self.file1_list):
                self.log(f"  [{i+1}] {filepath.split('/')[-1]}")
            self.log(f"发放记录表格：{self.file2}")
            self.log(f"  工作簿（A）：{self.sheet2}\n")

            wb = openpyxl.load_workbook(self.file2)

            # 步骤3：选择工作簿B（只读，一次性）
            self.log("【步骤3】选择发放记录表格的汇总工作簿...")
            sheet_b = self.ask_select_sheet(self.file2, "选择汇总工作簿（工作簿B）")
            if not sheet_b:
                self.log("  未选择工作簿B，跳过汇总拆分步骤。")
                wb.save(self.file2)
                self.log("全部处理完毕！")
                self.show_custom_dialog("完成", "数据已创建，未进行汇总拆分。", "info")
                return
            self.log(f"  工作簿（B）：{sheet_b}")

            # 读取B的初始状态（只读）
            b_state = init_b_state(self.file2, sheet_b)
            self.log(f"  已读取B的初始状态，共 {len(b_state)} 个收款人\n")

            # 串行处理每个代工费表格
            for i, (filepath, sheet_name) in enumerate(zip(self.file1_list, self.sheet1_list)):
                filename = filepath.split("/")[-1]
                self.log(f"{'='*60}")
                self.log(f"【文件 {i+1}/{len(self.file1_list)}】{filename}")
                self.log(f"{'='*60}")

                # 步骤1：汇总
                self.log("【步骤1】汇总'介绍人'和'金额'...")
                summary_data = process_first_excel(filepath, sheet_name)
                self.log(f"  汇总完成，共 {len(summary_data)} 条记录")
                self.log("  代工费表格已保存（汇总结果已写入原文件下方）")

                # 步骤2：解析公司前缀和备注1
                self.log("【步骤2】读取公司前缀并解析备注1...")
                company_prefix = get_company_prefix(filepath)
                self.log(f"  公司前缀：{company_prefix}")
                if not company_prefix:
                    self.log("  [警告] 未在'个税统计'工作簿中找到包含'发薪单位/发薪公司/公司'的单元格")
                remark1 = parse_remark1_from_filename(filename)
                self.log(f"  解析备注1：{remark1}")
                self.log("")

                # 步骤3：构建追加记录
                appended_records = build_appended_records(wb, summary_data)
                self.log(f"  已构建追加记录，共 {len(appended_records)} 条")

                # 步骤4：拆分判断（b_state会实时更新）
                all_used_names = {name for name, _, _ in appended_records}

                self.log("【步骤4】逐条判断追加记录并拆分...")
                split_records = []
                new_records = []
                split_start_row = None

                # 先预览累计（不修改）
                preview_state = dict(b_state)
                for name, amount, receiver_name in appended_records:
                    before = preview_state.get(receiver_name, 0.0) if receiver_name is not None else 0.0
                    after = before + amount
                    split_records.append((name, amount, receiver_name))
                    self.log(f"    需拆分: {name} ({receiver_name}) 累计 {before} + {amount} = {after}")
                    if receiver_name is not None:
                        preview_state[receiver_name] = after

                self.log(f"  判断完成，需拆分 {len(split_records)} 条。")

                # 执行实际拆分，获取结果
                if appended_records:
                    new_records, split_start_row = write_split_rows(
                        wb, appended_records, self.sheet2, b_state, all_used_names
                    )

                self.log("  拆分结果：")
                for candidate, amount, chosen_receiver in new_records:
                    receiver_display = chosen_receiver if chosen_receiver else ""
                    self.log(f"    {candidate}   {amount}   {receiver_display}")
                self.log("  发放记录表格（工作簿A）已保存。")

                # 步骤5：填充备注（使用当前文件的公司前缀和备注1）
                if split_records:
                    self.log("【步骤5】为拆分记录填充备注列...")
                    fill_split_rows(
                        wb, self.sheet2, split_start_row, new_records,
                        remark1, company_prefix
                    )
                    self.log("  填充完成，工作簿A已保存。")

                self.log(f"  文件{i+1}处理完成\n")

            wb.save(self.file2)
            self.log("="*60)
            self.log("全部处理完毕！")
            self.show_custom_dialog(
                "完成",
                f"所有 {len(self.file1_list)} 个文件已成功处理！\n可继续选择新文件进行处理。", "info"
            )
            self.app.status_label.config(text="代工费拆分 - 处理完成")

        except Exception as e:
            self.show_custom_dialog("处理错误", str(e), "error")
            self.log(f"  [错误] {e}")
        finally:
            self.btn_start.config(state=tk.NORMAL)
            if wb is not None:
                wb.close()
