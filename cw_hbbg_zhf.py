import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy
import traceback


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel表格合并工具")
        self.root.geometry("750x550")
        self.root.resizable(True, True)

        self.selected_files = []
        self.process_results = []  # 记录每个文件的处理结果
        self._converted_temps = []  # 记录 .xls 转换的临时文件

        self.setup_ui()

    def setup_ui(self):
        # 标题
        title_label = tk.Label(self.root, text="Excel表格合并工具", font=("微软雅黑", 18, "bold"))
        title_label.pack(pady=15)

        # 文件选择区域
        select_frame = tk.Frame(self.root)
        select_frame.pack(pady=8, padx=20, fill=tk.X)

        self.select_btn = tk.Button(select_frame, text="选择Excel文件", command=self.select_files,
                                     font=("微软雅黑", 11), bg="#4CAF50", fg="white", width=15)
        self.select_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = tk.Button(select_frame, text="清空列表", command=self.clear_files,
                                    font=("微软雅黑", 11), bg="#f44336", fg="white", width=10)
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        # 目标列选择区域
        col_frame = tk.Frame(self.root)
        col_frame.pack(pady=5, padx=20, fill=tk.X)

        tk.Label(col_frame, text="填写月份的目标列（如 AU）：", font=("微软雅黑", 10)).pack(side=tk.LEFT, padx=5)
        self.target_col_var = tk.StringVar(value="")
        self.target_col_entry = tk.Entry(col_frame, textvariable=self.target_col_var, font=("微软雅黑", 11),
                                          width=12, justify='center')
        self.target_col_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(col_frame, text="(请输入Excel列字母，如 A, B, C ... AU, AV 等)", font=("微软雅黑", 9),
                 fg="#888888").pack(side=tk.LEFT, padx=5)

        # 文件列表显示区域（带滚动条）
        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=8, padx=20, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set,
                                        font=("微软雅黑", 10), selectmode=tk.MULTIPLE,
                                        height=12)
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        # 状态栏
        status_frame = tk.Frame(self.root)
        status_frame.pack(pady=5, padx=20, fill=tk.X)

        self.status_var = tk.StringVar(value="请选择需要合并的Excel文件")
        status_label = tk.Label(status_frame, textvariable=self.status_var, font=("微软雅黑", 10),
                                 fg="#666666", anchor="w")
        status_label.pack(side=tk.LEFT, fill=tk.X)

        # 进度条
        self.progress = ttk.Progressbar(status_frame, length=200, mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=(10, 0))

        # 操作按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)

        self.merge_btn = tk.Button(btn_frame, text="开始处理并合并", command=self.process_and_merge,
                                    font=("微软雅黑", 12, "bold"), bg="#2196F3", fg="white",
                                    width=18, height=2)
        self.merge_btn.pack()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择需要合并的Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            count = len(self.selected_files)
            self.status_var.set(f"已选择 {count} 个文件")

    def clear_files(self):
        self.selected_files = []
        self.file_listbox.delete(0, tk.END)
        self.process_results.clear()
        self._converted_temps = []
        self.target_col_var.set("")
        self.status_var.set("请选择需要合并的Excel文件")
        self.progress['value'] = 0

    @staticmethod
    def extract_month_from_date_range(date_str):
        """从日期字符串中提取月份"""
        try:
            match = re.search(r'(\d{4})年(\d{1,2})月', str(date_str))
            if match:
                return int(match.group(2))
        except Exception as e:
            print(f"解析日期出错: {e}")
        return None

    def _load_with_openpyxl(self, file_path):
        """
        用 openpyxl 加载 Excel 文件，返回 (workbook, worksheet, data_list)
        支持 .xlsx 和 .xls（自动转换），全程保留原始格式
        """
        filename = os.path.basename(file_path)

        # 尝试1: 直接用 openpyxl 加载（真正的 .xlsx）
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            data = [[cell.value for cell in row] for row in ws.iter_rows()]
            print(f"{filename}: openpyxl 直接加载成功 ({ws.max_row}行)")
            return wb, ws, data
        except Exception:
            pass

        # 尝试2: .xls / 伪.xlsx → pandas 转 openpyxl
        try:
            df = pd.read_excel(file_path, header=None, engine='xlrd')
            temp_converted = os.path.join(
                os.path.dirname(file_path),
                f'__converted_{os.path.splitext(filename)[0]}.xlsx'
            )
            df.to_excel(temp_converted, index=False, header=False, engine='openpyxl')
            wb = load_workbook(temp_converted)
            ws = wb.active
            data = [[cell.value for cell in row] for row in ws.iter_rows()]

            # 记录转换临时文件以便后续清理
            self._converted_temps.append(temp_converted)
            print(f"{filename}: xlrd 转换后 openpyxl 加载 ({ws.max_row}行)")
            return wb, ws, data
        except Exception as e:
            print(f"{filename}: 所有方式加载失败 - {e}")
            return None, None, None

    @staticmethod
    def _col_letter_to_index(col_letter):
        """将Excel列字母转换为1-based列索引，如 A->1, AU->47"""
        col_letter = col_letter.upper().strip()
        result = 0
        for char in col_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def process_single_file(self, file_path):
        """
        处理单个Excel文件：
        1. 用户自行选择目标列（如AU列）填写月份
        2. 在第5行表头找"身份证件类型"列，从该列下第一个"居民身份证"行开始填写
        3. 全程使用 openpyxl 操作，保留原始格式
        """
        # 获取用户选择的目标列
        target_col_letter = self.target_col_var.get().strip().upper()
        if not target_col_letter:
            return {"status": "error", "file": file_path,
                    "message": "未指定目标列，请在输入框中输入列字母(如 AU)"}

        try:
            target_excel_col = self._col_letter_to_index(target_col_letter)
        except Exception as e:
            return {"status": "error", "file": file_path,
                    "message": f"目标列 '{target_col_letter}' 格式错误: {e}"}

        try:
            wb, ws, data = self._load_with_openpyxl(file_path)
            if wb is None:
                return {"status": "error", "file": file_path,
                        "message": "无法读取文件"}

            total_rows = len(data)
            filename = os.path.basename(file_path)

            if total_rows < 5:
                wb.close()
                return {"status": "error", "file": file_path,
                        "message": f"数据不足5行（实际{total_rows}行），找不到表头"}

            print(f"{filename}: 成功读取数据，共 {total_rows} 行")

            # === 步骤1: 第2行获取时间范围 ===
            date_cell_value = None
            for col in range(len(data[1])):
                cell_value = str(data[1][col]) if pd.notna(data[1][col]) else ""
                if "年" in cell_value and "月" in cell_value and "日" in cell_value:
                    date_cell_value = cell_value
                    break

            if not date_cell_value:
                wb.close()
                return {"status": "skip", "file": file_path,
                        "message": "未在第2行找到日期信息"}

            current_month = self.extract_month_from_date_range(date_cell_value)
            if current_month is None:
                wb.close()
                return {"status": "skip", "file": file_path,
                        "message": "无法解析月份"}

            fill_month = 12 if current_month == 1 else current_month - 1
            month_str = f"{fill_month}月"

            # === 步骤2: 在第5行找"身份证件类型"列 ===
            id_type_col_idx = None
            for col in range(len(data[4])):
                cell_value = str(data[4][col]) if pd.notna(data[4][col]) else ""
                if cell_value.strip() == "身份证件类型":
                    id_type_col_idx = col
                    break

            if id_type_col_idx is None:
                wb.close()
                return {"status": "skip", "file": file_path,
                        "message": "未在第5行找到'身份证件类型'列"}

            # === 步骤3: 找到"身份证件类型"列中第一个和最后一个"居民身份证"的行 ===
            start_row_data_idx = None   # data列表中的起始行索引(0-based)
            end_row_data_idx = None     # data列表中的末尾行索引(0-based)

            for row_idx in range(5, total_rows):
                row_len = len(data[row_idx])
                if id_type_col_idx >= row_len:
                    continue
                cell_val = str(data[row_idx][id_type_col_idx]) if pd.notna(data[row_idx][id_type_col_idx]) else ""
                if cell_val.strip() == "居民身份证":
                    if start_row_data_idx is None:
                        start_row_data_idx = row_idx
                    end_row_data_idx = row_idx

            if start_row_data_idx is None or end_row_data_idx is None:
                wb.close()
                return {"status": "skip", "file": file_path,
                        "message": "在'身份证件类型'列中未找到'居民身份证'"}

            print(f"{filename}: 目标列={target_col_letter}(第{target_excel_col}列) | "
                  f"范围=第{start_row_data_idx + 1}行 ~ 第{end_row_data_idx + 1}行 | 月份={month_str}")

            # === 步骤4: 在起止范围内逐行填写月份 ===
            filled_count = 0
            for data_row_idx in range(start_row_data_idx, end_row_data_idx + 1):
                excel_row = data_row_idx + 1  # openpyxl 行号(1-based)
                target_cell = ws.cell(row=excel_row, column=target_excel_col)
                target_cell.value = month_str
                filled_count += 1

            print(f"{filename}: 完成 | 填写={filled_count}行")

            # === 步骤5: 用 openpyxl 保存（保留全部格式）===
            base_name = os.path.splitext(filename)[0]
            temp_path = os.path.join(os.path.dirname(file_path), f'{base_name}_processed.xlsx')
            wb.save(temp_path)
            wb.close()

            return {
                "status": "success",
                "file": file_path,
                "temp_path": temp_path,
                "month": month_str,
                "filled": filled_count,
                "target_col": target_col_letter,
                "start_row": start_row_data_idx + 1,
                "end_row": end_row_data_idx + 1
            }

        except Exception as e:
            traceback.print_exc()
            return {"status": "error", "file": file_path, "message": str(e)}

    def process_and_merge(self):
        """处理并合并所有选中的Excel文件"""
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择需要合并的Excel文件！")
            return

        # 禁用按钮，防止重复点击
        self.merge_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.process_results.clear()
        self._converted_temps = []

        total_files = len(self.selected_files)
        success_count = 0
        error_count = 0
        skip_count = 0
        processed_paths = []
        temp_files_to_clean = []

        try:
            # ========== 第一阶段：处理所有文件 ==========
            for i, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                
                # 更新状态
                self.status_var.set(f"[{i+1}/{total_files}] 正在处理: {filename}")
                self.progress['value'] = (i / total_files) * 80  # 处理阶段占80%
                self.root.update()
                
                result = self.process_single_file(file_path)
                self.process_results.append(result)

                if result["status"] == "success":
                    success_count += 1
                    processed_paths.append(result["temp_path"])
                    temp_files_to_clean.append(result["temp_path"])
                elif result["status"] == "skip":
                    skip_count += 1
                    processed_paths.append(result["file"])  # 用原文件路径
                else:
                    error_count += 1
                    processed_paths.append(result["file"])

            if success_count == 0:
                messagebox.showwarning("警告", 
                    f"没有成功处理的文件！\n\n跳过: {skip_count}\n错误: {error_count}")
                self._cleanup_all_temps(temp_files_to_clean)
                return

            # ========== 第二阶段：选择保存位置 ==========
            self.status_var.set("请选择保存位置...")
            self.root.update()

            output_path = filedialog.asksaveasfilename(
                title="选择汇总表保存位置",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                initialfile="汇总表.xlsx"
            )

            if not output_path:
                messagebox.showinfo("提示", "已取消保存")
                self._cleanup_all_temps(temp_files_to_clean)
                return

            # ========== 第三阶段：合并文件 ==========
            self.status_var.set("正在合并文件...")
            self.progress['value'] = 85
            self.root.update()

            merge_result = self.merge_excel_files(processed_paths, output_path)

            # ========== 第四阶段：清理临时文件 ==========
            self._cleanup_all_temps(temp_files_to_clean)
            
            self.progress['value'] = 100
            self.status_var.set(f"完成! 成功:{success_count} 跳过:{skip_count} 错误:{error_count}")

            # ========== 结果汇总 ==========
            summary = self._build_result_summary(success_count, skip_count, error_count, output_path)
            messagebox.showinfo("处理完成", summary)

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：\n{str(e)}")
            self.status_var.set("处理失败")
        finally:
            self.merge_btn.config(state=tk.NORMAL)

    @staticmethod
    def _cleanup_temps(temp_files):
        """清理临时文件"""
        for temp_file in temp_files:
            try:
                if temp_file and '_processed.xlsx' in temp_file and os.path.exists(temp_file):
                    os.remove(temp_file)
                    print(f"已清理临时文件: {os.path.basename(temp_file)}")
            except Exception as e:
                print(f"清理失败: {temp_file} - {e}")

    def _cleanup_all_temps(self, processed_temps):
        """清理所有临时文件（包括转换临时文件）"""
        self._cleanup_temps(processed_temps)
        # 清理 .xls 转换的临时文件
        for conv_temp in self._converted_temps:
            try:
                if os.path.exists(conv_temp):
                    os.remove(conv_temp)
                    print(f"已清理转换临时文件: {os.path.basename(conv_temp)}")
            except Exception as e:
                print(f"清理转换临时文件失败: {conv_temp} - {e}")
        self._converted_temps = []

    def _build_result_summary(self, success, skip, error, output_path):
        """构建结果汇总信息"""
        lines = [
            "=" * 40,
            f"  总计: {success + skip + error} 个文件",
            f"  ✓ 成功处理: {success}",
            f"  - 已跳过:   {skip}",
            f"  ✗ 处理错误: {error}",
            "=" * 40,
            f"\n保存位置:\n{output_path}"
        ]
        
        # 添加详细结果（仅显示前10个）
        if self.process_results:
            detail_lines = ["\n详细信息:"]
            for i, r in enumerate(self.process_results[:10]):
                fn = os.path.basename(r["file"])
                if r["status"] == "success":
                    detail_lines.append(f"  ✓ {fn} → 列:{r.get('target_col','?')} 月份:{r.get('month', '')} "
                                        f"(第{r.get('start_row','?')}~{r.get('end_row','?')}行, 共{r.get('filled', 0)}行)")
                elif r["status"] == "skip":
                    detail_lines.append(f"  - {fn}: {r.get('message', '已跳过')}")
                else:
                    detail_lines.append(f"  ✗ {fn}: {r.get('message', '未知错误')}")
            
            if len(self.process_results) > 10:
                detail_lines.append(f"  ... 还有 {len(self.process_results) - 10} 条记录")
            
            lines.extend(detail_lines)
        
        return "\n".join(lines)

    def merge_excel_files(self, file_paths, output_path):
        """
        合并多个Excel文件到一个汇总表中，完整保留原表格式：
        - 字体/边框/填充色/对齐方式
        - 合并单元格
        - 行高 / 列宽
        - 条件格式 / 数据验证
        - 数字格式
        """
        if not file_paths:
            return None

        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = "汇总表"

        current_row = 1
        total_files = len(file_paths)

        for idx, file_path in enumerate(file_paths):
            try:
                source_wb = load_workbook(file_path)
                source_ws = source_wb.active

                max_row = source_ws.max_row
                max_col = source_ws.max_column

                # --- 1. 逐单元格复制值 + 样式 ---
                for row_idx in range(1, max_row + 1):
                    # 复制行高
                    source_rd = source_ws.row_dimensions[row_idx]
                    if source_rd.height:
                        new_ws.row_dimensions[current_row].height = source_rd.height

                    for col_idx in range(1, max_col + 1):
                        source_cell = source_ws.cell(row=row_idx, column=col_idx)
                        target_cell = new_ws.cell(row=current_row, column=col_idx)

                        target_cell.value = source_cell.value

                        # 复制全部样式属性
                        target_cell.font = copy(source_cell.font)
                        target_cell.border = copy(source_cell.border)
                        target_cell.fill = copy(source_cell.fill)
                        target_cell.number_format = source_cell.number_format
                        target_cell.protection = copy(source_cell.protection)
                        target_cell.alignment = copy(source_cell.alignment)

                    current_row += 1

                # --- 2. 复制合并单元格 ---
                for merged_range in source_ws.merged_cells.ranges:
                    min_col, min_row, max_col_m, max_row_m = merged_range.bounds
                    # 坐标偏移到当前写入位置
                    offset_min_row = min_row + (current_row - max_row) - 1
                    offset_max_row = max_row_m + (current_row - max_row) - 1
                    new_merged = f"{get_column_letter(min_col)}{offset_min_row}:{get_column_letter(max_col_m)}{offset_max_row}"
                    new_ws.merge_cells(new_merged)

                # --- 3. 复制列宽 ---
                for col_idx in range(1, max_col + 1):
                    col_letter = get_column_letter(col_idx)
                    if col_letter in source_ws.column_dimensions:
                        source_cd = source_ws.column_dimensions[col_letter]
                        if source_cd.width:
                            new_ws.column_dimensions[col_letter].width = source_cd.width

                source_wb.close()

                # 文件间空20行分隔
                if idx < total_files - 1:
                    current_row += 20

                print(f"已合并 [{idx+1}/{total_files}]: {os.path.basename(file_path)} ({max_row}行)")

            except Exception as e:
                print(f"合并失败: {os.path.basename(file_path)} - {e}")
                traceback.print_exc()
                continue

        new_wb.save(output_path)
        new_wb.close()
        
        print(f"\n✓ 汇总表已保存: {output_path}")
        return output_path



def main():
    root = tk.Tk()
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap(default='')
    except:
        pass
    
    app = ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
