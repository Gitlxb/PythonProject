# -- coding: utf-8 --
# @Time : 2025-05-05 11:34
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : zh_pdf_Excel.py
# @Software: PyCharm

import pdfplumber
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tqdm import tqdm
from tkinter import ttk
import threading
import logging
from openpyxl import load_workbook
from openpyxl.styles import Alignment


class PDFToExcelConverter:
    def __init__(self, parent=None):
        # 如果提供了父窗口，使用Toplevel；否则创建新的Tk实例
        if parent:
            self.root = tk.Toplevel(parent)
            self.root.transient(parent)  # 设置为父窗口的临时窗口
            self.root.grab_set()  # 设置为模态窗口
        else:
            self.root = tk.Tk()
        
        self.root.title("PDF转Excel工具 - 对齐列版")
        self.root.geometry("550x350")

        # 初始化变量
        self.pdf_path = ""
        self.excel_path = ""
        self.conversion_in_progress = False
        self.append_mode = tk.BooleanVar(value=False)

        # 配置日志
        logging.basicConfig(filename='pdf_to_excel.log', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

        # 创建UI元素
        self.create_widgets()
        
        # 如果有父窗口，提升当前窗口到前面
        if parent:
            self.root.lift()
            self.root.focus_force()

    def create_widgets(self):
        """创建界面元素"""
        # PDF文件选择部分
        pdf_frame = tk.Frame(self.root)
        pdf_frame.pack(pady=10)

        tk.Label(pdf_frame, text="PDF文件:").pack(side=tk.LEFT)
        self.pdf_entry = tk.Entry(pdf_frame, width=45)
        self.pdf_entry.pack(side=tk.LEFT, padx=5)

        pdf_btn = tk.Button(pdf_frame, text="浏览...", command=self.select_pdf_file)
        pdf_btn.pack(side=tk.LEFT)

        # Excel文件选择部分
        excel_frame = tk.Frame(self.root)
        excel_frame.pack(pady=10)

        tk.Label(excel_frame, text="Excel输出:").pack(side=tk.LEFT)
        self.excel_entry = tk.Entry(excel_frame, width=45)
        self.excel_entry.pack(side=tk.LEFT, padx=5)

        excel_btn = tk.Button(excel_frame, text="浏览...", command=self.select_excel_path)
        excel_btn.pack(side=tk.LEFT)

        # 选项部分
        option_frame = tk.Frame(self.root)
        option_frame.pack(pady=5)

        tk.Checkbutton(option_frame, text="追加到现有文件", variable=self.append_mode).pack(side=tk.LEFT)

        # 进度条
        self.progress_label = tk.Label(self.root, text="准备就绪")
        self.progress_label.pack(pady=5)

        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=450, mode='determinate')
        self.progress.pack(pady=10)

        # 转换按钮
        self.convert_btn = tk.Button(self.root, text="开始转换", command=self.start_conversion)
        self.convert_btn.pack(pady=20)

    def select_pdf_file(self):
        """选择PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.pdf_path = file_path
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, file_path)

            # 自动设置默认输出文件名
            default_name = os.path.splitext(os.path.basename(file_path))[0] + "_对齐.xlsx"
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, os.path.join(os.path.dirname(file_path), default_name))

    def select_excel_path(self):
        """选择Excel输出路径"""
        default_name = self.excel_entry.get() or "output_对齐.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=os.path.basename(default_name)
        )
        if file_path:
            self.excel_path = file_path
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)

    def update_progress(self, value, text):
        """更新进度条"""
        self.progress['value'] = value
        self.progress_label.config(text=text)
        self.root.update_idletasks()

    def extract_table_with_alignment(self, table, chars):
        """提取表格数据并保留对齐信息"""
        aligned_data = []
        for row in table:
            aligned_row = []
            for cell in row:
                if cell is None:
                    aligned_row.append('')
                    continue

                # 获取字符级别的对齐信息
                cell_text = []
                for char in chars:
                    if char['text'] == cell:
                        cell_text.append({
                            'text': char['text'],
                            'x0': char['x0'],
                            'x1': char['x1'],
                            'width': char['width']
                        })
                        break

                if cell_text:
                    aligned_row.append(cell_text[0])
                else:
                    aligned_row.append({'text': cell, 'x0': 0, 'x1': len(cell), 'width': len(cell)})
            aligned_data.append(aligned_row)
        return aligned_data

    def process_table(self, page, page_num):
        """处理单个表格，返回对齐后的DataFrame"""
        tables = page.extract_tables()
        if not tables:
            return pd.DataFrame()

        # 提取字符级别的信息用于对齐
        chars = page.chars

        all_tables = []
        for table in tables:
            if not table:
                continue

            # 获取对齐后的表格数据
            aligned_table = self.extract_table_with_alignment(table, chars)
            if not aligned_table:
                continue

            # 处理表头
            headers = []
            header_row = aligned_table[0]
            for cell in header_row:
                headers.append(cell['text'] if isinstance(cell, dict) else str(cell))

            # 处理数据行
            data = []
            for row in aligned_table[1:]:
                data_row = []
                for cell in row:
                    data_row.append(cell['text'] if isinstance(cell, dict) else str(cell))
                data.append(data_row)

            # 创建DataFrame
            df = pd.DataFrame(data, columns=headers)

            # 添加元信息
            df['数据类型'] = '表格'
            df['来源页码'] = page_num

            all_tables.append(df)

        return pd.concat(all_tables, ignore_index=True) if all_tables else pd.DataFrame()

    def process_text(self, page, page_num):
        """处理文本内容，返回DataFrame"""
        text = page.extract_text()
        if not text:
            return pd.DataFrame()

        # 提取字符级别的信息用于对齐
        chars = page.chars
        if not chars:
            return pd.DataFrame()

        # 按行组织字符
        lines = {}
        for char in chars:
            # 估算行号（基于y坐标）
            line_key = round(char['top'], 1)
            if line_key not in lines:
                lines[line_key] = []
            lines[line_key].append(char)

        # 处理每行文本
        aligned_lines = []
        for line_key in sorted(lines.keys()):
            line_chars = sorted(lines[line_key], key=lambda x: x['x0'])

            # 构建行文本
            line_text = ''.join([char['text'] for char in line_chars])
            aligned_lines.append(line_text)

        if not aligned_lines:
            return pd.DataFrame()

        # 创建DataFrame
        df = pd.DataFrame(aligned_lines, columns=['内容'])
        df['数据类型'] = '文本'
        df['来源页码'] = page_num

        return df

    def apply_excel_alignment(self, worksheet, df, aligned_data):
        """应用Excel单元格对齐"""
        # 默认对齐方式
        default_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # 应用对齐到所有单元格
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.alignment = default_alignment

        # 如果有对齐数据，可以在这里添加特定对齐逻辑
        # ...

    def save_to_excel(self, df, excel_path, sheet_name="PDF内容"):
        """保存DataFrame到Excel并保持对齐"""
        try:
            # 创建新文件或追加到现有文件
            if self.append_mode.get() and os.path.exists(excel_path):
                book = load_workbook(excel_path)
                writer = pd.ExcelWriter(excel_path, engine='openpyxl')
                writer.book = book

                if sheet_name in book.sheetnames:
                    startrow = book[sheet_name].max_row
                    if startrow == 1:  # 只有标题行
                        startrow = 0
                else:
                    startrow = 0

                df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=(startrow == 0))
                writer.save()
                writer.close()

                # 重新打开文件应用格式
                book = load_workbook(excel_path)
                worksheet = book[sheet_name]
                self.apply_excel_alignment(worksheet, df, None)
                book.save(excel_path)
                book.close()
            else:
                # 创建新文件
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    # 获取worksheet对象应用格式
                    worksheet = writer.sheets[sheet_name]
                    self.apply_excel_alignment(worksheet, df, None)
        except Exception as e:
            raise Exception(f"保存Excel文件失败: {str(e)}")

    def conversion_thread(self):
        """执行转换的线程"""
        try:
            self.conversion_in_progress = True
            self.convert_btn.config(state=tk.DISABLED)
            self.update_progress(0, "正在准备转换...")
            logging.info("开始转换过程")

            if not self.pdf_path or not self.excel_path:
                messagebox.showwarning("警告", "请先选择PDF文件和输出路径!")
                return

            # 创建数据容器
            all_data = []

            # 先计算总页数用于进度条
            with pdfplumber.open(self.pdf_path) as pdf:
                total_pages = len(pdf.pages)

            self.update_progress(5, f"开始转换 {total_pages} 页...")
            logging.info(f"PDF总页数: {total_pages}")

            with pdfplumber.open(self.pdf_path) as pdf:
                for i, page in enumerate(tqdm(pdf.pages, desc="转换进度"), 1):
                    progress = 5 + (i / total_pages) * 90
                    self.update_progress(progress, f"正在处理第 {i}/{total_pages} 页...")

                    # 处理表格
                    table_df = self.process_table(page, i)
                    if not table_df.empty:
                        all_data.append(table_df)
                        # 添加分隔行
                        all_data.append(pd.DataFrame([{'数据类型': '分隔线', '来源页码': i}]))

                    # 处理文本
                    text_df = self.process_text(page, i)
                    if not text_df.empty:
                        all_data.append(text_df)
                        # 添加分隔行
                        all_data.append(pd.DataFrame([{'数据类型': '分隔线', '来源页码': i}]))

            # 合并所有数据
            if not all_data:
                self.update_progress(100, "转换完成 - 没有提取到内容")
                messagebox.showwarning("警告", "没有从PDF中提取到任何内容!")
                logging.warning("没有提取到任何内容")
                return

            combined_df = pd.concat(all_data, ignore_index=True)
            combined_df.fillna('', inplace=True)

            self.update_progress(95, "正在写入Excel文件...")
            logging.info(f"总共提取了 {len(combined_df)} 行数据")

            # 保存到Excel
            self.save_to_excel(combined_df, self.excel_path)

            self.update_progress(100, "转换完成!")
            messagebox.showinfo("成功", f"PDF已成功转换为Excel:\n{self.excel_path}")
            logging.info(f"转换成功，文件已保存到: {self.excel_path}")

        except Exception as e:
            error_msg = f"转换过程中发生错误: {str(e)}"
            self.update_progress(100, error_msg)
            messagebox.showerror("错误", error_msg)
            logging.error(error_msg, exc_info=True)

        finally:
            self.conversion_in_progress = False
            self.convert_btn.config(state=tk.NORMAL)
            logging.info("转换过程结束")

    def start_conversion(self):
        """开始转换"""
        if not self.conversion_in_progress:
            thread = threading.Thread(target=self.conversion_thread)
            thread.start()

    def run(self):
        """运行主程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = PDFToExcelConverter()
    app.run()