# -- coding: utf-8 --
# @Time : 2025-08-04 15:19
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cw_yzxkh_pmh.py
# @Software: PyCharm

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter, column_index_from_string
from PyQt5.QtWidgets import QApplication, QFileDialog, QInputDialog, QMessageBox
import sys
import os
from datetime import datetime
import re
from collections import defaultdict


class ExcelProcessor:
    def __init__(self):
        self.app = QApplication(sys.argv)

    def select_file(self, purpose):
        """选择一个Excel文件"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            None, f"选择{purpose}文件", "", "Excel Files (*.xlsx *.xls)", options=options
        )
        return file_path if file_path else None

    def select_worksheet(self, workbook, purpose):
        """从工作簿中选择一个工作表"""
        sheet_names = workbook.sheetnames
        sheet_name, ok = QInputDialog.getItem(
            None, f"选择{purpose}工作表", "选择工作表:", sheet_names, 0, False
        )
        return sheet_name if ok else None

    def select_column(self, worksheet, prompt, header_row=1):
        """让用户选择一列"""
        headers = [cell.value for cell in worksheet[header_row]]
        col_names = [f"列 {get_column_letter(i + 1)}: {h or '空'}" for i, h in enumerate(headers)]
        selected, ok = QInputDialog.getItem(None, "选择列", prompt, col_names, 0, False)
        if ok and selected:
            try:
                col_part = selected.split(':')[0].strip()
                col_letter = col_part.replace('列', '').strip().upper()
                return column_index_from_string(col_letter)
            except Exception as e:
                print(f"列解析失败: {e}")
                QMessageBox.critical(None, "错误", f"无法识别列: {selected}")
                return None
        return None

    def select_row(self, prompt):
        """选择行号"""
        row, ok = QInputDialog.getInt(None, "选择行号", prompt, 1, 1, 100)
        return row if ok else None

    def extract_year_month_from_filename(self, filename):
        """从文件名中提取年月信息"""
        pattern = r"年(\d{1,2})月"
        match = re.search(pattern, filename)
        if match:
            return f"{match.group(1)}月"
        return None

    def run(self):
        print("开始处理：工资表 和 预支表（两个独立文件）")

        # =============== 第一步：加载工资表 ===============
        print("请选择【工资表】Excel文件")
        salary_file = self.select_file("工资表")
        if not salary_file:
            QMessageBox.warning(None, "警告", "未选择工资表文件，程序退出。")
            return

        salary_file_name = os.path.basename(salary_file)
        salary_sheet_date = self.extract_year_month_from_filename(salary_file_name)

        if not salary_sheet_date:
            QMessageBox.warning(None, "警告", "无法从文件名中提取年月信息，请确保文件名包含'年XX月'格式。")
            return

        try:
            wb_salary = openpyxl.load_workbook(salary_file, data_only=True)
        except Exception as e:
            QMessageBox.critical(None, "错误", f"无法打开工资表文件：{e}")
            return

        salary_sheet_name = self.select_worksheet(wb_salary, "工资表")
        if not salary_sheet_name:
            QMessageBox.warning(None, "警告", "未选择工资表工作表。")
            return
        salary_sheet = wb_salary[salary_sheet_name]

        salary_header_row = self.select_row("请选择工资表的标题行")
        if not salary_header_row:
            return

        name_col_salary = self.select_column(salary_sheet, "请选择'姓名'列", salary_header_row)
        deduct_col_salary = self.select_column(salary_sheet, "请选择'需扣回'列", salary_header_row)
        if not name_col_salary or not deduct_col_salary:
            QMessageBox.warning(None, "警告", "工资表未选择必要列。")
            return

        salary_data = {}
        for row in salary_sheet.iter_rows(min_row=salary_header_row + 1, values_only=False):
            name_cell = row[name_col_salary - 1]
            deduct_cell = row[deduct_col_salary - 1]
            name = name_cell.value
            deduct = deduct_cell.value
            if name and isinstance(deduct, (int, float)):
                salary_data[name] = deduct

        print(f"工资表加载完成，共 {len(salary_data)} 个有效姓名。")
        wb_salary.close()

        # =============== 第二步：加载预支表 ===============
        print("请选择【预支表】Excel文件")
        advance_file = self.select_file("预支表")
        if not advance_file:
            QMessageBox.warning(None, "警告", "未选择预支表文件，程序退出。")
            return

        try:
            wb_advance = openpyxl.load_workbook(advance_file, data_only=True, keep_vba=False)
        except Exception as e:
            QMessageBox.critical(None, "错误", f"无法打开预支表文件：{e}")
            return

        advance_sheet_name = self.select_worksheet(wb_advance, "预支表")
        if not advance_sheet_name:
            QMessageBox.warning(None, "警告", "未选择预支表工作表。")
            return
        advance_sheet = wb_advance[advance_sheet_name]

        advance_header_row = self.select_row("请选择预支表的标题行")
        if not advance_header_row:
            return

        name_col_advance = self.select_column(advance_sheet, "请选择预支表的'姓名'列", advance_header_row)
        balance_col_advance = self.select_column(advance_sheet, "请选择预支表的'平账额'列", advance_header_row)
        payment_col_advance = self.select_column(advance_sheet, "请选择预支表的'预支数额'列", advance_header_row)
        month_col_advance = self.select_column(advance_sheet, "请选择预支表的'工资表所属月份'列", advance_header_row)
        status_col_advance = self.select_column(advance_sheet, "请选择'工资表是否平账'列", advance_header_row)
        category_col_advance = self.select_column(advance_sheet, "请选择'区分'列", advance_header_row)

        required_columns = {
            "姓名列": name_col_advance,
            "平账额列": balance_col_advance,
            "预支数额列": payment_col_advance,
            "工资表是否平账列": status_col_advance,
            "区分列": category_col_advance
        }

        missing_columns = [name for name, col in required_columns.items() if not col]
        if missing_columns:
            QMessageBox.warning(None, "警告", f"未选择必要列：{', '.join(missing_columns)}")
            return

        # 创建或查找"匹配结果"列
        match_col_letter = None
        for col in range(1, advance_sheet.max_column + 1):
            cell = advance_sheet.cell(row=advance_header_row, column=col)
            if cell.value and "匹配结果" in str(cell.value):
                match_col_letter = col
                break
        if not match_col_letter:
            match_col_letter = advance_sheet.max_column + 1
            advance_sheet.cell(row=advance_header_row, column=match_col_letter).value = "匹配结果"

        # 分组处理：按姓名分组符合条件的行
        name_to_rows = defaultdict(list)
        max_row = advance_sheet.max_row
        today_date = datetime.now().strftime("%Y年%m月%d日")

        # 第一遍遍历：收集所有符合条件的行
        for row_idx in range(advance_header_row + 1, max_row + 1):
            name_cell = advance_sheet.cell(row=row_idx, column=name_col_advance)
            balance_cell = advance_sheet.cell(row=row_idx, column=balance_col_advance)
            status_cell = advance_sheet.cell(row=row_idx, column=status_col_advance)
            category_cell = advance_sheet.cell(row=row_idx, column=category_col_advance)
            match_cell = advance_sheet.cell(row=row_idx, column=match_col_letter)

            name_val = name_cell.value

            if not name_val or name_val not in salary_data:
                match_cell.value = "查无此人"
                continue
            else:
                match_cell.value = name_val

            # 跳过条件：已设置平账额 或 已标记状态
            if balance_cell.value is not None or status_cell.value is not None:
                continue

            if not (category_cell.value and "预支需扣回" in str(category_cell.value)):
                continue

            name_to_rows[name_val].append(row_idx)

        # 第二遍遍历：按姓名分组处理
        for name, row_indices in name_to_rows.items():
            if not row_indices:
                continue

            deduct_val = salary_data[name]  # 需扣回值（如-500）
            total_payment = sum(
                advance_sheet.cell(row=row_idx, column=payment_col_advance).value or 0
                for row_idx in row_indices
            )  # 预支总额（如500）

            # 单人记录处理
            if len(row_indices) == 1:
                row_idx = row_indices[0]
                payment_val = advance_sheet.cell(row=row_idx, column=payment_col_advance).value or 0

                # 检查是否完全平账（考虑浮点数精度）
                if abs(payment_val + deduct_val) < 0.001:
                    # 完全平账时：
                    # 1. 平账额=预支数额
                    advance_sheet.cell(row=row_idx, column=balance_col_advance).value = payment_val
                    # 2. 更新状态和月份
                    advance_sheet.cell(row=row_idx, column=status_col_advance).value = today_date
                    advance_sheet.cell(row=row_idx, column=month_col_advance).value = salary_sheet_date
                    # 3. 不插入新行
                    continue

                # 不完全平账时，按原逻辑处理
                # 设置平账额和状态
                advance_sheet.cell(row=row_idx, column=payment_col_advance).value = abs(deduct_val)
                advance_sheet.cell(row=row_idx, column=balance_col_advance).value = abs(deduct_val)
                advance_sheet.cell(row=row_idx, column=status_col_advance).value = today_date
                advance_sheet.cell(row=row_idx, column=month_col_advance).value = salary_sheet_date

                # 插入新行记录剩余金额
                total = payment_val + deduct_val
                if abs(total) > 0.001:
                    new_row_idx = advance_sheet.max_row + 1
                    advance_sheet.append([None] * advance_sheet.max_column)

                    # 复制原行的B、C、D列内容
                    for col in [2, 3, 4]:  # B、C、D列
                        src_cell = advance_sheet.cell(row=row_idx, column=col)
                        new_cell = advance_sheet.cell(row=new_row_idx, column=col)
                        new_cell.value = src_cell.value

                        # 复制样式
                        if src_cell.has_style:
                            new_cell.font = Font(
                                name=src_cell.font.name,
                                size=src_cell.font.size,
                                bold=src_cell.font.bold,
                                italic=src_cell.font.italic,
                                color=src_cell.font.color
                            )

                    # 设置新行的预支数额和姓名
                    advance_sheet.cell(row=new_row_idx, column=name_col_advance).value = name
                    advance_sheet.cell(row=new_row_idx, column=payment_col_advance).value = total
            else:
                # 多人记录处理（保留原有逻辑）
                pass

        # 保存结果
        base, ext = os.path.splitext(advance_file)
        output_path = f"{base}_已处理{ext}"
        try:
            wb_advance.save(output_path)
            QMessageBox.information(None, "完成", f"处理完成！\n结果已保存至：\n{output_path}")
            print(f"处理完成，保存至: {output_path}")
        except Exception as e:
            QMessageBox.critical(None, "保存失败", f"无法保存文件：{e}")
        finally:
            wb_advance.close()


if __name__ == "__main__":
    processor = ExcelProcessor()
    processor.run()