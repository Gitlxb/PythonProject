from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import warnings
from datetime import datetime
import gc

warnings.filterwarnings('ignore')

class ExcelProcessor:
    def __init__(self):
        self.workbook = None
        self.sheet_names = []
        self.current_file_path = None
        self._column_cache = {}  # 列位置缓存
        self.check_round = 1  # 核对轮次计数器（第1轮使用Sheet3/4/5，第2轮使用Sheet6/7/8，以此类推）
    
    def close_workbook(self):
        """关闭工作簿释放资源"""
        if self.workbook:
            try:
                self.workbook.close()
            except:
                pass
            finally:
                self.workbook = None
                self.sheet_names = []
                self._column_cache.clear()
                self.check_round = 1  # 重置轮次
                gc.collect()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_workbook()
        return False
    
    def _get_round_sheet_numbers(self):
        """根据当前轮次获取Sheet编号
        
        Returns:
            (names_num, data_num, result_num): 姓名表、数据表、结果表的编号
        """
        # 第1轮: 3,4,5; 第2轮: 6,7,8; 第3轮: 9,10,11
        base = 3 + (self.check_round - 1) * 3
        return base, base + 1, base + 2
    
    def increment_round(self):
        """增加轮次计数器，用于下一轮核对"""
        # 防止轮次无限增长（最多支持30轮，即Sheet90/91/92）
        if self.check_round >= 30:
            raise ValueError("轮次过多（已达30轮），建议保存文件后重新开始")
        
        self.check_round += 1
        self._column_cache.clear()  # 清空列缓存，避免不同轮次混淆
        return self.check_round
    
    def get_current_check_round(self):
        """获取当前核对轮次"""
        return self.check_round

    def select_excel_file(self, file_path):
        """选择 Excel 文件并读取工作表"""
        try:
            # 先关闭已有工作簿
            self.close_workbook()
            
            # 使用 openpyxl 读取工作簿
            self.workbook = load_workbook(file_path, data_only=True)
            self.sheet_names = self.workbook.sheetnames
            self.current_file_path = file_path
            self._column_cache.clear()  # 清空缓存
            return True, f"成功读取文件：{os.path.basename(file_path)}", self.sheet_names
        except Exception as e:
            return False, f"读取文件失败：{str(e)}", []

    def _find_column_by_header(self, worksheet, header_name, row=2, cache_key=None):
        """通用方法：通过表头名称查找列索引（带缓存）"""
        if cache_key and cache_key in self._column_cache:
            return self._column_cache[cache_key]
        
        for col in range(1, worksheet.max_column + 1):
            cell_value = worksheet.cell(row=row, column=col).value
            if cell_value and str(cell_value).strip() == header_name:
                if cache_key:
                    self._column_cache[cache_key] = col
                return col
        return None
    
    def _find_column_contains(self, worksheet, keyword, row=2, cache_key=None):
        """通用方法：通过关键词查找包含该词的列索引（带缓存）"""
        if cache_key and cache_key in self._column_cache:
            return self._column_cache[cache_key]
        
        for col in range(1, worksheet.max_column + 1):
            cell_value = worksheet.cell(row=row, column=col).value
            if cell_value and keyword in str(cell_value):
                if cache_key:
                    self._column_cache[cache_key] = col
                return col
        return None
    
    def detect_and_prepare(self, target_sheet_name):
        """检测功能 - 检查第二行是否有姓名1列，若无则在预支数额前插入，并创建核对工作表"""
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"
    
            target_sheet = self.workbook[target_sheet_name]
    
            # 查找"预支数额"列（在第二行）
            pre_advance_col = self._find_column_contains(target_sheet, "预支数额", cache_key=f"{target_sheet_name}_预支数额")
            if not pre_advance_col:
                return False, "未找到'预支数额'列"
    
            # 检查第二行是否已有"姓名1"列
            has_name1 = self._find_column_by_header(target_sheet, "姓名1", cache_key=f"{target_sheet_name}_姓名1") is not None
    
            # 如果没有"姓名1"列，在"预支数额"前插入
            if not has_name1:
                target_sheet.insert_cols(pre_advance_col)
                target_sheet.cell(row=2, column=pre_advance_col).value = "姓名1"
                # 清除相关缓存
                self._column_cache.clear()
    
            # 创建 Sheet3 和 Sheet4
            self._create_check_sheets()
    
            return True, f"检测完成！已确保'姓名1'列存在，并创建了核对工作表"
    
        except Exception as e:
            return False, f"检测失败：{str(e)}"

    def _create_check_sheets(self):
        """内部方法：创建工作表用于核对（根据当前轮次创建对应的Sheet）"""
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"
            
            names_num, data_num, result_num = self._get_round_sheet_numbers()
            names_sheet = f'Sheet{names_num}'
            data_sheet = f'Sheet{data_num}'
            result_sheet = f'Sheet{result_num}'

            # 创建Sheet
            if names_sheet not in self.workbook.sheetnames:
                self.workbook.create_sheet(names_sheet)
            if data_sheet not in self.workbook.sheetnames:
                self.workbook.create_sheet(data_sheet)
            if result_sheet not in self.workbook.sheetnames:
                self.workbook.create_sheet(result_sheet)

            return True, f"已创建 {names_sheet}、{data_sheet} 和 {result_sheet}"

        except Exception as e:
            return False, f"创建工作表失败：{str(e)}"


    def copy_names_from_workbook(self, source_worksheet, source_sheet_name):
        """从另一个工作簿的指定工作表复制姓名列到当前工作簿的Sheet3/6等（批量操作优化）"""
        try:
            if not self.workbook:
                return False, "请先选择主 Excel 文件"
            
            names_num, data_num, result_num = self._get_round_sheet_numbers()
            names_sheet_name = f'Sheet{names_num}'
            
            # 如果Sheet已存在，删除后重建（比逐行删除快）
            if names_sheet_name in self.workbook.sheetnames:
                del self.workbook[names_sheet_name]
            names_sheet = self.workbook.create_sheet(names_sheet_name)

            # 找到姓名列（从第二行开始识别）
            name_col = self._find_column_contains(source_worksheet, "姓名", cache_key=f"{source_sheet_name}_姓名")
            # 排除"姓名1"等变体
            if name_col:
                cell_value = str(source_worksheet.cell(row=2, column=name_col).value)
                if "1" in cell_value:
                    name_col = None
                    # 重新查找不包含数字的姓名列
                    for col in range(1, source_worksheet.max_column + 1):
                        cell_val = str(source_worksheet.cell(row=2, column=col).value)
                        if "姓名" in cell_val and "1" not in cell_val:
                            name_col = col
                            break
    
            if not name_col:
                return False, "未找到'姓名'列"
    
            # 批量读取姓名数据到列表
            max_row = source_worksheet.max_row
            names_data = []
            for row in range(3, max_row + 1):
                name_value = source_worksheet.cell(row=row, column=name_col).value
                if name_value is not None and str(name_value).strip() != "":
                    names_data.append(name_value)
            
            # 批量写入Sheet
            for idx, name_value in enumerate(names_data, start=1):
                names_sheet.cell(row=idx, column=1).value = name_value
    
            return True, f"已复制姓名数据到 {names_sheet_name}"
    
        except Exception as e:
            return False, f"复制姓名数据失败：{str(e)}"
    
    def apply_vlookup_formula(self, merged_sheet_name, selected_units=None):
        """在工作表中应用姓名匹配 - 使用Sheet3/6的姓名数据进行匹配，并筛选数据到Sheet4/7（批量优化）
        
        Args:
            merged_sheet_name: 要匹配的工作表名称
            selected_units: 用户选择的单位列表，None表示不筛选单位
        """
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"
            
            merged_sheet = self.workbook[merged_sheet_name]
            names_num, data_num, result_num = self._get_round_sheet_numbers()
            names_sheet_name = f'Sheet{names_num}'
            data_sheet_name = f'Sheet{data_num}'
            names_sheet = self.workbook[names_sheet_name]
            
            # 动态查找列
            name_col = self._find_column_by_header(merged_sheet, "姓名", cache_key=f"{merged_sheet_name}_姓名")
            if not name_col:
                return False, "未找到'姓名'列"
            
            # 查找单位列
            unit_col = self._find_column_by_header(merged_sheet, "单位", cache_key=f"{merged_sheet_name}_单位")
            
            balance_col = self._find_column_contains(merged_sheet, "平账额", cache_key=f"{merged_sheet_name}_平账额")
            if not balance_col:
                return False, "未找到'平账额'列"
            
            advance_col = self._find_column_contains(merged_sheet, "预支数额", cache_key=f"{merged_sheet_name}_预支数额")
            if not advance_col:
                return False, "未找到'预支数额'列"
            
            max_row = merged_sheet.max_row
            
            # 从Sheet3/6获取所有姓名数据（批量读取）
            sheet_names_set = set()
            for row in range(1, names_sheet.max_row + 1):
                name_value = names_sheet.cell(row=row, column=1).value
                if name_value and str(name_value).strip() != '':
                    sheet_names_set.add(str(name_value).strip())
            
            # 批量读取合并表格数据到内存
            rows_data = []
            matched_units = set()  # 收集所有匹配到的单位
            
            for row in range(3, max_row + 1):
                name_value = merged_sheet.cell(row=row, column=name_col).value
                balance_value = merged_sheet.cell(row=row, column=balance_col).value
                advance_value = merged_sheet.cell(row=row, column=advance_col).value
                unit_value = merged_sheet.cell(row=row, column=unit_col).value if unit_col else None
                
                rows_data.append({
                    'row': row,
                    'name': name_value,
                    'balance': balance_value,
                    'advance': advance_value,
                    'unit': unit_value
                })
                
                # 收集匹配到的单位
                if name_value and str(name_value).strip() in sheet_names_set:
                    if unit_value and str(unit_value).strip() != '':
                        matched_units.add(str(unit_value).strip())
            
            # 如果没有选择单位且有多个单位，返回单位列表供用户选择
            if selected_units is None and len(matched_units) > 1:
                return True, f"UNITS:{'|'.join(sorted(matched_units))}"
            
            # 在内存中处理匹配逻辑
            match_count = 0
            empty_balance_rows = []
            
            for data in rows_data:
                name_value = data['name']
                # 匹配姓名：检查姓名是否在Sheet3/6中
                if name_value and str(name_value).strip() in sheet_names_set:
                    # 如果指定了单位筛选，检查单位是否匹配
                    if selected_units:
                        unit_str = str(data['unit']).strip() if data['unit'] else ''
                        if unit_str not in selected_units:
                            continue
                    
                    match_count += 1
                    # 筛选有效行：姓名匹配成功 且 平账额为空
                    if data['balance'] is None or str(data['balance']).strip() == '':
                        empty_balance_rows.append(data)
            
            # 创建或清空Sheet4/7（删除后重建比逐行删除快）
            if data_sheet_name in self.workbook.sheetnames:
                del self.workbook[data_sheet_name]
            data_sheet = self.workbook.create_sheet(data_sheet_name)
            
            # 设置Sheet的标题
            data_sheet.cell(row=1, column=1).value = "姓名"
            data_sheet.cell(row=1, column=2).value = "预支数额"
            
            # 批量写入数据到Sheet
            for idx, data in enumerate(empty_balance_rows, start=2):
                data_sheet.cell(row=idx, column=1).value = data['name']
                data_sheet.cell(row=idx, column=2).value = data['advance']
            
            return True, f"已匹配 {match_count} 条数据，最终筛选 {len(empty_balance_rows)} 条数据到 {data_sheet_name}"
            
        except Exception as e:
            return False, f"应用姓名匹配失败：{str(e)}"
    
    def process_deduction_data(self, source_worksheet, source_sheet_name, selected_column_type=None):
        """处理需扣回和预支扣款数据并复制到 Sheet4（支持三列：需扣回、预支扣款、预支扣回）
            
        Args:
            source_worksheet: 源工作表
            source_sheet_name: 源工作表名称
            selected_column_type: 用户选择的列类型 ('需扣回', '预支扣款', '预支扣回')，None表示自动选择
        """
        try:
            if not self.workbook:
                return False, "请先选择主 Excel 文件"
            
            # 找到姓名列、需扣回列、预支扣款列和预支扣回列
            name_col = self._find_column_by_header(source_worksheet, "姓名", cache_key=f"{source_sheet_name}_姓名")
            deduction_col = self._find_column_contains(source_worksheet, "需扣回", cache_key=f"{source_sheet_name}_需扣回")
            advance_deduction_col = self._find_column_contains(source_worksheet, "预支扣款", cache_key=f"{source_sheet_name}_预支扣款")
            advance_return_col = self._find_column_contains(source_worksheet, "预支扣回", cache_key=f"{source_sheet_name}_预支扣回")
            
            if not name_col:
                return False, "未找到'姓名'列"
            
            if not deduction_col and not advance_deduction_col and not advance_return_col:
                return False, "未找到'需扣回'列、'预支扣款'列或'预支扣回'列"
            
            # 检测哪些列有有效数据
            available_columns = []
            if deduction_col:
                # 检查需扣回列是否有数据
                for row in range(3, source_worksheet.max_row + 1):
                    val = source_worksheet.cell(row=row, column=deduction_col).value
                    if val is not None and str(val).strip() != '':
                        available_columns.append(('需扣回', deduction_col))
                        break
            
            if advance_deduction_col:
                # 检查预支扣款列是否有数据
                for row in range(3, source_worksheet.max_row + 1):
                    val = source_worksheet.cell(row=row, column=advance_deduction_col).value
                    if val is not None and str(val).strip() != '':
                        available_columns.append(('预支扣款', advance_deduction_col))
                        break
            
            if advance_return_col:
                # 检查预支扣回列是否有数据
                for row in range(3, source_worksheet.max_row + 1):
                    val = source_worksheet.cell(row=row, column=advance_return_col).value
                    if val is not None and str(val).strip() != '':
                        available_columns.append(('预支扣回', advance_return_col))
                        break
            
            if not available_columns:
                return False, "未找到有效的需扣回或预支扣款数据"
            
            # 如果有多列且有用户选择，使用用户选择的列；否则使用第一列
            if len(available_columns) > 1 and selected_column_type:
                # 查找用户选择的列
                target_col = None
                for col_name, col_idx in available_columns:
                    if col_name == selected_column_type:
                        target_col = col_idx
                        break
                if not target_col:
                    return False, f"未找到'{selected_column_type}'列"
            else:
                # 自动选择第一列
                target_col = available_columns[0][1]
            
            # 筛选有效行：姓名有效且目标列有数据
            deduction_rows = []
            for row in range(3, source_worksheet.max_row + 1):
                name_value = source_worksheet.cell(row=row, column=name_col).value
                
                # 检查姓名是否有效
                if not name_value or str(name_value).strip() == '' or str(name_value).strip() == '总计':
                    continue
                
                # 获取目标列的值
                deduction_value = source_worksheet.cell(row=row, column=target_col).value
                
                # 如果有有效数据，则记录该行
                if deduction_value is not None and str(deduction_value).strip() != '':
                    deduction_rows.append({
                        'name': name_value,
                        'deduction': deduction_value
                    })
            
            if not deduction_rows:
                return False, "未找到有效的需扣回或预支扣款数据"
            
            # 获取当前轮次的Sheet编号
            names_num, data_num, result_num = self._get_round_sheet_numbers()
            data_sheet_name = f'Sheet{data_num}'
            
            # 保存用户选择的列类型到对应Sheet的 AZ1 单元格（用于平账时判断）
            if data_sheet_name not in self.workbook.sheetnames:
                return False, f"{data_sheet_name} 不存在，请先完成第一步"
            
            sheet4 = self.workbook[data_sheet_name]
            # 在 AZ1 保存列类型标记
            sheet4.cell(row=1, column=52).value = selected_column_type if selected_column_type else "预支扣款"
            
            # 定义动态列索引（避免硬编码）
            DEDUCTION_NAME_COL = 5
            DEDUCTION_VALUE_COL = 6
            ADVANCE_NAME_COL = 9
            ADVANCE_VALUE_COL = 10
            COMPARE_COL = 11
            SUM_COL = 12
            DEDUCTION_SUM_NAME_COL = 13
            DEDUCTION_SUM_VALUE_COL = 14
            
            # 设置第 5、6 列的标题（第6列标题使用用户选择的列名）
            sheet4.cell(row=1, column=DEDUCTION_NAME_COL).value = "姓名"
            # 使用用户选择的列名作为标题，如果没有选择则使用默认
            column_header = selected_column_type if selected_column_type else "需扣回/预支扣款"
            sheet4.cell(row=1, column=DEDUCTION_VALUE_COL).value = column_header
            
            # 复制数据到第 5、6 列
            for idx, data in enumerate(deduction_rows, start=2):
                sheet4.cell(row=idx, column=DEDUCTION_NAME_COL).value = data['name']
                sheet4.cell(row=idx, column=DEDUCTION_VALUE_COL).value = data['deduction']
            
            # 辅助函数：复制列数据并计算总计
            def copy_columns_with_total(src_name_col, src_value_col, dest_name_col, dest_value_col, title_name, title_value):
                """通用方法：复制两列数据并添加总计（不写入总计行，由外部统一处理）"""
                sheet4.cell(row=1, column=dest_name_col).value = title_name
                sheet4.cell(row=1, column=dest_value_col).value = title_value
                
                max_row_idx = 1
                total_value = 0
                
                for row in range(2, sheet4.max_row + 1):
                    name_value = sheet4.cell(row=row, column=src_name_col).value
                    value = sheet4.cell(row=row, column=src_value_col).value
                    
                    if name_value and str(name_value).strip() != '总计':
                        sheet4.cell(row=row, column=dest_name_col).value = name_value
                        sheet4.cell(row=row, column=dest_value_col).value = value
                        max_row_idx = row
                        
                        if value is not None:
                            try:
                                total_value += float(value)
                            except (ValueError, TypeError):
                                pass
                
                return max_row_idx, total_value
            
            # 对第1-2列数据进行姓名去重和求和，然后写入第9-10列
            # 先收集所有姓名和金额
            name_advance_dict = {}  # {"姓名": 总金额}
            for row in range(2, sheet4.max_row + 1):
                name_value = sheet4.cell(row=row, column=1).value
                advance_value = sheet4.cell(row=row, column=2).value
                
                if name_value and str(name_value).strip() != '' and str(name_value).strip() != '总计':
                    name_str = str(name_value).strip()
                    if name_str not in name_advance_dict:
                        name_advance_dict[name_str] = 0
                    
                    if advance_value is not None:
                        try:
                            name_advance_dict[name_str] += float(advance_value)
                        except (ValueError, TypeError):
                            pass
            
            # 复制第 5、6 列数据到第 13、14 列（使用上方已定义的 DEDUCTION_SUM_* 变量）
            max_row_13_14, total_deduction = copy_columns_with_total(
                5, 6, DEDUCTION_SUM_NAME_COL, DEDUCTION_SUM_VALUE_COL,
                "姓名", "求和项：需扣回/预支扣款"
            )
            
            # 根据第13列的姓名顺序，写入第9-10列（对齐第13列，使用上方已定义的 ADVANCE_* 变量）
            sheet4.cell(row=1, column=ADVANCE_NAME_COL).value = "姓名"
            sheet4.cell(row=1, column=ADVANCE_VALUE_COL).value = "求和项：预支数额"
            
            max_row_9_10 = 1
            total_advance = 0
            
            # 遍历第13列的姓名，按顺序写入第9列，并从字典中取金额写入第10列
            for row in range(2, sheet4.max_row + 1):
                name_13 = sheet4.cell(row=row, column=DEDUCTION_SUM_NAME_COL).value
                if name_13 and str(name_13).strip() != '' and str(name_13).strip() != '总计':
                    name_str = str(name_13).strip()
                    # 第9列：写入姓名（与第13列相同）
                    sheet4.cell(row=row, column=ADVANCE_NAME_COL).value = name_13
                    # 第10列：从求和字典中取值，如果没有则为0
                    advance_total = name_advance_dict.get(name_str, 0)
                    sheet4.cell(row=row, column=ADVANCE_VALUE_COL).value = advance_total
                    max_row_9_10 = row
                    total_advance += advance_total
            
            # 确定最大行数，确保总计行对齐
            max_data_row = max(max_row_9_10, max_row_13_14)
            total_row = max_data_row + 1
            
            # 在统一位置写入总计行
            if max_data_row > 1:
                sheet4.cell(row=total_row, column=ADVANCE_NAME_COL).value = "总计"
                sheet4.cell(row=total_row, column=ADVANCE_VALUE_COL).value = total_advance
                sheet4.cell(row=total_row, column=DEDUCTION_SUM_NAME_COL).value = "总计"
                sheet4.cell(row=total_row, column=DEDUCTION_SUM_VALUE_COL).value = total_deduction
            
            # 在第 11 列逐行比较第 9 列和第 13 列的姓名是否相同
            for row in range(2, total_row):
                name_9 = sheet4.cell(row=row, column=ADVANCE_NAME_COL).value
                name_13 = sheet4.cell(row=row, column=DEDUCTION_SUM_NAME_COL).value
                sheet4.cell(row=row, column=COMPARE_COL).value = (name_9 == name_13)
            
            # 在第 12 列逐行计算第 10 列和第 14 列的数值相加
            for row in range(2, total_row):
                val_10 = sheet4.cell(row=row, column=ADVANCE_VALUE_COL).value
                val_14 = sheet4.cell(row=row, column=DEDUCTION_SUM_VALUE_COL).value
                
                sum_value = 0
                if val_10 is not None:
                    try:
                        sum_value += float(val_10)
                    except (ValueError, TypeError):
                        pass
                if val_14 is not None:
                    try:
                        sum_value += float(val_14)
                    except (ValueError, TypeError):
                        pass
                
                sheet4.cell(row=row, column=SUM_COL).value = sum_value
            
            return True, f"已处理 {len(deduction_rows)} 条需扣回/预支扣款数据到 {data_sheet_name}"
            
        except Exception as e:
            return False, f"处理需扣回数据失败：{str(e)}"

    def balance_accounts(self, salary_sheet_name, month):
        """平账功能（批量优化）"""
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"

            salary_sheet = self.workbook[salary_sheet_name]
            
            # 获取当前轮次的Sheet编号
            names_num, data_num, result_num = self._get_round_sheet_numbers()
            data_sheet_name = f'Sheet{data_num}'
            result_sheet_name = f'Sheet{result_num}'
            
            if data_sheet_name not in self.workbook.sheetnames:
                return False, f"{data_sheet_name} 不存在，请先完成核对"
            
            sheet4 = self.workbook[data_sheet_name]

            # 动态查找相关列
            name_col = self._find_column_by_header(
                salary_sheet, "姓名", cache_key=f"{salary_sheet_name}_姓名")
            advance_col = self._find_column_contains(
                salary_sheet, "预支数额", cache_key=f"{salary_sheet_name}_预支数额")
            balance_col = self._find_column_contains(
                salary_sheet, "平账额", cache_key=f"{salary_sheet_name}_平账额")
            
            # 查找平账标记列（支持多种命名）
            settled_col = self._find_column_contains(
                salary_sheet, "是否平账", cache_key=f"{salary_sheet_name}_平账标记1")
            if not settled_col:
                settled_col = self._find_column_contains(
                    salary_sheet, "工资表是否平账", cache_key=f"{salary_sheet_name}_平账标记2")
            
            # 查找月份列（支持多种命名）
            month_col = self._find_column_contains(
                salary_sheet, "工资表所属月份", cache_key=f"{salary_sheet_name}_月份1")
            if not month_col:
                month_col = self._find_column_contains(
                    salary_sheet, "工资表月份", cache_key=f"{salary_sheet_name}_月份2")

            # 直接使用传入的月份
            if not month:
                month = "未知月份"

            # 读取 Sheet4 AZ1 单元格的标记，判断使用哪两列数据
            marker_col = 52  # AZ列
            column_type_marker = sheet4.cell(row=1, column=marker_col).value
            
            # 根据标记决定数据源列
            if column_type_marker == "需扣回":
                # 使用第5-6列的姓名去匹配第1-2列的数据
                source_name_col = 5  # 第5列：需扣回姓名
                source_value_col = 6  # 第6列：需扣回金额
                target_name_col = 1  # 第1列：原始姓名
                target_advance_col = 2  # 第2列：原始预支数额
                use_matching = True  # 需要匹配
            else:
                # 使用第1-2列（预支扣款/预支扣回）
                source_name_col = 1
                source_value_col = 2
                target_name_col = 1
                target_advance_col = 2
                use_matching = False  # 不需要匹配，直接使用
            
            # 读取 Sheet4 第12列（SUM_COL）的数据，用于判断是否需要拆分
            SUM_COL = 12  # 第12列
            DEDUCTION_SUM_NAME_COL = 13  # 第13列：汇总后的姓名
            split_data = {}  # {"姓名": sum_value}
            
            for row in range(2, sheet4.max_row + 1):
                name_13 = sheet4.cell(row=row, column=DEDUCTION_SUM_NAME_COL).value  # 第13列姓名（汇总后的）
                sum_value = sheet4.cell(row=row, column=SUM_COL).value  # 第12列求和值
                
                if name_13 and str(name_13).strip() != '' and str(name_13).strip() != '总计' and sum_value is not None:
                    try:
                        sum_num = float(sum_value)
                        name_str = str(name_13).strip()
                        
                        if sum_num > 0:
                            # 大于0：需要拆分
                            split_data[name_str] = sum_num
                    except (ValueError, TypeError):
                        pass
            
            print(f"[DEBUG] 需要拆分的姓名: {split_data}")
            
            # 获取 Sheet4 数据（支持同一姓名多条记录）
            sheet4_data = {}  # 格式：{"姓名": [{"advance": 预支数额, "row": 行号, "used": False}, ...]}
            print(f"[DEBUG] Sheet4 最大行数: {sheet4.max_row}")
            print(f"[DEBUG] 列类型标记: {column_type_marker}")
            print(f"[DEBUG] 是否使用匹配: {use_matching}")
            
            if use_matching:
                # 需扣回模式：先读取第5-6列的姓名，然后匹配第1-2列的数据
                # 步骤1：收集第5列的所有有效姓名
                deduction_names = set()
                for row in range(2, sheet4.max_row + 1):
                    name_5 = sheet4.cell(row=row, column=source_name_col).value
                    if name_5 and str(name_5).strip() != '' and str(name_5).strip() != '总计':
                        deduction_names.add(str(name_5).strip())
                
                print(f"[DEBUG] 第5列的有效姓名数: {len(deduction_names)}")
                
                # 步骤2：遍历第1-2列，只保留姓名在第5列中出现的记录
                for row in range(2, sheet4.max_row + 1):
                    name = sheet4.cell(row=row, column=target_name_col).value
                    advance = sheet4.cell(row=row, column=target_advance_col).value
                    
                    if name and str(name).strip() != '' and str(name).strip() != '总计' and advance is not None:
                        name_str = str(name).strip()
                        # 只有当姓名在第5列中出现时，才加入平账数据
                        if name_str in deduction_names:
                            if name_str not in sheet4_data:
                                sheet4_data[name_str] = []
                            sheet4_data[name_str].append({"advance": advance, "row": row, "used": False})
                            print(f"[DEBUG] Sheet4 第{row}行（匹配成功）: 姓名={name}, 预支数额={advance}")
            else:
                # 预支扣款/扣回模式：直接使用第1-2列
                for row in range(2, sheet4.max_row + 1):
                    name = sheet4.cell(row=row, column=source_name_col).value
                    advance = sheet4.cell(row=row, column=source_value_col).value
                    
                    if name and str(name).strip() != '' and str(name).strip() != '总计' and advance is not None:
                        name_str = str(name).strip()
                        if name_str not in sheet4_data:
                            sheet4_data[name_str] = []
                        sheet4_data[name_str].append({"advance": advance, "row": row, "used": False})
                        print(f"[DEBUG] Sheet4 第{row}行: 姓名={name}, 预支数额={advance}")
            
            print(f"[DEBUG] 最终 sheet4_data 条目数: {len(sheet4_data)}")
            
            # 批量读取工资表数据到内存（只读取平账额为空的行）
            salary_rows = []
            matched_count = 0
            empty_balance_count = 0
            for row in range(3, salary_sheet.max_row + 1):
                name_value = salary_sheet.cell(row=row, column=name_col).value
                advance_value = salary_sheet.cell(row=row, column=advance_col).value
                balance_value = salary_sheet.cell(row=row, column=balance_col).value
                
                # 只处理平账额为空的行
                if balance_value is None or str(balance_value).strip() == '':
                    salary_rows.append({
                        'row': row,
                        'name': name_value,
                        'advance': advance_value
                    })
                    empty_balance_count += 1
                    # 检查是否匹配（姓名存在于Sheet4）
                    if name_value:
                        name_str = str(name_value).strip()
                        if name_str in sheet4_data:
                            matched_count += 1
            print(f"[DEBUG] 平账额为空的行数: {empty_balance_count}, 匹配到的行数: {matched_count}")
            
            # 在内存中处理平账逻辑
            updated_count = 0
            # 使用 xxxx.x.x 格式的日期（无前置0，点分隔）
            now = datetime.now()
            today_date = f"{now.year}.{now.month}.{now.day}"  # 如2026.4.10
            updated_rows = []  # 记录成功平账的行号
            new_rows_to_add = []  # 记录需要新增的行数据（拆分产生的）
            no_balance_rows = set()  # 记录不平账的行号（不参与平账）
            
            # 第一步：处理需要拆分的姓名
            # 第12列的值表示"不需要平账的金额"，需要从工资表中拆分出来
                        
            for name_str, no_balance_amount in split_data.items():
                if name_str not in sheet4_data:
                    continue
                
                # 找到该姓名的所有工资表记录
                matching_salary_rows = [r for r in salary_rows if r['name'] and str(r['name']).strip() == name_str]
                
                if not matching_salary_rows:
                    continue
                
                # 计算该姓名的总预支金额
                total_advance = sum(float(r['advance']) if r['advance'] else 0 for r in matching_salary_rows)
                
                if total_advance <= 0:
                    continue
                
                # 实际需要平账的金额 = 总额 - 不需要平账的金额
                need_balance_amount = total_advance - no_balance_amount
                
                if need_balance_amount <= 0:
                    # 如果不需要平账，所有记录都标记为不平账
                    print(f"[DEBUG] {name_str} 无需平账，总额={total_advance}, 不平账金额={no_balance_amount}")
                    continue
                
                print(f"[DEBUG] 拆分准备: {name_str}, 总额={total_advance}, 不平账={no_balance_amount}, 需平账={need_balance_amount}")
                
                # 按金额从大到小排序
                matching_salary_rows.sort(key=lambda x: float(x['advance']) if x['advance'] else 0, reverse=True)
                
                remaining_no_balance = no_balance_amount  # 剩余需要标记为“不平账”的金额
                
                for salary_record in matching_salary_rows:
                    if remaining_no_balance <= 0:
                        break
                    
                    try:
                        original_advance = float(salary_record['advance']) if salary_record['advance'] else 0
                    except (ValueError, TypeError):
                        continue
                    
                    if original_advance <= 0:
                        continue
                    
                    # 判断当前记录是否可以作为“不平账”部分
                    if original_advance >= remaining_no_balance:
                        # 当前记录足够覆盖不平账金额
                        if abs(original_advance - remaining_no_balance) < 0.01:
                            # 刚好相等：整条记录作为不平账，保持原样不修改
                            # 标记为已处理，后续不参与平账
                            new_rows_to_add.append({
                                'source_row': salary_record['row'],
                                'advance': original_advance,
                                'is_split': True,
                                'should_balance': False,  # 不平账
                                'keep_original': True  # 保留原记录，不修改
                            })
                            # 记录不平账的行号，后续平账时跳过
                            no_balance_rows.add(salary_record['row'])
                            print(f"[DEBUG] 完全不平账: {name_str}, 金额={original_advance}（保留原记录）")
                        else:
                            # 需要拆分：一部分不平账，一部分平账
                            no_balance_part = remaining_no_balance
                            balance_part = original_advance - remaining_no_balance
                            
                            # 修改原记录为平账部分
                            salary_sheet.cell(row=salary_record['row'], column=advance_col).value = balance_part
                            
                            # 新增不平账部分
                            new_rows_to_add.append({
                                'source_row': salary_record['row'],
                                'advance': no_balance_part,
                                'is_split': True,
                                'should_balance': False  # 不平账
                            })
                            
                            print(f"[DEBUG] 拆分: {name_str}, 原金额={original_advance}, 不平账={no_balance_part}, 平账={balance_part}")
                        
                        remaining_no_balance = 0
                        break
                    else:
                        # 当前记录全部作为不平账部分，保持原样不修改
                        new_rows_to_add.append({
                            'source_row': salary_record['row'],
                            'advance': original_advance,
                            'is_split': True,
                            'should_balance': False,  # 不平账
                            'keep_original': True  # 保留原记录，不修改
                        })
                        # 记录不平账的行号，后续平账时跳过
                        no_balance_rows.add(salary_record['row'])
                        remaining_no_balance -= original_advance
                        print(f"[DEBUG] 部分不平账: {name_str}, 金额={original_advance}, 剩余不平账={remaining_no_balance}")
            
            # 第二步：正常平账逻辑（处理未拆分的记录）
            for data in salary_rows:
                # 跳过不平账的行
                if data['row'] in no_balance_rows:
                    print(f"[DEBUG] 跳过不平账行: {data['name']}, 金额={data['advance']}, 行号={data['row']}")
                    continue
                
                if data['name']:
                    name_str = str(data['name']).strip()
                    if name_str in sheet4_data:
                        records = sheet4_data[name_str]
                        
                        # 遍历该姓名的所有预支数额记录，找到金额匹配的
                        for record in records:
                            sheet4_advance = record["advance"]
                            print(f"[DEBUG] 找到匹配: {name_str}, 工资表预支={data['advance']}, Sheet4预支={sheet4_advance}, 已使用={record['used']}")
                            
                            # 如果这条记录已经被使用过，跳过
                            if record["used"]:
                                print(f"[DEBUG] 记录已使用，跳过: {name_str}, Sheet4预支={sheet4_advance}")
                                continue
                            
                            try:
                                salary_advance = float(data['advance']) if data['advance'] is not None else None
                                sheet4_advance_num = float(sheet4_advance) if sheet4_advance is not None else None
                                
                                if salary_advance is not None and sheet4_advance_num is not None:
                                    if abs(salary_advance - sheet4_advance_num) < 0.01:
                                        # 标记该记录已使用
                                        record["used"] = True
                                        # 批量更新单元格
                                        salary_sheet.cell(row=data['row'], column=balance_col).value = data['advance']
                                        if settled_col:
                                            salary_sheet.cell(row=data['row'], column=settled_col).value = today_date
                                        if month_col:
                                            salary_sheet.cell(row=data['row'], column=month_col).value = month
                                        updated_count += 1
                                        updated_rows.append(data['row'])
                                        print(f"[DEBUG] ✓ 平账成功: {name_str}, 金额={salary_advance}")
                                        # 找到匹配后立即跳出，避免重复处理
                                        break
                                    else:
                                        print(f"[DEBUG] 金额不匹配: {name_str}, 工资表={salary_advance}, Sheet4={sheet4_advance_num}")
                            except (ValueError, TypeError) as e:
                                print(f"[DEBUG] 转换错误: {name_str}, 错误: {e}")
                                pass
            
            # 第四步：添加拆分产生的新行（优化：批量添加到末尾，避免逐行插入）
            if new_rows_to_add:
                # 收集所有需要添加的行数据
                rows_to_append = []
                
                for new_row_data in new_rows_to_add:
                    # 如果标记为保留原记录，则不需要添加新行（原记录已经在工资表中）
                    if new_row_data.get('keep_original', False):
                        print(f"[DEBUG] 跳过添加: 行{new_row_data['source_row']}（保留原记录）")
                        continue
                    
                    source_row = new_row_data['source_row']
                    new_advance = new_row_data['advance']
                    should_balance = new_row_data.get('should_balance', False)
                    
                    # 读取源行的完整数据
                    row_data = {}
                    for col in range(1, salary_sheet.max_column + 1):
                        cell_value = salary_sheet.cell(row=source_row, column=col).value
                        row_data[col] = cell_value
                    
                    # 修改预支数额
                    row_data[advance_col] = new_advance
                    
                    # 确保姓名正确
                    source_name = row_data.get(name_col)
                    
                    # 如果不平账，清空相关列
                    if not should_balance:
                        if balance_col:
                            row_data[balance_col] = None
                        if settled_col:
                            row_data[settled_col] = None
                        if month_col:
                            row_data[month_col] = None
                    
                    rows_to_append.append({
                        'data': row_data,
                        'source_row': source_row,
                        'should_balance': should_balance,
                        'name': source_name
                    })

                # 注意：不能直接用 max_row，因为 openpyxl 的 max_row 包含曾使用但已清空的"幽灵行"
                # 需要向上扫描找到真正有数据的最后一行
                raw_max = salary_sheet.max_row
                actual_last_row = raw_max
                for check_row in range(raw_max, 0, -1):
                    # 检查该行是否有任何非空单元格
                    row_has_data = False
                    for col in range(1, min(salary_sheet.max_column + 1, 20)):  # 只检查前20列，覆盖主要数据区
                        val = salary_sheet.cell(row=check_row, column=col).value
                        if val is not None and str(val).strip() != '':
                            row_has_data = True
                            break
                    if row_has_data:
                        actual_last_row = check_row
                        break
                
                print(f"[DEBUG] 原始max_row={raw_max}, 实际最后数据行={actual_last_row}, 将从{actual_last_row + 1}开始追加")
                
                for idx, row_info in enumerate(rows_to_append, start=1):
                    target_row = actual_last_row + idx
                    row_data = row_info['data']
                    
                    for col, value in row_data.items():
                        salary_sheet.cell(row=target_row, column=col).value = value
                    
                    status = "不平账" if not row_info['should_balance'] else "平账"
                    print(f"[DEBUG] 新增{status}行: 行{target_row}, 姓名={row_info['name']}, 金额={row_data[advance_col]}")
            
            # 第五步：为平账成功的行标红（排除拆分新增的行）
            if updated_rows:
                red_font = Font(color="FF0000")  # 红色字体
                for row_num in updated_rows:
                    for col in range(1, salary_sheet.max_column + 1):
                        cell = salary_sheet.cell(row=row_num, column=col)
                        # 只标记有内容的单元格
                        if cell.value is not None and str(cell.value).strip() != '':
                            cell.font = red_font

            # 创建或清空 Sheet5/8，保存平账记录（完整行数据）
            if result_sheet_name in self.workbook.sheetnames:
                del self.workbook[result_sheet_name]
            sheet5 = self.workbook.create_sheet(result_sheet_name)
            
            # 复制工资表的标题行到 Sheet5
            for col in range(1, salary_sheet.max_column + 1):
                header_value = salary_sheet.cell(row=2, column=col).value
                if header_value:
                    sheet5.cell(row=1, column=col).value = header_value
            
            # 写入平账成功的完整行数据
            for idx, row_num in enumerate(updated_rows, start=2):
                for col in range(1, salary_sheet.max_column + 1):
                    cell_value = salary_sheet.cell(row=row_num, column=col).value
                    sheet5.cell(row=idx, column=col).value = cell_value

            return True, f"已平账{updated_count}条记录，月份：{month}，数据已保存到 {result_sheet_name}"

        except Exception as e:
            return False, f"平账失败：{str(e)}"

    def save_workbook(self, save_path=None):
        """保存工作簿"""
        try:
            if not self.workbook:
                return False, "没有工作簿可保存"

            # 如果没有提供保存路径，使用原文件路径
            if save_path is None:
                if not self.current_file_path:
                    return False, "请先选择保存位置"
                save_path = self.current_file_path
                # 不再重新加载工作簿，直接使用内存中的工作簿（包含所有修改）
                # self.workbook = load_workbook(self.current_file_path, data_only=True)

            self.workbook.save(save_path)
            return True, f"文件已保存到：{save_path}"

        except Exception as e:
            return False, f"保存失败：{str(e)}"


# 创建全局实例
excel_processor = ExcelProcessor()