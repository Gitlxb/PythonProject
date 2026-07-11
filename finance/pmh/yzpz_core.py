from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import warnings
import gc

warnings.filterwarnings('ignore')


class ExcelProcessor:
    def __init__(self):
        self.workbook = None
        self.sheet_names = []
        self.current_file_path = None
        self._column_cache = {}  # 列位置缓存
        self.check_round = 1  # 核对轮次计数器

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
                self.check_round = 1
                gc.collect()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_workbook()
        return False

    @staticmethod
    def _is_valid_name(name_value):
        """检查姓名是否有效（非空且不是'总计'）"""
        if not name_value:
            return False
        name_str = str(name_value).strip()
        return name_str != '' and name_str != '总计'

    @staticmethod
    def _to_float(value, default=0):
        """安全地将值转换为浮点数"""
        if value is None:
            return default
        try:
            return float(value)
        except (ValueError, TypeError):
            return default

    def _get_round_sheet_numbers(self):
        """根据当前轮次获取Sheet编号"""
        base = 3 + (self.check_round - 1) * 2
        return base, base + 1

    def increment_round(self):
        """增加轮次计数器"""
        if self.check_round >= 30:
            raise ValueError("轮次过多（已达30轮），建议保存文件后重新开始")
        self.check_round += 1
        self._column_cache.clear()
        return self.check_round

    def get_current_check_round(self):
        """获取当前核对轮次"""
        return self.check_round

    def select_excel_file(self, file_path):
        """选择 Excel 文件并读取工作表"""
        try:
            self.close_workbook()
            self.workbook = load_workbook(file_path, data_only=True)
            self.sheet_names = self.workbook.sheetnames
            self.current_file_path = file_path
            self._column_cache.clear()
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
        """检测功能 - 创建核对工作表"""
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"

            target_sheet = self.workbook[target_sheet_name]
            self._create_check_sheets()
            return True, "检测完成！已创建核对工作表"
        except Exception as e:
            return False, f"检测失败：{str(e)}"

    def _create_check_sheets(self):
        """内部方法：创建工作表用于核对（根据当前轮次创建对应的Sheet）"""
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"

            names_num, data_num = self._get_round_sheet_numbers()
            names_sheet = f'Sheet{names_num}'
            data_sheet = f'Sheet{data_num}'

            if names_sheet not in self.workbook.sheetnames:
                self.workbook.create_sheet(names_sheet)
            if data_sheet not in self.workbook.sheetnames:
                self.workbook.create_sheet(data_sheet)

            return True, f"已创建 {names_sheet} 和 {data_sheet}"
        except Exception as e:
            return False, f"创建工作表失败：{str(e)}"

    def save_workbook(self, save_path=None):
        """保存工作簿"""
        try:
            if not self.workbook:
                return False, "没有工作簿可保存"

            if save_path is None:
                if not self.current_file_path:
                    return False, "请先选择保存位置"
                save_path = self.current_file_path

            self.workbook.save(save_path)
            return True, f"文件已保存到：{save_path}"
        except Exception as e:
            return False, f"保存失败：{str(e)}"

    # ==================== 核对方法 ====================

    def copy_names_from_workbook(self, source_worksheet, source_sheet_name):
        """从另一个工作簿的指定工作表复制姓名列到当前工作簿的Sheet3/6等（批量操作优化）"""
        try:
            if not self.workbook:
                return False, "请先选择 Excel 文件"

            names_num, data_num = self._get_round_sheet_numbers()
            names_sheet_name = f'Sheet{names_num}'

            if names_sheet_name in self.workbook.sheetnames:
                del self.workbook[names_sheet_name]
            names_sheet = self.workbook.create_sheet(names_sheet_name)

            name_col = self._find_column_contains(source_worksheet, "姓名", cache_key=f"{source_sheet_name}_姓名")

            if not name_col:
                return False, "未找到'姓名'列"

            max_row = source_worksheet.max_row
            names_data = []
            for row in range(3, max_row + 1):
                name_value = source_worksheet.cell(row=row, column=name_col).value
                if name_value is not None and str(name_value).strip() != '':
                    names_data.append(name_value)

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
            names_num, data_num = self._get_round_sheet_numbers()
            names_sheet_name = f'Sheet{names_num}'
            data_sheet_name = f'Sheet{data_num}'
            names_sheet = self.workbook[names_sheet_name]

            name_col = self._find_column_by_header(merged_sheet, "姓名", cache_key=f"{merged_sheet_name}_姓名")
            if not name_col:
                return False, "未找到'姓名'列"

            unit_col = self._find_column_by_header(merged_sheet, "单位", cache_key=f"{merged_sheet_name}_单位")
            balance_col = self._find_column_contains(merged_sheet, "平账额", cache_key=f"{merged_sheet_name}_平账额")
            if not balance_col:
                return False, "未找到'平账额'列"

            advance_col = self._find_column_contains(merged_sheet, "预支数额",
                                                     cache_key=f"{merged_sheet_name}_预支数额")
            if not advance_col:
                return False, "未找到'预支数额'列"

            max_row = merged_sheet.max_row

            sheet_names_set = set()
            for row in range(1, names_sheet.max_row + 1):
                name_value = names_sheet.cell(row=row, column=1).value
                if name_value and str(name_value).strip() != '':
                    sheet_names_set.add(str(name_value).strip())

            rows_data = []
            matched_units = set()

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

                if selected_units is None and name_value and str(name_value).strip() in sheet_names_set:
                    if unit_value and str(unit_value).strip() != '':
                        matched_units.add(str(unit_value).strip())

            if selected_units is None and len(matched_units) > 1:
                return True, f"UNITS:{'|'.join(sorted(matched_units))}"

            match_count = 0
            empty_balance_rows = []

            for data in rows_data:
                name_value = data['name']
                if name_value and str(name_value).strip() in sheet_names_set:
                    if selected_units and unit_col:
                        unit_str = str(data['unit']).strip() if data['unit'] else ''
                        if unit_str not in selected_units:
                            continue

                    match_count += 1
                    if data['balance'] is None or str(data['balance']).strip() == '':
                        empty_balance_rows.append(data)

            if data_sheet_name in self.workbook.sheetnames:
                del self.workbook[data_sheet_name]
            data_sheet = self.workbook.create_sheet(data_sheet_name)

            data_sheet.cell(row=1, column=1).value = "姓名"
            data_sheet.cell(row=1, column=2).value = "预支数额"

            for idx, data in enumerate(empty_balance_rows, start=2):
                data_sheet.cell(row=idx, column=1).value = data['name']
                data_sheet.cell(row=idx, column=2).value = data['advance']

            return True, f"已匹配 {match_count} 条数据，最终筛选 {len(empty_balance_rows)} 条数据到 {data_sheet_name}"

        except Exception as e:
            return False, f"应用姓名匹配失败：{str(e)}"


# 全局实例，供 GUI 导入使用
excel_processor = ExcelProcessor()

# 导入 yzpz_logic 模块，将 process_deduction_data 和 balance_accounts 方法绑定到 ExcelProcessor 类
from finance.pmh import yzpz_logic
