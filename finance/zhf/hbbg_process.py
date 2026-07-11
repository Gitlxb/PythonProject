# -*- coding: utf-8 -*-
"""
处理模式核心引擎

从 Excel 文件中提取月份信息，在目标列填写月份，然后合并为汇总表。
与 GUI 完全解耦：通过回调函数报告进度/状态，不直接操控 Tkinter 组件。

用法:
    from hbbg_process import ProcessMergeEngine

    engine = ProcessMergeEngine(
        target_col_letter="AU",
        progress_callback=lambda pct, text: ...,
    )
    result = engine.process_single_file(file_path)
    engine.merge_processed_files(processed_paths, output_path)
    engine.cleanup()
"""
import os
import re
import traceback

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from copy import copy


class ProcessMergeEngine:
    """处理模式：填写月份列 + 合并汇总"""

    def __init__(self, target_col_letter="", progress_callback=None, status_callback=None):
        """
        参数:
            target_col_letter: 目标列字母，如 "AU"
            progress_callback: Callable(pct: float, text: str) 进度回调
            status_callback:   Callable(text: str) 状态栏回调
        """
        self.target_col_letter = target_col_letter.strip().upper()
        self._progress_cb = progress_callback or (lambda p, t: None)
        self._status_cb = status_callback or (lambda t: None)
        self._converted_temps = []   # xlrd 转换产生的临时文件
        self.process_results = []    # 每次批处理的结果列表

    # ================================================================
    #  静态工具方法
    # ================================================================

    @staticmethod
    def extract_month_from_date_range(date_str):
        """从 "2024年10月01日-2024年10月31日" 提取月份数字"""
        try:
            match = re.search(r'(\d{4})年(\d{1,2})月', str(date_str))
            if match:
                return int(match.group(2))
        except Exception as e:
            print(f"解析日期出错: {e}")
        return None

    @staticmethod
    def col_letter_to_index(col_letter):
        """列字母 → 1-based 列索引（A→1, J→10, AA→27）"""
        col_letter = col_letter.upper().strip()
        result = 0
        for char in col_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    @staticmethod
    def get_actual_max_row(ws):
        """检测实际有数据的最后一行，避免 max_row 因格式虚高"""
        max_row = ws.max_row
        max_col = ws.max_column
        for row_idx in range(max_row, 0, -1):
            for col_idx in range(1, max_col + 1):
                if ws.cell(row=row_idx, column=col_idx).value is not None:
                    return row_idx
        return 0

    @staticmethod
    def build_result_summary(success, skip, error, output_path, results=None):
        """构建处理结果摘要字符串"""
        lines = [
            "=" * 40,
            f"  总计: {success + skip + error} 个文件",
            f"  ✓ 成功处理: {success}",
            f"  - 已跳过:   {skip}",
            f"  ✗ 处理错误: {error}",
            "=" * 40,
            f"\n保存位置:\n{output_path}"
        ]
        if results:
            detail_lines = ["\n详细信息:"]
            for i, r in enumerate(results[:10]):
                fn = os.path.basename(r["file"])
                if r["status"] == "success":
                    detail_lines.append(
                        f"  ✓ {fn} → 列:{r.get('target_col','?')} 月份:{r.get('month','')} "
                        f"(第{r.get('start_row','?')}~{r.get('end_row','?')}行, 共{r.get('filled',0)}行)")
                elif r["status"] == "skip":
                    detail_lines.append(f"  - {fn}: {r.get('message','已跳过')}")
                else:
                    detail_lines.append(f"  ✗ {fn}: {r.get('message','未知错误')}")
            if len(results) > 10:
                detail_lines.append(f"  ... 还有 {len(results) - 10} 条记录")
            lines.extend(detail_lines)
        return "\n".join(lines)

    # ================================================================
    #  文件加载
    # ================================================================

    def _load_with_openpyxl(self, file_path):
        """加载 Excel：优先 openpyxl，失败则 xlrd 转存后加载"""
        filename = os.path.basename(file_path)
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            data = [[cell.value for cell in row] for row in ws.iter_rows()]
            print(f"{filename}: openpyxl 直接加载成功 ({ws.max_row}行)")
            return wb, ws, data
        except Exception:
            pass

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
            self._converted_temps.append(temp_converted)
            print(f"{filename}: xlrd 转换后加载 ({wb.active.max_row}行)")
            return wb, ws, data
        except Exception as e:
            print(f"{filename}: 所有方式加载失败 - {e}")
            return None, None, None

    # ================================================================
    #  核心处理逻辑
    # ================================================================

    def process_single_file(self, file_path):
        """
        处理单个文件：查找日期→计算月份→在目标列填入月份。

        返回: dict {"status": "success"|"skip"|"error", "file": ..., "message": ..., ...}
        """
        if not self.target_col_letter:
            return {"status": "error", "file": file_path,
                    "message": "未指定目标列，请在输入框中输入列字母(如 AU)"}
        try:
            target_excel_col = self.col_letter_to_index(self.target_col_letter)
        except Exception as e:
            return {"status": "error", "file": file_path,
                    "message": f"目标列 '{self.target_col_letter}' 格式错误: {e}"}

        try:
            wb, ws, data = self._load_with_openpyxl(file_path)
            if wb is None:
                return {"status": "error", "file": file_path, "message": "无法读取文件"}

            total_rows = len(data)
            filename = os.path.basename(file_path)
            if total_rows < 5:
                wb.close()
                return {"status": "error", "file": file_path,
                        "message": f"数据不足5行（实际{total_rows}行），找不到表头"}

            # 第2行获取时间范围
            date_cell_value = None
            for col in range(len(data[1])):
                cell_value = str(data[1][col]) if pd.notna(data[1][col]) else ""
                if "年" in cell_value and "月" in cell_value and "日" in cell_value:
                    date_cell_value = cell_value
                    break
            if not date_cell_value:
                wb.close()
                return {"status": "skip", "file": file_path, "message": "未在第2行找到日期信息"}

            current_month = self.extract_month_from_date_range(date_cell_value)
            if current_month is None:
                wb.close()
                return {"status": "skip", "file": file_path, "message": "无法解析月份"}

            fill_month = 12 if current_month == 1 else current_month - 1
            month_str = f"{fill_month}月"

            # 第5行找"身份证件类型"列
            id_type_col_idx = None
            for col in range(len(data[4])):
                cell_value = str(data[4][col]) if pd.notna(data[4][col]) else ""
                if cell_value.strip() == "身份证件类型":
                    id_type_col_idx = col
                    break
            if id_type_col_idx is None:
                wb.close()
                return {"status": "skip", "file": file_path, "message": "未在第5行找到'身份证件类型'列"}

            # 找到"居民身份证"的起止行
            start_row_data_idx = None
            end_row_data_idx = None
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

            print(f"{filename}: 目标列={self.target_col_letter}(第{target_excel_col}列) | "
                  f"范围=第{start_row_data_idx+1}行~第{end_row_data_idx+1}行 | 月份={month_str}")

            # 填写月份
            filled_count = 0
            for data_row_idx in range(start_row_data_idx, end_row_data_idx + 1):
                excel_row = data_row_idx + 1
                target_cell = ws.cell(row=excel_row, column=target_excel_col)
                target_cell.value = month_str
                filled_count += 1

            print(f"{filename}: 完成 | 填写={filled_count}行")

            base_name = os.path.splitext(filename)[0]
            temp_path = os.path.join(os.path.dirname(file_path), f'{base_name}_processed.xlsx')
            wb.save(temp_path)
            wb.close()

            return {
                "status": "success", "file": file_path, "temp_path": temp_path,
                "month": month_str, "filled": filled_count,
                "target_col": self.target_col_letter,
                "start_row": start_row_data_idx + 1, "end_row": end_row_data_idx + 1
            }

        except Exception as e:
            traceback.print_exc()
            return {"status": "error", "file": file_path, "message": str(e)}

    # ================================================================
    #  合并汇总
    # ================================================================

    def merge_processed_files(self, file_paths, output_path):
        """
        用实际数据范围替代 max_row，避免空行过多。
        将已处理的文件（或跳过的原文件）合并到一个汇总工作簿中。
        """
        if not file_paths:
            return

        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = "汇总表"
        current_row = 1
        total_files = len(file_paths)

        for idx, file_path in enumerate(file_paths):
            try:
                source_wb = load_workbook(file_path)
                source_ws = source_wb.active

                actual_max_row = self.get_actual_max_row(source_ws)
                max_col = source_ws.max_column

                if actual_max_row == 0:
                    source_wb.close()
                    continue

                # 复制内容（行高 + 值 + 样式）
                for row_idx in range(1, actual_max_row + 1):
                    src_rd = source_ws.row_dimensions[row_idx]
                    if src_rd.height:
                        new_ws.row_dimensions[current_row].height = src_rd.height

                    for col_idx in range(1, max_col + 1):
                        source_cell = source_ws.cell(row=row_idx, column=col_idx)
                        target_cell = new_ws.cell(row=current_row, column=col_idx)
                        target_cell.value = source_cell.value
                        if source_cell.has_style:
                            target_cell.font = copy(source_cell.font)
                            target_cell.border = copy(source_cell.border)
                            target_cell.fill = copy(source_cell.fill)
                            target_cell.number_format = source_cell.number_format
                            target_cell.protection = copy(source_cell.protection)
                            target_cell.alignment = copy(source_cell.alignment)

                    current_row += 1

                # 复制合并单元格
                for merged_range in source_ws.merged_cells.ranges:
                    min_col, min_row, max_col_m, max_row_m = merged_range.bounds
                    if min_row > actual_max_row:
                        continue
                    max_row_m_clamped = min(max_row_m, actual_max_row)
                    offset_min_row = min_row + (current_row - actual_max_row) - 1
                    offset_max_row = max_row_m_clamped + (current_row - actual_max_row) - 1
                    new_merged = f"{get_column_letter(min_col)}{offset_min_row}:{get_column_letter(max_col_m)}{offset_max_row}"
                    try:
                        new_ws.merge_cells(new_merged)
                    except Exception:
                        pass

                # 复制列宽
                for col_idx in range(1, max_col + 1):
                    col_letter = get_column_letter(col_idx)
                    if col_letter in source_ws.column_dimensions:
                        src_cd = source_ws.column_dimensions[col_letter]
                        if src_cd.width:
                            new_ws.column_dimensions[col_letter].width = src_cd.width

                source_wb.close()

                # 文件间空20行
                if idx < total_files - 1:
                    current_row += 20

                print(f"已合并 [{idx+1}/{total_files}]: {os.path.basename(file_path)} ({actual_max_row}行)")

            except Exception as e:
                print(f"合并失败: {os.path.basename(file_path)} - {e}")
                traceback.print_exc()
                continue

        new_wb.save(output_path)
        new_wb.close()
        print(f"\n✓ 汇总表已保存: {output_path}")

    # ================================================================
    #  清理
    # ================================================================

    def cleanup(self, temp_files=None):
        """清理临时文件"""
        if temp_files:
            for temp_file in temp_files:
                try:
                    if temp_file and '_processed.xlsx' in temp_file and os.path.exists(temp_file):
                        os.remove(temp_file)
                        print(f"已清理临时文件: {os.path.basename(temp_file)}")
                except Exception as e:
                    print(f"清理失败: {temp_file} - {e}")

        for conv_temp in self._converted_temps:
            try:
                if os.path.exists(conv_temp):
                    os.remove(conv_temp)
                    print(f"已清理转换临时文件: {os.path.basename(conv_temp)}")
            except Exception as e:
                print(f"清理转换临时文件失败: {conv_temp} - {e}")
        self._converted_temps = []
