# -- coding: utf-8 --
# @File : bijiao_frame.py
# @Description: 数据比较工具 - 核心工具函数（日期解析、表头查找、结果合并、Excel保存）

from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd
import numpy as np


# ==================== Excel保存功能 ====================

def save_to_excel(final_df, save_path):
    """将结果保存到Excel文件"""
    if final_df is None or final_df.empty:
        raise ValueError("没有数据可以保存")

    wb = Workbook()
    ws = wb.active
    ws.title = "对比结果"

    # 样式定义
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    true_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    false_fill = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")

    data_font = Font(name="微软雅黑", size=10)
    data_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # 表头
    headers = ['姓名', '预支金额', '身份证号', '比较结果', '差额', '姓名', '预支数额', '备注', '运营中心']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

    # 数据行
    for row_idx, (_, row_data) in enumerate(final_df.iterrows(), 2):
        values = [
            str(row_data['姓名']),
            float(row_data['预支金额']),
            str(row_data['身份证号']),
            str(row_data['比较结果']),
            float(row_data['差额']),
            str(row_data['姓名(右)']),
            float(row_data['预支数额(右)']),
            str(row_data.get('备注', '')),
            str(row_data.get('运营中心', ''))
        ]

        is_true = row_data['比较结果'] == 'TRUE'

        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border
            cell.fill = true_fill if is_true else false_fill

    # 列宽
    col_widths_map = {1: 14, 2: 14, 3: 22, 4: 12, 5: 12, 6: 14, 7: 14, 8: 20, 9: 20}
    for col_idx, width in col_widths_map.items():
        ws.column_dimensions[chr(64 + col_idx)].width = width

    # 总计行（紧跟数据行下方）
    total_row = len(final_df) + 2
    total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_font = Font(name="微软雅黑", size=10, bold=True)
    total_left = float(final_df['预支金额'].sum())
    total_diff = float(final_df['差额'].sum())
    total_right = float(final_df['预支数额(右)'].sum())

    total_cells = [
        (1, "总计"),
        (2, total_left),
        (5, total_diff),
        (7, total_right),
    ]
    for col_idx, value in total_cells:
        cell = ws.cell(row=total_row, column=col_idx, value=value)
        cell.font = total_font
        cell.alignment = data_alignment
        cell.border = thin_border
        cell.fill = total_fill

    # 统计信息（总计行下方空一行开始）
    stat_row = total_row + 2
    ws.cell(row=stat_row, column=1, value="统计信息").font = Font(name="微软雅黑", size=10, bold=True)

    true_count = len(final_df[final_df['比较结果'] == 'TRUE'])
    false_count = len(final_df[final_df['比较结果'] == 'FALSE'])

    ws.cell(row=stat_row + 1, column=1, value="相符（TRUE）").font = data_font
    ws.cell(row=stat_row + 1, column=2, value=true_count).font = data_font
    ws.cell(row=stat_row + 2, column=1, value="不符（FALSE）").font = data_font
    ws.cell(row=stat_row + 2, column=2, value=false_count).font = data_font
    ws.cell(row=stat_row + 3, column=1, value="总计").font = data_font
    ws.cell(row=stat_row + 3, column=2, value=len(final_df)).font = data_font

    wb.save(save_path)
    wb.close()


# ==================== 工具函数 ====================

def format_number(value):
    """格式化数字"""
    try:
        num = float(value)
        if num == int(num):
            return str(int(num))
        return f"{num:.2f}"
    except (ValueError, TypeError):
        return str(value)


def remark_by_diff(diff):
    """根据差额生成备注"""
    if diff > 0:
        return "申请表多报了"
    elif diff < 0:
        return "明细表有但没申请"
    return ""


# ==================== 日期解析和表头查找功能 ====================

def parse_date_flexible(series):
    """更灵活地解析日期，兼容多种格式，统一转为日期（不含时间）"""
    # 第一步：如果已经是 datetime 类型，直接 normalize
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors='coerce').dt.normalize()

    # 第二步：先用 pandas 向量化批量解析（最稳定、最快）
    batch_parsed = pd.to_datetime(series, errors='coerce')
    if batch_parsed.notna().any():
        return batch_parsed.dt.normalize()

    # 第三步：对批量解析失败的值，逐元素兜底
    def _try_convert(val):
        if pd.isna(val) or str(val).strip() in ('', 'nan', 'NaT', 'None'):
            return pd.NaT

        # 已经是 datetime / Timestamp / np.datetime64
        if isinstance(val, (pd.Timestamp, datetime)):
            return pd.Timestamp(val).normalize()
        if isinstance(val, np.datetime64):
            return pd.Timestamp(val).normalize()

        str_val_orig = str(val).strip()

        # pandas 直接解析
        try:
            result = pd.to_datetime(str_val_orig, errors='coerce')
            if pd.notna(result):
                return result.normalize()
        except Exception:
            pass

        # dayfirst
        try:
            result = pd.to_datetime(str_val_orig, errors='coerce', dayfirst=True)
            if pd.notna(result):
                return result.normalize()
        except Exception:
            pass

        # Excel 日期序列号
        try:
            num_val = float(str_val_orig)
            if 1 < num_val < 100000:
                return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=int(num_val))).normalize()
        except (ValueError, TypeError):
            pass

        # 中文格式
        str_val = str_val_orig.replace('年', '-').replace('月', '-').replace('日', '')
        try:
            result = pd.to_datetime(str_val, errors='coerce')
            if pd.notna(result):
                return result.normalize()
        except Exception:
            pass

        # 斜杠格式（兼容单数字月日）
        if '/' in str_val_orig:
            for fmt in ['%Y/%m/%d %H:%M:%S', '%Y/%m/%d', '%Y/%m/%d %H:%M',
                        '%Y/%m/%d %I:%M:%S %p']:
                try:
                    return pd.Timestamp(datetime.strptime(str_val_orig, fmt)).normalize()
                except ValueError:
                    continue

        return pd.NaT

    return series.apply(_try_convert)


def find_header_row(df, required_cols, search_rows=10):
    """
    在DataFrame的前几行中查找表头行
    返回：(header_row_index, column_mapping)
    column_mapping: {标准列名: 实际列索引}
    """
    df_str = df.astype(str)

    for row_idx in range(min(search_rows, len(df))):
        row_values = df_str.iloc[row_idx].tolist()
        found_cols = {}

        for std_col, keywords in required_cols.items():
            for col_idx, val in enumerate(row_values):
                val_clean = str(val).strip().replace('\n', '').replace('\r', '')
                if val_clean in keywords:
                    found_cols[std_col] = col_idx
                    break

            if len(found_cols) == len(required_cols):
                return row_idx, found_cols

    return None, None


# ==================== 结果合并功能 ====================

def merge_results(df1_processed, df2_processed, multi_id_names):
    """
    合并两个表格的结果并进行对比，返回最终的DataFrame
    
    若姓名属于 multi_id_names，则按 (姓名, 运营中心) 维度匹配；
    否则仍按姓名维度匹配。
    """
    if df1_processed is None or df2_processed is None:
        raise ValueError("请先处理两个表格的数据")

    result_rows = []
    multi_id_names = set(multi_id_names) if multi_id_names else set()

    # 构建右表映射
    right_map_name = {}          # 普通姓名 → 金额
    right_map_name_unit = {}     # (姓名, 运营中心) → 金额

    for _, right_row in df2_processed.iterrows():
        name = right_row['姓名']
        amount = right_row['预支数额']
        unit = str(right_row.get('运营中心', '')).strip()
        if name in multi_id_names:
            right_map_name_unit[(name, unit)] = amount
        else:
            right_map_name[name] = amount

    for _, left_row in df1_processed.iterrows():
        left_name = left_row['姓名']
        left_amount = left_row['预支金额']
        left_id = left_row['身份证号']
        left_unit = str(left_row.get('运营中心', '')).strip()

        if left_name in multi_id_names:
            key = (left_name, left_unit)
            if key in right_map_name_unit:
                right_amount = right_map_name_unit[key]
                if abs(left_amount - right_amount) < 0.01:
                    compare_result = "TRUE"
                    diff = 0.0
                else:
                    compare_result = "FALSE"
                    diff = left_amount - right_amount
                result_rows.append({
                    '姓名': left_name,
                    '预支金额': left_amount,
                    '身份证号': left_id,
                    '比较结果': compare_result,
                    '差额': round(diff, 2),
                    '姓名(右)': left_name,
                    '预支数额(右)': right_amount,
                    '备注': remark_by_diff(diff),
                    '运营中心': left_unit
                })
            else:
                diff = left_amount
                result_rows.append({
                    '姓名': left_name,
                    '预支金额': left_amount,
                    '身份证号': left_id,
                    '比较结果': 'FALSE',
                    '差额': round(diff, 2),
                    '姓名(右)': '',
                    '预支数额(右)': 0,
                    '备注': remark_by_diff(diff),
                    '运营中心': left_unit
                })
        else:
            if left_name in right_map_name:
                right_amount = right_map_name[left_name]
                if abs(left_amount - right_amount) < 0.01:
                    compare_result = "TRUE"
                    diff = 0.0
                else:
                    compare_result = "FALSE"
                    diff = left_amount - right_amount
                result_rows.append({
                    '姓名': left_name,
                    '预支金额': left_amount,
                    '身份证号': left_id,
                    '比较结果': compare_result,
                    '差额': round(diff, 2),
                    '姓名(右)': left_name,
                    '预支数额(右)': right_amount,
                    '备注': remark_by_diff(diff),
                    '运营中心': ''
                })
            else:
                diff = left_amount
                result_rows.append({
                    '姓名': left_name,
                    '预支金额': left_amount,
                    '身份证号': left_id,
                    '比较结果': 'FALSE',
                    '差额': round(diff, 2),
                    '姓名(右)': '',
                    '预支数额(右)': 0,
                    '备注': remark_by_diff(diff),
                    '运营中心': ''
                })

    left_names = set(df1_processed['姓名'])
    for _, right_row in df2_processed.iterrows():
        right_name = right_row['姓名']
        if right_name not in left_names:
            diff = -right_row['预支数额']
            result_rows.append({
                '姓名': right_name,
                '预支金额': 0,
                '身份证号': '',
                '比较结果': 'FALSE',
                '差额': round(diff, 2),
                '姓名(右)': right_name,
                '预支数额(右)': right_row['预支数额'],
                '备注': remark_by_diff(diff),
                '运营中心': str(right_row.get('运营中心', '')).strip()
            })

    final_df = pd.DataFrame(result_rows)

    if not final_df.empty:
        col_order = ['姓名', '预支金额', '身份证号', '比较结果', '差额', '姓名(右)', '预支数额(右)', '备注', '运营中心']
        final_df = final_df[col_order]
        final_df = final_df.sort_values(
            by=['姓名', '预支金额'], ascending=[True, False]
        ).reset_index(drop=True)

    return final_df
