# -- coding: utf-8 --
# @File : bijiao_core.py
# @Description: 工人预支申请表 vs 预支明细表 对比工具 - 核心功能模块

import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
import warnings

from .bijiao_frame import save_to_excel, format_number, remark_by_diff, parse_date_flexible, find_header_row, merge_results

warnings.filterwarnings('ignore')


class WorkerAdvanceComparatorCore:
    """工人预支对比工具 - 核心功能类"""

    def __init__(self):
        # 第一个表格（工人预支申请表）
        self.file1_path = None
        self.sheet1_name = None
        self.df_raw1 = None
        self.df1_processed = None
        self.header_row1 = None

        # 第二个表格（预支明细表）
        self.file2_path = None
        self.sheet2_name = None
        self.df_raw2 = None
        self.df2_processed = None
        self.header_row2 = None

        # 表1日期范围（工人预支申请表）
        self.start_date1 = None
        self.end_date1 = None

        # 表2日期范围（预支明细表）
        self.start_date2 = None
        self.end_date2 = None

        # 通用
        self.final_df = None

        # 预解析的中间数据（供分步处理使用）
        self._table1_parsed_data = None
        self._table2_parsed_data = None

        # 记录表1中"同姓名但多身份证号"的姓名集合（供表2条件拆分用）
        self._multi_id_names = set()

        # ========== 手动列映射配置（默认None=自动查找） ==========
        # 表1固定列（工人预支申请表）：F=预支数额, I=姓名, N=提交时间, S=身份证号
        self.table1_fixed_header_row = None  # 若设置，则跳过自动查找，直接使用该行作为表头
        self.table1_fixed_cols = {}          # {'姓名': 列索引, '预支金额': ..., ...}
        # 表2固定列（预支明细表）
        self.table2_fixed_header_row = None
        self.table2_fixed_cols = {}

    def set_file1(self, file_path, sheet_name):
        """设置第一个文件路径和工作表"""
        self.file1_path = file_path
        self.sheet1_name = sheet_name
        # 清空相关缓存，防止重选文件后使用旧数据
        self._table1_parsed_data = None
        self.df1_processed = None
        self.df_raw1 = None
        self.header_row1 = None
        self.start_date1 = None
        self.end_date1 = None
        self._multi_id_names = set()

    def set_file2(self, file_path, sheet_name):
        """设置第二个文件路径和工作表"""
        self.file2_path = file_path
        self.sheet2_name = sheet_name
        # 清空相关缓存，防止重选文件后使用旧数据
        self._table2_parsed_data = None
        self.df2_processed = None
        self.df_raw2 = None
        self.header_row2 = None
        self.start_date2 = None
        self.end_date2 = None
        self._multi_id_names = set()

    def set_table1_fixed_layout(self, header_row=0, name_col=8, amount_col=5, id_col=18, date_col=13, unit_col=3):
        """
        手动设置表1（工人预支申请表）的固定列位置
        :param header_row: 表头行索引（0-based），默认0=第1行
        :param name_col: 姓名列索引（0-based），默认8=I列
        :param amount_col: 预支金额列索引（0-based），默认5=F列
        :param id_col: 身份证号列索引（0-based），默认18=S列
        :param date_col: 提交时间列索引（0-based），默认13=N列
        :param unit_col: 运营中心列索引（0-based），默认3=D列
        """
        self.table1_fixed_header_row = header_row
        self.table1_fixed_cols = {
            '姓名': name_col,
            '预支金额': amount_col,
            '身份证号': id_col,
            '提交时间': date_col,
            '运营中心': unit_col
        }

    def set_table2_fixed_layout(self, header_row=0, name_col=4, amount_col=5, date_col=1, unit_col=2):
        """
        手动设置表2（预支明细表）的固定列位置
        默认: B列=预支日期(1), C列=运营中心(2), E列=姓名(4), F列=预支数额(5), 第2行表头(header_row=1)
        """
        self.table2_fixed_header_row = header_row
        self.table2_fixed_cols = {
            '姓名': name_col,
            '预支数额': amount_col,
            '预支日期': date_col,
            '运营中心': unit_col
        }

    def set_date_range1(self, start_date, end_date):
        """设置表1（工人预支申请表）的日期范围"""
        self.start_date1 = pd.to_datetime(start_date)
        self.end_date1 = pd.to_datetime(end_date)
        self.end_date1 = self.end_date1 + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    def set_date_range2(self, start_date, end_date):
        """设置表2（预支明细表）的日期范围"""
        self.start_date2 = pd.to_datetime(start_date)
        self.end_date2 = pd.to_datetime(end_date)
        self.end_date2 = self.end_date2 + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    def set_date_range_both(self, start_date, end_date):
        """统一设置两个表的日期范围（相同范围）"""
        self.set_date_range1(start_date, end_date)
        self.set_date_range2(start_date, end_date)

    def _parse_date_flexible(self, series):
        """更灵活地解析日期，兼容多种格式，统一转为日期（不含时间）"""
        return parse_date_flexible(series)

    def find_header_row(self, df, required_cols, search_rows=10):
        """
        在DataFrame的前几行中查找表头行
        返回：(header_row_index, column_mapping)
        column_mapping: {标准列名: 实际列索引}
        """
        return find_header_row(df, required_cols, search_rows)

    def _load_table1_core(self):
        """读取表1原始数据并解析列，返回解析后的DataFrame（不做日期筛选）"""
        if not self.file1_path or not self.sheet1_name:
            raise ValueError("请先设置第一个文件路径和工作表")

        self.df_raw1 = pd.read_excel(
            self.file1_path, sheet_name=self.sheet1_name, header=None
        )

        # ── 判断是否使用固定列布局 ──
        if self.table1_fixed_cols:
            header_row = self.table1_fixed_header_row
            col_mapping = self.table1_fixed_cols
            self.header_row1 = header_row
            print(f"[表1-固定布局] 使用手动指定的固定列位置：表头行={header_row}, 列映射={col_mapping}")
        else:
            # 原有自动查找逻辑
            required_cols = {
                '姓名': ['姓名'],
                '预支金额': ['预支金额'],
                '身份证号': ['身份证号'],
                '提交时间': ['提交时间']
            }
            header_row, col_mapping = self.find_header_row(self.df_raw1, required_cols)
            if header_row is None:
                raise ValueError(
                    "在第一个表格（工人预支申请表）中未找到完整的表头信息。\n"
                    "需要包含：姓名、预支金额、身份证号、提交时间"
                )
            self.header_row1 = header_row
            missing_cols = [k for k in required_cols.keys() if k not in col_mapping]
            if missing_cols:
                raise ValueError(f"第一个表格中缺少以下列: {', '.join(missing_cols)}")

        data_df = self.df_raw1.iloc[header_row + 1:].copy()
        data_df.columns = list(self.df_raw1.columns)
        data_df = data_df.reset_index(drop=True)

        name_col_idx = col_mapping['姓名']
        amount_col_idx = col_mapping['预支金额']
        id_col_idx = col_mapping['身份证号']
        date_col_idx = col_mapping['提交时间']

        result_df = pd.DataFrame()
        result_df['姓名'] = data_df.iloc[:, name_col_idx]

        # ── 处理金额列：支持文本金额转换 ──
        raw_amount = data_df.iloc[:, amount_col_idx]
        # 先尝试直接转数字
        result_df['预支金额'] = pd.to_numeric(raw_amount, errors='coerce')
        # 对转换失败的值（NaN），再尝试解析文本金额
        nan_mask = result_df['预支金额'].isna()
        if nan_mask.any():
            # 将无法转数字的文本值进行特殊处理
            text_amounts = raw_amount.astype(str).str.strip()
            # 常见中文金额文本 → 0
            zero_keywords = ['未预支', '无', '零', '没有', '', 'nan', 'None', 'NaN']
            for kw in zero_keywords:
                text_amounts = text_amounts.replace(kw, '0', regex=False)
            # 中文数字金额：去掉"元/块/毛/分/钱/整"等单位词
            text_amounts = text_amounts.str.replace(r'[元块毛分钱整]$', '', regex=True)
            text_amounts = text_amounts.str.replace(',', '').str.replace('，', '')
            # 尝试再次转换
            converted = pd.to_numeric(text_amounts, errors='coerce')
            # 只填充之前为NaN的位置
            result_df.loc[nan_mask, '预支金额'] = converted[nan_mask]
            # 如果还有转换失败的，记录诊断信息
            still_nan = result_df['预支金额'].isna()
            if still_nan.any():
                failed_samples = raw_amount[still_nan].dropna().unique()[:10]
                print(f"[表1-金额诊断] 以下金额值无法解析（前10个唯一值）: {failed_samples}")

        result_df['身份证号'] = data_df.iloc[:, id_col_idx].astype(str).str.strip()

        # 运营中心列（若配置了）
        unit_col_idx = col_mapping.get('运营中心')
        if unit_col_idx is not None:
            result_df['运营中心'] = data_df.iloc[:, unit_col_idx].astype(str).str.strip()

        # 增加原始值采样，方便诊断
        raw_dates = data_df.iloc[:, date_col_idx].dropna().head(5).tolist()
        print(f"[表1-日期诊断] 原始日期列前5个非空值: {raw_dates}")
        print(f"[表1-日期诊断] 原始值类型分布: {data_df.iloc[:, date_col_idx].apply(lambda x: type(x).__name__).value_counts().to_dict()}")

        result_df['提交时间'] = self._parse_date_flexible(data_df.iloc[:, date_col_idx])

        parsed_dates = result_df['提交时间'].dropna().head(5).tolist()
        print(f"[表1-日期诊断] 解析后前5个有效日期: {parsed_dates}")
        print(f"[表1-日期诊断] 解析成功率: {result_df['提交时间'].notna().sum()} / {len(result_df)}")

        # 姓名/金额列诊断
        print(f"[表1-姓名诊断] 原始姓名列前5个非空值: {data_df.iloc[:, name_col_idx].dropna().head(5).tolist()}")
        print(f"[表1-金额诊断] 原始金额列前5个非空值: {data_df.iloc[:, amount_col_idx].dropna().head(5).tolist()}")
        print(f"[表1-金额诊断] 转换后非空数: {result_df['预支金额'].notna().sum()}, 空值数: {result_df['预支金额'].isna().sum()}")
        print(f"[表1-金额诊断] 转换后金额前5个非空值: {result_df['预支金额'].dropna().head(5).tolist()}")
        print(f"[表1-姓名诊断] 姓名为空/空白/nan的行数: {(result_df['姓名'].astype(str).str.strip().isin(['', 'nan'])).sum()}")

        result_df = result_df.reset_index(drop=True)
        return result_df

    def get_date_range_table1(self):
        """预加载表1，返回可用日期范围 (min_date_str, max_date_str)"""
        result_df = self._load_table1_core()
        self._table1_parsed_data = result_df

        valid_dates = result_df['提交时间'].dropna()
        total_rows = len(result_df)
        valid_count = len(valid_dates)
        print(f"[表1-范围诊断] 总行数: {total_rows}, 有效日期数: {valid_count}")
        if valid_count > 0:
            print(f"[表1-范围诊断] 最小日期: {valid_dates.min()}, 最大日期: {valid_dates.max()}")

        if valid_dates.empty:
            return "无有效日期", "无有效日期"
        return valid_dates.min().strftime("%Y-%m-%d"), valid_dates.max().strftime("%Y-%m-%d")

    def process_first_table(self):
        """基于已存储的表1数据和日期范围，执行筛选、汇总

        汇总策略：
        1. 先按姓名全局判断身份证号唯一性
           A. 全局只有一个唯一非空身份证号  → 按 [姓名] 汇总（跨运营中心），填充该身份证
           B. 全局全部无身份证号            → 按 [姓名] 汇总，身份证=''
           C. 全局存在多个不同非空身份证号  → 进入步骤2（运营中心子判断）
        2. 在场景C下，对该姓名的每个运营中心子组再判断：
           a. 该运营中心内只有一个唯一非空身份证（其余空白）
              → 按 [姓名, 运营中心] 汇总，填充该身份证
           b. 该运营中心内全部空白身份证
              → 按 [姓名, 运营中心] 汇总，身份证=''
           c. 该运营中心内仍有多个不同非空身份证
              → 按 [姓名, 运营中心, 身份证号] 拆分汇总
        """
        if self._table1_parsed_data is None:
            self._table1_parsed_data = self._load_table1_core()

        result_df = self._table1_parsed_data.copy()
        total_rows = len(result_df)
        valid_dates = result_df['提交时间'].notna().sum()
        na_dates = total_rows - valid_dates

        result_df = result_df.dropna(subset=['姓名', '预支金额'])
        result_df = result_df[result_df['姓名'].astype(str).str.strip() != '']
        result_df = result_df[result_df['姓名'].astype(str).str.strip() != 'nan']
        result_df = result_df.reset_index(drop=True)

        mask = (result_df['提交时间'] >= self.start_date1) & (result_df['提交时间'] <= self.end_date1)
        result_df = result_df[mask].copy()

        if result_df.empty:
            msg = "在第一个表格中，所选时间范围内没有数据。\n"
            if total_rows > 0:
                msg += f"原因诊断：共读取 {total_rows} 行数据，其中 {valid_dates} 行成功识别日期，{na_dates} 行日期解析失败。\n"
                msg += f"筛选范围：{self.start_date1} ~ {self.end_date1}"
            raise ValueError(msg)

        result_df['身份证号'] = result_df['身份证号'].replace('nan', '').replace('NaN', '').replace('None', '').astype(str).str.strip()

        if '运营中心' not in result_df.columns:
            result_df['运营中心'] = ''

        multi_id_names = set()
        result_parts = []

        for name, group in result_df.groupby('姓名', sort=False):
            global_unique_ids = set(group['身份证号'].dropna().unique()) - {''}

            if len(global_unique_ids) == 0:
                # 场景B：全局全部空白身份证号
                df_sum = group.groupby(['姓名'], as_index=False)['预支金额'].sum()
                df_sum['身份证号'] = ''
                df_sum['运营中心'] = ''
                result_parts.append(df_sum)
            elif len(global_unique_ids) == 1:
                # 场景A：全局只有一个唯一非空身份证号
                the_id = list(global_unique_ids)[0]
                df_sum = group.groupby(['姓名'], as_index=False)['预支金额'].sum()
                df_sum['身份证号'] = the_id
                unit = group[group['身份证号'] == the_id]['运营中心'].iloc[0]
                df_sum['运营中心'] = unit
                result_parts.append(df_sum)
            else:
                # 场景C：全局存在多个不同非空身份证号 → 按运营中心子判断
                multi_id_names.add(name)
                for unit, unit_group in group.groupby('运营中心', sort=False):
                    unit_unique_ids = set(unit_group['身份证号'].dropna().unique()) - {''}

                    if len(unit_unique_ids) == 0:
                        # a. 该运营中心下全部空白
                        df_sum = unit_group.groupby(['姓名', '运营中心'], as_index=False)['预支金额'].sum()
                        df_sum['身份证号'] = ''
                        result_parts.append(df_sum)
                    elif len(unit_unique_ids) == 1:
                        # b. 该运营中心下只有一个唯一非空身份证（其余空白）
                        the_id = list(unit_unique_ids)[0]
                        df_sum = unit_group.groupby(['姓名', '运营中心'], as_index=False)['预支金额'].sum()
                        df_sum['身份证号'] = the_id
                        result_parts.append(df_sum)
                    else:
                        # c. 该运营中心下仍有多个不同非空身份证
                        df_sum = unit_group.groupby(['姓名', '运营中心', '身份证号'], as_index=False)['预支金额'].sum()
                        result_parts.append(df_sum)

        self._multi_id_names = multi_id_names
        self.df1_processed = pd.concat(result_parts, ignore_index=True)
        self.df1_processed = self.df1_processed[['姓名', '预支金额', '身份证号', '运营中心']]
        return self.df1_processed

    def _load_table2_core(self):
        """读取表2原始数据并解析列，返回解析后的DataFrame（不做日期筛选）"""
        if not self.file2_path or not self.sheet2_name:
            raise ValueError("请先设置第二个文件路径和工作表")

        self.df_raw2 = pd.read_excel(
            self.file2_path, sheet_name=self.sheet2_name, header=None
        )

        # 固定第2行为表头（0-based索引1），不再自动扫描前10行
        header_row = 1
        self.header_row2 = header_row

        if not self.table2_fixed_cols:
            raise ValueError("表2固定列布局未设置，请先调用 set_table2_fixed_layout()")
        col_mapping = self.table2_fixed_cols
        print(f"[表2-固定布局] 固定第2行为表头，列映射={col_mapping}")

        data_df = self.df_raw2.iloc[header_row + 1:].copy()
        data_df.columns = list(self.df_raw2.columns)
        data_df = data_df.reset_index(drop=True)

        date_col_idx = col_mapping['预支日期']
        name_col_idx = col_mapping['姓名']
        amount_col_idx = col_mapping['预支数额']

        result_df = pd.DataFrame()
        result_df['预支日期'] = self._parse_date_flexible(data_df.iloc[:, date_col_idx])
        result_df['姓名'] = data_df.iloc[:, name_col_idx]

        # 运营中心列（若配置了）
        unit_col_idx = col_mapping.get('运营中心')
        if unit_col_idx is not None:
            result_df['运营中心'] = data_df.iloc[:, unit_col_idx].astype(str).str.strip()

        # ── 处理金额列：支持文本金额转换 ──
        raw_amount = data_df.iloc[:, amount_col_idx]
        result_df['预支数额'] = pd.to_numeric(raw_amount, errors='coerce')
        nan_mask = result_df['预支数额'].isna()
        if nan_mask.any():
            text_amounts = raw_amount.astype(str).str.strip()
            zero_keywords = ['未预支', '无', '零', '没有', '', 'nan', 'None', 'NaN']
            for kw in zero_keywords:
                text_amounts = text_amounts.replace(kw, '0', regex=False)
            text_amounts = text_amounts.str.replace(r'[元块毛分钱整]$', '', regex=True)
            text_amounts = text_amounts.str.replace(',', '').str.replace('，', '')
            converted = pd.to_numeric(text_amounts, errors='coerce')
            result_df.loc[nan_mask, '预支数额'] = converted[nan_mask]
            still_nan = result_df['预支数额'].isna()
            if still_nan.any():
                failed_samples = raw_amount[still_nan].dropna().unique()[:10]
                print(f"[表2-金额诊断] 以下金额值无法解析（前10个唯一值）: {failed_samples}")

        # 增加原始值采样，方便诊断
        raw_dates = data_df.iloc[:, date_col_idx].dropna().head(5).tolist()
        print(f"[表2-日期诊断] 原始日期列前5个非空值: {raw_dates}")
        print(f"[表2-日期诊断] 原始值类型分布: {data_df.iloc[:, date_col_idx].apply(lambda x: type(x).__name__).value_counts().to_dict()}")

        parsed_dates = result_df['预支日期'].dropna().head(5).tolist()
        print(f"[表2-日期诊断] 解析后前5个有效日期: {parsed_dates}")
        print(f"[表2-日期诊断] 解析成功率: {result_df['预支日期'].notna().sum()} / {len(result_df)}")

        # 姓名/金额列诊断
        print(f"[表2-姓名诊断] 原始姓名列前5个非空值: {data_df.iloc[:, name_col_idx].dropna().head(5).tolist()}")
        print(f"[表2-金额诊断] 原始金额列前5个非空值: {data_df.iloc[:, amount_col_idx].dropna().head(5).tolist()}")
        print(f"[表2-金额诊断] 转换后非空数: {result_df['预支数额'].notna().sum()}, 空值数: {result_df['预支数额'].isna().sum()}")
        print(f"[表2-金额诊断] 转换后金额前5个非空值: {result_df['预支数额'].dropna().head(5).tolist()}")
        print(f"[表2-姓名诊断] 姓名为空/空白/nan的行数: {(result_df['姓名'].astype(str).str.strip().isin(['', 'nan'])).sum()}")

        result_df = result_df.reset_index(drop=True)
        return result_df

    def get_date_range_table2(self):
        """预加载表2，返回可用日期范围 (min_date_str, max_date_str)"""
        result_df = self._load_table2_core()
        self._table2_parsed_data = result_df

        valid_dates = result_df['预支日期'].dropna()
        total_rows = len(result_df)
        valid_count = len(valid_dates)
        print(f"[表2-范围诊断] 总行数: {total_rows}, 有效日期数: {valid_count}")
        if valid_count > 0:
            print(f"[表2-范围诊断] 最小日期: {valid_dates.min()}, 最大日期: {valid_dates.max()}")

        if valid_dates.empty:
            return "无有效日期", "无有效日期"
        return valid_dates.min().strftime("%Y-%m-%d"), valid_dates.max().strftime("%Y-%m-%d")

    def process_second_table(self):
        """基于已存储的表2数据和日期范围，执行筛选、汇总

        若表1中存在同姓名多身份证号的情况（multi_id_names），
        则表2中对应姓名需按运营中心拆分汇总；否则按姓名汇总。
        """
        if self._table2_parsed_data is None:
            self._table2_parsed_data = self._load_table2_core()

        result_df = self._table2_parsed_data.copy()
        total_rows = len(result_df)
        valid_dates = result_df['预支日期'].notna().sum()
        na_dates = total_rows - valid_dates

        result_df = result_df.dropna(subset=['姓名', '预支数额'])
        result_df = result_df[result_df['姓名'].astype(str).str.strip() != '']
        result_df = result_df[result_df['姓名'].astype(str).str.strip() != 'nan']
        result_df = result_df.reset_index(drop=True)

        mask = (result_df['预支日期'] >= self.start_date2) & (result_df['预支日期'] <= self.end_date2)
        result_df = result_df[mask].copy()

        if result_df.empty:
            msg = "在第二个表格中，所选时间范围内没有数据。\n"
            if total_rows > 0:
                msg += f"原因诊断：共读取 {total_rows} 行数据，其中 {valid_dates} 行成功识别日期，{na_dates} 行日期解析失败。\n"
                msg += f"筛选范围：{self.start_date2} ~ {self.end_date2}"
            raise ValueError(msg)

        # 运营中心列兜底
        if '运营中心' not in result_df.columns:
            result_df['运营中心'] = ''

        if self._multi_id_names:
            mask_multi = result_df['姓名'].isin(self._multi_id_names)
            # multi_id 姓名：按 [姓名, 运营中心] 分组
            df_multi = result_df[mask_multi].groupby(['姓名', '运营中心'], as_index=False)['预支数额'].sum()
            # 普通姓名：按 [姓名] 分组，运营中心留空
            df_normal = result_df[~mask_multi].groupby(['姓名'], as_index=False)['预支数额'].sum()
            df_normal['运营中心'] = ''
            self.df2_processed = pd.concat([df_multi, df_normal], ignore_index=True)
        else:
            self.df2_processed = result_df.groupby(['姓名'], as_index=False)['预支数额'].sum()
            self.df2_processed['运营中心'] = ''

        return self.df2_processed

    def merge_results(self):
        """合并两个表格的结果并进行对比，返回最终的DataFrame

        若姓名属于 multi_id_names，则按 (姓名, 运营中心) 维度匹配；
        否则仍按姓名维度匹配。
        """
        self.final_df = merge_results(self.df1_processed, self.df2_processed, self._multi_id_names)
        return self.final_df
