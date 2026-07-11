# -*- coding: utf-8 -*-
"""
Excel汇总处理工具 - 核心处理模块
根据备注2列内容缩写，汇总打款金额，生成供应商汇总表

包含：SummaryProcessor 类
工具函数请参考：hs_utils.py
数据核对请参考：hs_checker.py
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
import re
import traceback
from typing import List, Tuple, Optional
import warnings
warnings.filterwarnings('ignore')

from finance.cxm.hs_utils import match_month_in_text


class SummaryProcessor:
    """
    汇总处理核心类
    功能包括：
    - 读取Excel文件和工作表
    - 检测"备注2"列并应用缩写规则
    - 按供应商和公司汇总打款金额
    - 生成汇总报表
    """

    # 空白字符正则表达式（包括零宽空格、不间断空格等）
    BLANK_PATTERN = r'[\s\u200b\u200c\u200d\u00a0\u3000]+'

    def __init__(self):
        """初始化处理器"""
        self.current_file_path: Optional[str] = None
        self.sheet_names: List[str] = []
        self.current_sheet: Optional[str] = None
        self.data_frame: Optional[pd.DataFrame] = None
        self.output_data: Optional[pd.DataFrame] = None

    def load_file(self, file_path: str) -> Tuple[bool, str, List[str]]:
        """
        加载Excel文件

        Args:
            file_path: Excel文件路径

        Returns:
            (success, message, sheet_names): 是否成功、消息、工作表列表
        """
        try:
            self.current_file_path = file_path
            wb = load_workbook(file_path, data_only=True)
            self.sheet_names = wb.sheetnames
            wb.close()  # 及时释放资源，后续用pd.read_excel读取数据
            return True, f"成功读取文件：{os.path.basename(file_path)}", self.sheet_names
        except Exception as e:
            return False, f"读取文件失败：{str(e)}", []

    def load_sheet(self, sheet_name: str) -> Tuple[bool, str]:
        """
        加载指定工作表

        Args:
            sheet_name: 工作表名称

        Returns:
            (success, message): 是否成功、消息
        """
        try:
            self.current_sheet = sheet_name
            self.data_frame = pd.read_excel(self.current_file_path, sheet_name=sheet_name)
            return True, f"成功加载工作表：{sheet_name}"
        except Exception as e:
            return False, f"加载工作表失败：{str(e)}"

    def _clean_blank_chars(self, text: str) -> str:
        """移除所有空白字符（包括不可见字符）"""
        return re.sub(self.BLANK_PATTERN, '', str(text))

    def _build_unique_key(self, row, has_payee_name: bool):
        """构建供应商唯一键"""
        supplier = row['供应商名称']
        if pd.isna(supplier) or supplier == '':
            return None
        supplier_str = str(supplier).strip()

        if has_payee_name and pd.notna(row['收款姓名']) and str(row['收款姓名']).strip() != '':
            payee_name = str(row['收款姓名']).strip()
            return (supplier_str, payee_name)
        return supplier_str

    def _match_supplier(self, row_supplier, target_supplier: str) -> bool:
        """匹配供应商名称（支持空白字符容错）"""
        if pd.isna(row_supplier):
            return False
        row_clean = str(row_supplier).strip()
        if row_clean == target_supplier:
            return True
        return self._clean_blank_chars(row_clean) == target_supplier

    def _match_payee(self, row_payee, target_payee: str, has_payee_name: bool) -> bool:
        """匹配收款姓名（支持空白字符容错）"""
        if not has_payee_name:
            return True
        if pd.notna(row_payee):
            row_clean = str(row_payee).strip()
            if row_clean == target_payee:
                return True
            return self._clean_blank_chars(row_clean) == target_payee
        return target_payee == ''

    def _abbreviate_company_name(self, company_name: str) -> Optional[str]:
        """
        根据缩写规则处理公司名称

        Args:
            company_name: 原始公司名称

        Returns:
            缩写后的公司名称
        """
        # 空白或空值返回"（空白）"
        if pd.isna(company_name) or company_name == '':
            return "（空白）"

        # 使用正则表达式移除所有空白字符（包括不可见字符）
        company_str = str(company_name)
        cleaned_str = re.sub(self.BLANK_PATTERN, '', company_str)

        if cleaned_str == '':
            return "（空白）"

        company_str = company_str.strip()

        # 规则1: "刘先锋现金卡支付" → "刘先锋现金卡"
        if company_str == "刘先锋现金卡支付":
            return "刘先锋现金卡"

        # 规则2: "XXX对公打款"或"XXX代发打款" → "XXX"
        for suffix in ["对公打款", "代发打款"]:
            if company_str.endswith(suffix):
                return company_str.replace(suffix, "")

        # 不符合规则，返回原值
        return company_str

    def _extract_payment_type(self, remark2: str) -> Optional[str]:
        """从备注2值中提取打款类型"""
        if pd.isna(remark2):
            return None
        val_str = str(remark2).strip()

        if "刘先锋现金卡" in val_str:
            return "刘先锋现金卡"
        if val_str.endswith("对公打款"):
            return "对公打款"
        if val_str.endswith("代发打款"):
            return "代发打款"
        return None

    def detect_payment_types(self) -> List[str]:
        """检测备注2列中存在的打款类型"""
        if self.data_frame is None:
            return []
        if "备注2" not in self.data_frame.columns:
            return []

        types_found = set()
        for val in self.data_frame["备注2"].dropna():
            pt = self._extract_payment_type(val)
            if pt:
                types_found.add(pt)

        # 按固定顺序排序
        order = {"刘先锋现金卡": 0, "对公打款": 1, "代发打款": 2}
        return sorted(list(types_found), key=lambda x: order.get(x, 99))

    def detect_months(self) -> List[str]:
        """检测备注1列中存在的月份（支持X月和X-Y月格式，如1月、1-2月）"""
        if self.data_frame is None:
            return []
        if "备注1" not in self.data_frame.columns:
            return []

        months_found = set()
        for val in self.data_frame["备注1"].dropna():
            val_str = str(val).strip()
            # 精确提取单月份，避免"12月"被拆成"1月"和"2月"
            found = re.findall(r'(?<!\d)(\d{1,2}月)(?!\d)', val_str)
            for m in found:
                try:
                    num = int(m.replace('月', ''))
                    if 1 <= num <= 12:
                        months_found.add(m)
                except ValueError:
                    pass

            # 提取范围月份，如"1-2月"、"3-5月"
            range_found = re.findall(r'(?<!\d)(\d{1,2}-\d{1,2}月)(?!\d)', val_str)
            for m in range_found:
                try:
                    start, end = m.replace('月', '').split('-')
                    start_num = int(start)
                    end_num = int(end)
                    if 1 <= start_num <= 12 and 1 <= end_num <= 12:
                        months_found.add(m)
                except ValueError:
                    pass

        # 排序：单月份在前，范围月份在后，均按起始数字排序
        def sort_key(x):
            if '-' in x:
                start = int(x.split('-')[0])
                return (1, start)
            return (0, int(x.replace('月', '')))

        return sorted(list(months_found), key=sort_key)

    def detect_amount_columns(self) -> List[Tuple[int, str]]:
        """检测所有名为'打款金额'的列，返回列索引和列字母标识"""
        if self.data_frame is None:
            return []

        amount_cols = [(i, get_column_letter(i + 1)) for i, col in enumerate(self.data_frame.columns) if '打款金额' in str(col)]
        return amount_cols

    def process_data(self, selected_types: Optional[List[str]] = None,
                     selected_months: Optional[List[str]] = None,
                     selected_amount_col_idx: Optional[int] = None) -> Tuple[bool, str]:
        """
        处理数据：缩写、汇总、生成报表

        Returns:
            (success, message): 是否成功、消息
        """
        try:
            if self.data_frame is None:
                return False, "请先选择工作表"

            df = self.data_frame

            # 检测"备注2"列
            if "备注2" not in df.columns:
                return False, "未找到'备注2'列"

            # 使用用户选择的打款金额列
            if selected_amount_col_idx is not None:
                amount_col_idx = selected_amount_col_idx
            else:
                return False, "未选择或检测到'打款金额'列"

            # 应用缩写规则
            df['公司缩写'] = df['备注2'].apply(self._abbreviate_company_name)

            # 根据用户选择的打款类型筛选
            if selected_types:
                def match_payment_type(remark2):
                    if pd.isna(remark2):
                        return False
                    val_str = str(remark2).strip()
                    for pt in selected_types:
                        if pt == "刘先锋现金卡" and "刘先锋现金卡" in val_str:
                            return True
                        if pt == "对公打款" and val_str.endswith("对公打款"):
                            return True
                        if pt == "代发打款" and val_str.endswith("代发打款"):
                            return True
                    return False

                df = df[df['备注2'].apply(match_payment_type)]

                if df.empty:
                    return False, "没有符合所选打款类型的数据"

            # 根据用户选择的月份筛选（精确字符串匹配，不展开范围）
            if selected_months:
                if "备注1" not in df.columns:
                    return False, "未找到'备注1'列，无法进行月份筛选"

                df = df[df['备注1'].apply(
                    lambda val: any(match_month_in_text(m, val) for m in selected_months)
                )]

                if df.empty:
                    return False, f"没有符合所选月份的数据"

            # 获取所有唯一的供应商名称
            if "供应商名称" not in df.columns:
                return False, "未找到'供应商名称'列"

            # 获取所有唯一的缩写公司名称
            unique_companies = df['公司缩写'].unique().tolist()
            unique_companies = [c for c in unique_companies if c is not None and c != '']

            if not unique_companies:
                return False, "没有有效的公司缩写数据"

            # 确保"（空白）"列在最后（在"总计"列之前）
            if "（空白）" in unique_companies:
                unique_companies.remove("（空白）")
                unique_companies.append("（空白）")

            # 检查是否有"收款姓名"列
            has_payee_name = "收款姓名" in df.columns

            # 第一步：按供应商+收款姓名分组收集空白行和非空白行
            blank_groups = {}  # {key: [rows]}
            non_blank_groups = {}  # {key: {company: total_amount}}

            for _, row in df.iterrows():
                unique_key = self._build_unique_key(row, has_payee_name)
                if unique_key is None:
                    continue

                remark2 = row['备注2']
                is_blank = pd.isna(remark2) or self._clean_blank_chars(remark2) == ''

                if is_blank:
                    # 收集空白行
                    if unique_key not in blank_groups:
                        blank_groups[unique_key] = []
                    blank_groups[unique_key].append(row)
                else:
                    # 汇总非空白行的金额
                    company_abbrev = row['公司缩写']
                    if unique_key not in non_blank_groups:
                        non_blank_groups[unique_key] = {}

                    try:
                        if len(row) > amount_col_idx:
                            amount = row.iloc[amount_col_idx]
                            if pd.notna(amount):
                                amount_value = float(amount)
                                if company_abbrev in non_blank_groups[unique_key]:
                                    non_blank_groups[unique_key][company_abbrev] += amount_value
                                else:
                                    non_blank_groups[unique_key][company_abbrev] = amount_value
                    except (ValueError, TypeError, IndexError):
                        pass

            # 第二步：初始化汇总数据字典
            all_keys = list(set(list(blank_groups.keys()) + list(non_blank_groups.keys())))
            summary_data = {}
            for key in all_keys:
                if isinstance(key, tuple):
                    supplier_str, payee_name = key
                    summary_data[key] = {
                        'supplier': supplier_str,
                        'payee': payee_name,
                        'amounts': {company: None for company in unique_companies}
                    }
                else:
                    summary_data[key] = {
                        'supplier': key,
                        'payee': '',
                        'amounts': {company: None for company in unique_companies}
                    }

            # 第三步：填充非空白行的汇总数据
            for key, company_amounts in non_blank_groups.items():
                if key not in summary_data:
                    continue
                for company, amount in company_amounts.items():
                    if company != "（空白）" and company in summary_data[key]['amounts']:
                        summary_data[key]['amounts'][company] = amount

            # 第四步：处理空白行，汇总到"（空白）"列
            for key, rows in blank_groups.items():
                if key not in summary_data:
                    continue

                supplier_str = summary_data[key]['supplier']
                payee_name = summary_data[key]['payee']
                blank_total = 0.0

                for row in rows:
                    # 匹配供应商
                    if not self._match_supplier(row['供应商名称'], supplier_str):
                        continue

                    # 匹配收款姓名
                    if not self._match_payee(row.get('收款姓名'), payee_name, has_payee_name):
                        continue

                    # 累加金额
                    try:
                        if len(row) > amount_col_idx:
                            amount = row.iloc[amount_col_idx]
                            if pd.notna(amount):
                                amount_value = float(amount)
                                blank_total += amount_value
                    except (ValueError, TypeError, IndexError):
                        pass

                # 设置空白列总额
                if blank_total > 0 and "（空白）" in summary_data[key]['amounts']:
                    summary_data[key]['amounts']["（空白）"] = blank_total

            # 创建输出DataFrame
            output_columns = ['供应商名称'] + unique_companies
            output_rows = []

            for key in summary_data:
                row_data = {'供应商名称': summary_data[key]['supplier']}
                for company in unique_companies:
                    row_data[company] = summary_data[key]['amounts'][company]
                # 保存收款姓名，用于后续处理"公司"列
                row_data['_payee_name'] = summary_data[key]['payee']
                output_rows.append(row_data)

            self.output_data = pd.DataFrame(output_rows, columns=output_columns + ['_payee_name'])

            # 添加"总计"列
            company_cols = [col for col in output_columns if col != '供应商名称']
            self.output_data['总计'] = self.output_data[company_cols].sum(axis=1)

            # 添加"公司"列数据
            self._add_company_column()

            # 重新排列列顺序
            # "总计"列和"公司"列之间为"备注"列
            final_columns = ['供应商名称'] + unique_companies + ['总计', '备注', '公司']
            self.output_data = self.output_data.reindex(columns=final_columns)
            # 确保"备注"列值为空白
            self.output_data['备注'] = ''

            return True, f"处理完成！共{len(self.output_data)}个供应商，{len(unique_companies)}个公司"

        except Exception as e:
            error_msg = f"数据汇总失败：{str(e)}\n{traceback.format_exc()}"
            return False, error_msg

    def _add_company_column(self) -> None:
        """
        添加"公司"列数据

        从"收款姓名"列提取符合条件的值：
        - 包含"刘先锋现金卡支付"
        - 包含"个体工商户"
        - 包含"有限公司"
        - 包含"分公司"

        注意：现在每个供应商+收款姓名的组合是一行，所以"公司"列显示对应的收款姓名
        """
        if self.output_data is None or '_payee_name' not in self.output_data.columns:
            return

        # 向量化处理公司列
        keywords = ["刘先锋现金卡支付", "个体工商户", "个体户", "有限公司", "分公司", "服务部"]

        def extract_company(payee_name):
            if payee_name and any(keyword in payee_name for keyword in keywords):
                return payee_name
            return ''

        self.output_data['公司'] = self.output_data['_payee_name'].apply(extract_company)

        # 删除临时列
        if '_payee_name' in self.output_data.columns:
            del self.output_data['_payee_name']

    def _extract_date_from_sheet_name(self, sheet_name: str, include_timestamp: bool = False) -> str:
        """
        从工作表名称中提取日期部分

        Args:
            sheet_name: 工作表名称
            include_timestamp: 是否在无法提取日期时附加时间戳

        Returns:
            提取的日期字符串或清理后的名称
        """
        if not sheet_name:
            if include_timestamp:
                return f"供应商汇总_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}"
            return "供应商汇总"

        pattern = r'(\d{4}年\d{1,2}月)'
        match = re.search(pattern, sheet_name)

        if match:
            return match.group(1)
        else:
            clean_name = re.sub(r'[<>:"/\\|?*]', '', sheet_name)
            if clean_name:
                return clean_name
            if include_timestamp:
                return f"供应商汇总_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}"
            return "供应商汇总"



