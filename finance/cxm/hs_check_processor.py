# -*- coding: utf-8 -*-
"""
核对功能处理器 - CheckProcessor 类
处理两个独立子功能：
  1. 发放记录表汇总（去后缀合并供应商 + 透视表）
  2. 供应商对账单汇总（含负数金额求和 + 透视表）
"""

import re
import pandas as pd
import os
from typing import List, Dict, Tuple, Optional


class CheckProcessor:
    """核对功能的独立处理器"""

    def __init__(self):
        self.df_disbursement = None   # 发放记录表 DataFrame
        self.df_statement = None       # 对账单 DataFrame
        self.output_disbursement = None  # 发放记录汇总结果
        self.output_statement = None     # 对账单汇总结果

    # ============================================================
    #  子功能1：发放记录表汇总
    # ============================================================

    def load_disbursement(self, file_path: str, sheet_name: str) -> Tuple[bool, str]:
        """加载发放记录表"""
        try:
            self.df_disbursement = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
            return True, f"加载成功，共{len(self.df_disbursement)}行，" \
                   f"列：{', '.join(self.df_disbursement.columns.tolist())}"
        except Exception as e:
            return False, f"加载失败：{e}"

    def detect_disbursement_headers(self) -> dict:
        """
        检测发放记录表的表头（第一行）
        返回：{'supplier_col', 'amount_col_c', 'amount_col_h',
               'remark1_col', 'remark2_col', 'all_columns'}
        """
        if self.df_disbursement is None:
            return {}
        cols = self.df_disbursement.columns.tolist()
        result = {'all_columns': cols}

        # 供应商名称（精确匹配）
        for c in cols:
            if c.strip() == '供应商名称':
                result['supplier_col'] = c
                break

        # 备注1、备注2（精确匹配）
        for c in cols:
            if c.strip() == '备注1':
                result['remark1_col'] = c
            if c.strip() == '备注2':
                result['remark2_col'] = c

        # C列、H列（按位置：第3列=C，第8列=H，0-based index）
        if len(cols) >= 3:
            result['amount_col_c'] = cols[2]   # C = index 2
        if len(cols) >= 8:
            result['amount_col_h'] = cols[7]   # H = index 7

        return result

    @staticmethod
    def normalize_supplier(name) -> str:
        """
        去除供应商名称末尾的数字后缀
        '供应商-德斯邦人力'   → '供应商-德斯邦人力'
        '供应商-德斯邦人力1'  → '供应商-德斯邦人力'
        '供应商-德斯邦人力12' → '供应商-德斯邦人力'
        """
        if pd.isna(name):
            return ''
        s = str(name).strip()
        return re.sub(r'\d+$', '', s)

    @staticmethod
    def extract_month_from_remark1(text) -> Optional[str]:
        """从备注1提取月份：'4月瑞立服务费-端木财' → '4月'"""
        if pd.isna(text):
            return None
        s = str(text).strip()
        m = re.search(r'(\d{1,2}月)', s)
        return m.group(1) if m else None

    @staticmethod
    def extract_company_and_type_from_remark2(text) -> Tuple[Optional[str], Optional[str]]:
        """
        从备注2提取公司名和打款类型
        '温州锦途外包代发打款' → ('温州锦途', '代发打款')
        '芜湖才库对公打款'   → ('芜湖才库', '对公打款')
        '张三刘先锋现金卡支付' → ('张三', '刘先锋现金卡支付')
        """
        if pd.isna(text):
            return None, None
        s = str(text).strip()
        keywords = ['代发打款', '对公打款', '刘先锋现金卡支付']
        for kw in keywords:
            if kw in s:
                idx = s.index(kw)
                company = s[:idx].strip()
                # 去除公司名末尾的"外包"（如"温州锦途外包"→"温州锦途"）
                company = re.sub(r'外包$', '', company)
                # "合静"补全为"温州合静"（如"合静代发打款"→公司="温州合静"）
                if company == '合静':
                    company = '温州合静'
                return company, kw
        return None, None

    def process_disbursement(self,
                              amount_col: str,
                              selected_months: Optional[List[str]] = None,
                              selected_types: Optional[List[str]] = None) -> Tuple[bool, str]:
        """
        处理发放记录表，生成透视汇总表
        amount_col: 用户选择的打款金额列名（C列或H列）
        selected_months: 选择的月份列表（来自备注1），None=不过滤
        selected_types: 选择的打款类型列表（来自备注2），None=不过滤
        """
        if self.df_disbursement is None:
            return False, '请先加载发放记录表'

        df = self.df_disbursement.copy()
        headers = self.detect_disbursement_headers()

        # 1. 过滤月份（备注1）
        if selected_months and 'remark1_col' in headers:
            col = headers['remark1_col']
            mask = df[col].apply(
                lambda x: self.extract_month_from_remark1(x) in selected_months
                if pd.notna(x) else False
            )
            df = df[mask].copy()
            if df.empty:
                return False, '按月份过滤后无数据'

        # 2. 过滤打款类型（备注2）
        if selected_types and 'remark2_col' in headers:
            col = headers['remark2_col']
            mask = df[col].apply(
                lambda x: any(kw in str(x) for kw in selected_types) if pd.notna(x) else False
            )
            df = df[mask].copy()
            if df.empty:
                return False, '按打款类型过滤后无数据'

        if 'supplier_col' not in headers:
            return False, "未找到'供应商名称'列"
        supplier_col = headers['supplier_col']

        # 3. 提取月份（备注1）、公司+类型（备注2）
        if 'remark1_col' in headers:
            df['_month'] = df[headers['remark1_col']].apply(self.extract_month_from_remark1)
        else:
            df['_month'] = None

        if 'remark2_col' in headers:
            df['_company'], df['_type'] = zip(*df[headers['remark2_col']].apply(
                lambda x: self.extract_company_and_type_from_remark2(x)
            ))
        else:
            df['_company'] = None
            df['_type'] = None

        # 4. 归一化供应商名称（去数字后缀）
        df['_norm_supplier'] = df[supplier_col].apply(self.normalize_supplier)

        # 5. 构建透视表：行=归一化供应商，列=公司，值=amount_col 求和
        #    如果 _company 有 None，填充为'未知'
        df['_company'] = df['_company'].fillna('未知')

        # 只保留金额为数字的行
        df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
        df = df[df[amount_col].notna()].copy()

        if df.empty:
            return False, '打款金额列无有效数值'

        # 透视表
        pivot = df.pivot_table(
            index='_norm_supplier',
            columns='_company',
            values=amount_col,
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot.rename(columns={'_norm_supplier': '供应商名称'}, inplace=True)

        # 添加总计列
        company_cols = [c for c in pivot.columns if c != '供应商名称']
        pivot['总计'] = pivot[company_cols].sum(axis=1)

        # 列顺序：供应商名称 + 各公司列 + 总计
        final_cols = ['供应商名称'] + company_cols + ['总计']
        pivot = pivot[[c for c in final_cols if c in pivot.columns]]

        self.output_disbursement = pivot
        return True, f"发放记录汇总完成，共{len(pivot)}行，" \
                          f"公司数：{len(company_cols)}"

    # ============================================================
    #  子功能2：供应商对账单汇总
    # ============================================================

    def load_statement(self, file_path: str, sheet_name: str) -> Tuple[bool, str]:
        """加载供应商对账单"""
        try:
            self.df_statement = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
            return True, f"加载成功，共{len(self.df_statement)}行，" \
                   f"列：{', '.join(self.df_statement.columns.tolist())}"
        except Exception as e:
            return False, f"加载失败：{e}"

    def detect_statement_headers(self) -> dict:
        """
        检测对账单的表头
        返回：{'supplier_col', 'year_month_col', 'amount_col', 'company_col', 'all_columns'}
        """
        if self.df_statement is None:
            return {}
        cols = self.df_statement.columns.tolist()
        result = {'all_columns': cols}

        for c in cols:
            cs = c.strip()
            if cs == '供应商名称':
                result['supplier_col'] = c
            elif cs == '核算年月':
                result['year_month_col'] = c
            elif cs == '金额':
                result['amount_col'] = c
            elif cs == '公司':
                result['company_col'] = c

        return result

    def get_unique_year_months(self) -> List[str]:
        """获取对账单中核算年月的唯一值列表（排序）"""
        if self.df_statement is None:
            return []
        headers = self.detect_statement_headers()
        if 'year_month_col' not in headers:
            return []
        col = headers['year_month_col']
        vals = self.df_statement[col].dropna().astype(str).str.strip().unique().tolist()
        vals.sort()
        return vals

    def process_statement(self, selected_year_months: List[str]) -> Tuple[bool, str]:
        """
        处理对账单，按选择的核算年月过滤，生成透视汇总表
        selected_year_months: 选择的年月列表，如 ['2026-04', '2026-05']
        """
        if self.df_statement is None:
            return False, '请先加载对账单'

        headers = self.detect_statement_headers()
        for key in ['supplier_col', 'year_month_col', 'amount_col', 'company_col']:
            if key not in headers:
                return False, f"未找到必要列，已检测到列：{', '.join(headers.get('all_columns', []))}"

        df = self.df_statement.copy()
        col_ym = headers['year_month_col']
        col_amt = headers['amount_col']
        col_sup = headers['supplier_col']
        col_cmp = headers['company_col']

        # 过滤核算年月
        df[col_ym] = df[col_ym].apply(lambda x: str(x).strip() if pd.notna(x) else '')
        df = df[df[col_ym].isin(selected_year_months)].copy()
        if df.empty:
            return False, '按核算年月过滤后无数据'

        # 金额转数字（保留负数！）
        df[col_amt] = pd.to_numeric(df[col_amt], errors='coerce')
        df = df[df[col_amt].notna()].copy()

        if df.empty:
            return False, '金额列无有效数值'

        # 透视表：行=供应商名称，列=公司，值=金额 sum（含负数自动正确处理）
        pivot = df.pivot_table(
            index=col_sup,
            columns=col_cmp,
            values=col_amt,
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot.rename(columns={col_sup: '供应商名称'}, inplace=True)

        # 添加总计列（横向求和，保留符号）
        company_cols = [c for c in pivot.columns if c != '供应商名称']
        pivot['总计'] = pivot[company_cols].sum(axis=1)

        # 列顺序：供应商名称 + 各公司列 + 总计
        final_cols = ['供应商名称'] + company_cols + ['总计']
        pivot = pivot[[c for c in final_cols if c in pivot.columns]]

        self.output_statement = pivot
        return True, f"对账单汇总完成，共{len(pivot)}行，" \
                          f"公司数：{len(company_cols)}"

    # ============================================================
    #  核对对比：发放记录汇总 vs 对账单汇总
    # ============================================================

    def reconcile(self) -> Tuple[bool, str, pd.DataFrame]:
        """
        对比发放记录汇总和对账单汇总，生成核对结果表。
        以 (供应商名称, 公司) 为最小粒度逐格比对。

        返回: (成功与否, 消息, 核对结果DataFrame)
        DataFrame 列: [供应商名称, 公司, 发放记录金额, 对账单金额, 差额, 状态]

        状态判定：
          已核对   — 两边都有值且 A == B（严格相等）
          待核对   — 两边都有值但 A != B
          未匹配   — 只有一边有值（A有B无 或 A无B有）
        """
        if self.output_disbursement is None:
            return False, '请先完成发放记录汇总', pd.DataFrame()
        if self.output_statement is None:
            return False, '请先完成对账单汇总', pd.DataFrame()

        df_a = self.output_disbursement  # 发放记录汇总
        df_b = self.output_statement     # 对账单汇总

        # 归一化供应商名称用于匹配（去除"供应商-"前缀，使"供应商-兴程"与"兴程"能匹配）
        def _norm(supplier):
            return re.sub(r'^供应商-', '', str(supplier).strip())

        # 提取公司列名（排除 '供应商名称' 和 '总计'）
        companies_a = [c for c in df_a.columns if c not in ('供应商名称', '总计')]
        companies_b = [c for c in df_b.columns if c not in ('供应商名称', '总计')]
        all_companies = sorted(set(companies_a) | set(companies_b))

        # 收集所有 (供应商, 公司) 组合及其值
        rows = []
        for _, row_a in df_a.iterrows():
            supplier = _norm(row_a.get('供应商名称', ''))
            for company in all_companies:
                val_a = float(row_a.get(company, 0)) if company in df_a.columns else 0.0
                rows.append((supplier, company, val_a, 'A'))

        for _, row_b in df_b.iterrows():
            supplier = _norm(row_b.get('供应商名称', ''))
            for company in all_companies:
                val_b = float(row_b.get(company, 0)) if company in df_b.columns else 0.0
                rows.append((supplier, company, val_b, 'B'))

        # 按 (供应商, 公司) 聚合
        record = {}  # (supplier, company) -> {'a': val, 'b': val}
        for supplier, company, val, source in rows:
            key = (str(supplier).strip(), str(company).strip())
            if key not in record:
                record[key] = {'a': 0.0, 'b': 0.0}
            if source == 'A':
                record[key]['a'] = val
            else:
                record[key]['b'] = val

        # 构建结果行
        result_rows = []
        stats = {'matched': 0, 'pending': 0, 'unmatched_a': 0, 'unmatched_b': 0,
                 'pending_diff': 0.0, 'unmatched_a_amt': 0.0, 'unmatched_b_amt': 0.0}

        for (supplier, company), vals in sorted(record.items()):
            a_val = vals['a']
            b_val = vals['b']
            diff = a_val - b_val

            # 状态判定
            if a_val == 0 and b_val == 0:
                continue  # 双方空白不显示

            if abs(a_val) > 0 and abs(b_val) > 0:
                if a_val == b_val:
                    status = '已核对'
                    stats['matched'] += 1
                else:
                    status = '待核对'
                    stats['pending'] += 1
                    stats['pending_diff'] += abs(diff)
            elif abs(a_val) > 0:
                status = '未匹配'
                stats['unmatched_a'] += 1
                stats['unmatched_a_amt'] += a_val
            else:
                status = '未匹配'
                stats['unmatched_b'] += 1
                stats['unmatched_b_amt'] += abs(b_val)

            result_rows.append({
                '供应商名称': supplier,
                '公司': company,
                '发放记录金额': a_val if a_val != 0 else '',
                '对账单金额': b_val if b_val != 0 else '',
                '差额': diff if (abs(a_val) > 0 and abs(b_val) > 0) else (a_val if abs(a_val) > 0 else -b_val),
                '状态': status,
            })

        result_df = pd.DataFrame(result_rows)
        self.reconcile_stats = stats
        self.reconcile_result = result_df

        total = len(result_df)
        msg = (
            f"核对完成！共 {total} 条 | "
            f"已核对 {stats['matched']} 条 | "
            f"待核对 {stats['pending']} 条 | "
            f"未匹配(发放记录表) {stats['unmatched_a']} 条 | "
            f"未匹配(对账单表) {stats['unmatched_b']} 条"
        )
        return True, msg, result_df
