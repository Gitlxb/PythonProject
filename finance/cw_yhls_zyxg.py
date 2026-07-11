# -- coding: utf-8 --
# @Time : 2025-05-30 15:58
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : cw_yhls_zyxg.py
# @Software: PyCharm

import pandas as pd
from tkinter import Tk, filedialog
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkinter import ttk
import re


def select_file():
    """让用户选择Excel文件并返回文件路径"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
    )
    return file_path


def process_excel(file_path):
    """处理Excel文件"""
    # 读取Excel文件的所有工作表
    xls = pd.ExcelFile(file_path)

    # 检查是否存在"各种内容"工作表
    if "各种内容" not in xls.sheet_names:
        raise ValueError("Excel文件中缺少工作表: 各种内容")

    # 读取主数据工作表（假设第一个工作表是主数据）
    sheet_names = [name for name in xls.sheet_names if name != "各种内容"]
    if not sheet_names:
        raise ValueError("Excel文件中缺少主数据工作表")
    main_sheet_name = sheet_names[0]
    df = pd.read_excel(file_path, sheet_name=main_sheet_name)

    # 读取"各种内容"工作表中的所有配置内容
    config_df = pd.read_excel(file_path, sheet_name="各种内容", header=None)

    ##################

    # 从AS列(29列,索引28)获取所有非空选项
    as_options = [str(x).strip() for x in config_df.iloc[1:, 28].dropna().unique()]

    # 创建GUI弹窗让用户选择
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 弹出选择对话框
    selected_as = simpledialog.askstring(
        title="往来款公司选择",
        prompt="请从以下选项中选择要处理的往来款公司:\n" + "\n".join(as_options),
        initialvalue=as_options[0] if as_options else ""
    )

    ##################

    # A列第2行开始：代发其他需要匹配的短语
    replace_phrases = config_df.iloc[1:, 0].dropna().tolist()

    # B列第2行开始：交易失败退款追加内容
    refund_texts = config_df.iloc[1:, 1].dropna().tolist()
    refund_text = refund_texts[0] if refund_texts else "交易失败退款"  # 默认值

    # D列第2行开始：需要替换的手续费类型
    fee_types = config_df.iloc[1:, 3].dropna().tolist()

    # E列第2行开始：手续费类型对应的替换内容
    fee_replacements = config_df.iloc[1:, 4].dropna().tolist()

    # G列第2行开始：需要特殊处理的收(付)方名称（投资款处理）
    target_names = config_df.iloc[1:, 6].dropna().tolist()

    # I列第2行开始：需要特殊处理的往来款公司名称
    # J列第2行开始：往来款公司对应的替换名称
    # transfer_mapping = {}
    # for i in range(1, len(config_df)):
    #     source_name = config_df.iloc[i, 8]  # I列是第8列(0-based)
    #     target_name = config_df.iloc[i, 9]  # J列是第9列(0-based)
    #     if pd.notna(source_name) and pd.notna(target_name):
    #         transfer_mapping[source_name] = target_name

    ###############################
    # 从J列(第10列，索引9)获取所有非空选项作为可选公司列表
    target_companies = [str(x).strip() for x in config_df.iloc[1:, 9].dropna().unique()]

    # 创建GUI弹窗让用户选择替换"锦途"的公司名称
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # BB列(第54列)第2行开始：交易类型筛选条件
    transaction_type_filters = config_df.iloc[1:, 53].dropna().tolist()  # BB列是第54列(0-based)

    # BC列(第55列)第2行开始：对应的摘要替换内容
    transaction_summary_replacements = config_df.iloc[1:, 54].dropna().tolist()  # BC列是第55列(0-based)

    # 创建交易类型到摘要替换的映射字典
    transaction_type_mapping = dict(zip(transaction_type_filters, transaction_summary_replacements))

    # 弹出选择对话框
    selected_company = simpledialog.askstring(
        title="选择替换公司",
        prompt="请从J列中选择要替换'锦途'的公司名称：\n\n可用选项：\n" + "\n".join(target_companies),
        initialvalue=target_companies[0] if target_companies else ""
    )

    if not selected_company or selected_company.strip() not in target_companies:
        messagebox.showinfo("提示", "未选择有效公司或选择取消，程序终止。")
        exit()

    selected_company = selected_company.strip()

    # 建立I列和J列的映射关系
    transfer_mapping = {}
    for i in range(1, len(config_df)):
        source_name = str(config_df.iloc[i, 8]).strip() if pd.notna(config_df.iloc[i, 8]) else None  # I列
        target_name = str(config_df.iloc[i, 9]).strip() if pd.notna(config_df.iloc[i, 9]) else None  # J列
        if source_name and target_name:
            transfer_mapping[source_name] = target_name
    ###############################

    # L列第2行开始：内部转账账号配置
    internal_accounts = config_df.iloc[1:, 11].dropna().tolist()

    # N列第2行开始：需要清除的收(付)方名称内容
    names_to_clear = config_df.iloc[1:, 13].dropna().tolist()

    # P列第2行开始：交易类型筛选条件
    transaction_types = config_df.iloc[1:, 15].dropna().tolist()

    # R列第2行开始：收(付)方名称筛选条件
    payment_names = config_df.iloc[1:, 17].dropna().tolist()

    # T列第2行开始：新增收(付)方名称筛选条件
    payment_names_t = config_df.iloc[1:, 19].dropna().tolist()

    # V列第2行开始：新增摘要筛选条件
    summary_texts = config_df.iloc[1:, 21].dropna().tolist()

    # X列第2行开始：联动代发的收(付)方名称筛选条件
    lian_dong_payment_names = config_df.iloc[1:, 23].dropna().tolist()

    # Z列第2行开始：需要标记为黄色的收(付)方名称
    highlight_names = config_df.iloc[1:, 25].dropna().tolist()

    # AB列(第28列)第2行开始：交易类型关键词
    transaction_keywords = config_df.iloc[1:, 27].dropna().tolist()  # AB列是第28列(0-based)

    # AD列(第30列)第2行开始：行标黄_摘要
    summary_match_list = config_df.iloc[1:, 29].dropna().tolist()  # AD列是第30列(0-based)

    # AF列(第32列)第2行开始：行标黄_收(付)方名称
    gs_gs_wzj = config_df.iloc[1:, 31].dropna().tolist()  # AD列是第30列(0-based)

    # ===== 新增部分：跳过BG列配置的关键词 =====
    # 从"各种内容"工作表BG列获取关键词列表（第59列，索引58）
    bg_keywords = [str(x).strip() for x in config_df.iloc[1:, 58].dropna().unique()]

    # 创建跳过处理的掩码：摘要中包含BG列任一关键词的记录
    skip_mask = df['摘要'].apply(
        lambda x: any(keyword in str(x) for keyword in bg_keywords) if pd.notna(x) else False
    )

    # 创建手续费类型映射字典
    fee_mapping = {}
    for i in range(min(len(fee_types), len(fee_replacements))):
        fee_mapping[fee_types[i]] = fee_replacements[i]
    # 如果手续费类型比替换内容多，使用默认值"银行手续费"
    if len(fee_types) > len(fee_replacements):
        for fee_type in fee_types[len(fee_replacements):]:
            fee_mapping[fee_type] = "银行手续费"

    # 检查必要的列是否存在
    required_columns = ['流程实例号', '业务名称', '摘要', '收(付)方名称', '收(付)方账号', '借方金额', '贷方金额']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"主数据工作表中缺少必要的列: {col}")

    # 创建一个副本用于处理，不影响原始数据
    df_processed = df.copy()

    # 添加跳过处理标记列
    df_processed['_skip_processing'] = skip_mask

    # ===== 新增部分：清除指定收(付)方名称内容 =====
    if names_to_clear:  # 只有当有配置需要清除的名称时才处理
        # 对每个需要清除的内容进行处理
        for text_to_clear in names_to_clear:
            # 筛选出收(付)方名称包含当前要清除文本的记录
            mask = (df_processed['收(付)方名称'].str.contains(text_to_clear, na=False)) & (~df_processed['_skip_processing'])
            # 清除这些记录中收(付)方名称的指定内容
            df_processed.loc[mask, '收(付)方名称'] = df_processed.loc[mask, '收(付)方名称'].str.replace(text_to_clear, '',
                                                                                              regex=False)

    # ===== 第一部分处理：代发其他 =====
    # 找出流程实例号重复2次及以上的所有实例号
    duplicate_counts = df['流程实例号'].value_counts()
    duplicate_instances = duplicate_counts[duplicate_counts >= 2].index

    # 遍历所有重复的流程实例号
    for instance in duplicate_instances:
        # 获取当前流程实例号的所有记录
        group = df[df['流程实例号'] == instance]

        # 筛选出业务名称不为"联动代发"的记录（包括空值）
        target_rows = group[group['业务名称'] != '联动代发']

        # 处理这些行的摘要列
        for idx, row in target_rows.iterrows():
            # 检查摘要是否包含"各种内容"工作表中的任何短语
            if any(phrase in str(row['摘要']) for phrase in replace_phrases):
                # 找到同组的其他行（不包括当前行）
                other_rows = group[group.index != idx]
                if not other_rows.empty:
                    # 取第一行作为替换内容
                    replacement_text = other_rows.iloc[0]['摘要']
                    # 加上从B列读取的退款文本
                    new_text = f"{replacement_text}_{refund_text}"
                    # 检查是否跳过处理
                    if not df_processed.at[idx, '_skip_processing']:
                        df_processed.at[idx, '摘要'] = new_text

    # ===== 第二部分处理：联动代发 =====
    # 筛选业务名称为"联动代发"且收(付)方名称包含"公司"的记录
    mask = ((df_processed['业务名称'] == '联动代发') &
            (df_processed['收(付)方名称'].str.contains('公司', na=False)) &
            (~df_processed['_skip_processing']))
    df_processed.loc[mask, '摘要'] = df_processed.loc[mask, '收(付)方名称']

    # ===== 第三部分处理：银行手续费 =====
    # 筛选并替换摘要内容
    if fee_mapping:  # 只有当手续费类型映射不为空时才处理
        fee_mask = (df_processed['摘要'].isin(fee_mapping.keys()) &
                    (~df_processed['_skip_processing']))
        df_processed.loc[fee_mask, '摘要'] = df_processed.loc[fee_mask, '摘要'].map(fee_mapping)

    # ===== 第四部分处理：投资款 =====
    if target_names:  # 只有当有配置目标名称时才处理
        # 筛选出收(付)方名称包含G列中任一名称的记录
        name_mask = df_processed['收(付)方名称'].isin(target_names)

        # 在这些记录中，进一步筛选摘要包含"投资款"的记录
        invest_mask = df_processed['摘要'].str.contains('投资款', na=False)

        # 同时满足两个条件的记录
        combined_mask = name_mask & invest_mask & (~df_processed['_skip_processing'])

        # 修改这些记录的摘要内容
        df_processed.loc[combined_mask, '摘要'] = "投资款，" + df_processed.loc[combined_mask, '收(付)方名称']

    # ===== 往来款处理 =====
    if transfer_mapping:
        # 获取所有往来款公司名称（排除与selected_as相同的）
        transfer_companies = [name for name in transfer_mapping.keys()
                              if str(name).strip() != str(selected_as).strip()]

        # 筛选条件：
        # 1. 收(付)方名称在I列配置的公司中
        # 2. 且不等于selected_as（第一个弹窗选择的值）
        company_mask = df_processed['收(付)方名称'].apply(
            lambda x: str(x).strip() in transfer_companies
        )

        # 借方金额有内容的记录
        debit_mask = df_processed['借方金额'].notna() & (df_processed['借方金额'] != 0)

        # 贷方金额有内容的记录
        credit_mask = df_processed['贷方金额'].notna() & (df_processed['贷方金额'] != 0)

        # 处理借方金额有内容的记录
        debit_company_mask = company_mask & debit_mask & (~df_processed['_skip_processing'])
        df_processed.loc[debit_company_mask, '摘要'] = df_processed.loc[debit_company_mask].apply(
            lambda row: f"往来款，{selected_company}转{transfer_mapping[str(row['收(付)方名称']).strip()]}",
            axis=1
        )

        # 处理贷方金额有内容的记录
        credit_company_mask = company_mask & credit_mask & (~df_processed['_skip_processing'])
        df_processed.loc[credit_company_mask, '摘要'] = df_processed.loc[credit_company_mask].apply(
            lambda row: f"往来款，{transfer_mapping[str(row['收(付)方名称']).strip()]}转{selected_company}",
            axis=1
        )

        # 统计跳过的记录（收(付)方名称等于selected_as且有金额的记录）
        skipped_mask = ((df_processed['收(付)方名称'].apply(lambda x: str(x).strip() == str(selected_as).strip())) &
                        (df_processed['借方金额'].notna() | df_processed['贷方金额'].notna()))
        skipped_count = skipped_mask.sum()

    # ===== 新增部分：内部转账处理（基于AL列账号匹配）=====
    # 弹出第一个输入框：筛选"收(付)方名称"
    root = tk.Tk()
    root.withdraw()
    input_name_filter = simpledialog.askstring("输入", "请输入筛选内容（用于匹配内部转账'收(付)方名称'）:")

    # 弹出第二个输入框：用于摘要输出
    input_summary_text = simpledialog.askstring("输入", "请输入公司银行账号后四位（用于摘要输出）:")

    if input_name_filter and input_summary_text:  # 确保两个输入都不为空
        # 从"各种内容"工作表AL列获取账号列表（第38列，索引37）
        al_accounts = [str(x).strip() for x in config_df.iloc[1:, 37].dropna().unique()]

        # 筛选条件：
        # 1. "收(付)方名称"包含第一个输入的内容
        # 2. "收(付)方账号"在AL列中
        mask = (
                (df_processed['收(付)方名称'] == input_name_filter) &
                (df_processed['收(付)方账号'].astype(str).str.strip().isin(al_accounts)) &
                (~df_processed['_skip_processing'])
        )

        # 获取符合条件的记录
        filtered_df = df_processed[mask]

        # 处理借方金额有内容的记录
        debit_mask = filtered_df['借方金额'].notna() & (filtered_df['借方金额'] != 0)
        for idx in filtered_df[debit_mask].index:
            account = str(df_processed.at[idx, '收(付)方账号']).strip()
            last_four = account[-4:] if len(account) >= 4 else account  # 取后4位
            df_processed.at[idx, '摘要'] = f"内部转账，{input_summary_text}转入{last_four}"

        # 处理贷方金额有内容的记录
        credit_mask = filtered_df['贷方金额'].notna() & (filtered_df['贷方金额'] != 0)
        for idx in filtered_df[credit_mask].index:
            account = str(df_processed.at[idx, '收(付)方账号']).strip()
            last_four = account[-4:] if len(account) >= 4 else account  # 取后4位
            df_processed.at[idx, '摘要'] = f"内部转账，{last_four}转入{input_summary_text}"

        # 显示处理结果
        processed_count = debit_mask.sum() + credit_mask.sum()
        messagebox.showinfo(
            "内部转账处理完成",
            f"成功处理 {processed_count} 条记录\n"
            f"筛选条件：\n"
            f"- 收(付)方名称包含：'{input_name_filter}'\n"
            f"- 收(付)方账号匹配AL列\n\n"
            f"摘要格式：\n"
            f"1. 借方金额 → '内部转账，{input_summary_text}转入XXXX'\n"
            f"2. 贷方金额 → '内部转账，XXXX转入{input_summary_text}'"
        )

    # ===== 新增部分：劳务费处理 =====
    if transaction_types and payment_names:  # 只有当有配置交易类型和收(付)方名称时才处理
        # 筛选条件：
        # 1. 贷方金额有内容
        # 2. 交易类型在P列配置的列表中
        # 3. 收(付)方名称在R列配置的列表中
        labor_fee_mask = (
                (df_processed['贷方金额'].notna()) &
                (df_processed['贷方金额'] != 0) &
                (df_processed['交易类型'].isin(transaction_types)) &
                (df_processed['收(付)方名称'].isin(payment_names)) &
                (~df_processed['摘要'].str.contains('往来款', na=False)) &  # 新增条件
                (~df_processed['_skip_processing']))  # 跳过处理标记

        # 修改这些记录的摘要内容为"收(付)方名称_劳务费"
        df_processed.loc[labor_fee_mask, '摘要'] = df_processed.loc[labor_fee_mask, '收(付)方名称'] + "_劳务费"

        # 标记借方金额有内容的劳务费记录
        df_processed.loc[labor_fee_mask & (df_processed['借方金额'].notna()) & (df_processed['借方金额'] != 0),
                         '_标记黄色'] = True

    # ===== 新增部分：根据T列和V列筛选并修改摘要 =====
    if payment_names_t and summary_texts:  # 只有当有配置时才处理
        # 筛选条件：
        # 1. 收(付)方名称在T列配置的列表中
        # 2. 摘要在V列配置的列表中
        labor_mask = (
                (df_processed['收(付)方名称'].isin(payment_names_t)) &
                (df_processed['摘要'].isin(summary_texts)) &
                (~df_processed['摘要'].str.contains('往来款', na=False)) &  # 新增条件
                (~df_processed['_skip_processing']))  # 跳过处理标记
        # 修改这些记录的摘要内容为"收(付)方名称_劳务费"
        df_processed.loc[labor_mask, '摘要'] = df_processed.loc[labor_mask, '收(付)方名称'] + "_劳务费"

        # 标记借方金额有内容的劳务费记录
        df_processed.loc[labor_mask & (df_processed['借方金额'].notna()) & (df_processed['借方金额'] != 0),
                         '_标记黄色'] = True

    # ===== 新增部分：联动代发劳务费处理 =====
    if lian_dong_payment_names:  # 只有当有配置联动代发的收(付)方名称时才处理
        # 筛选条件：
        # 1. 业务名称为"联动代发"
        # 2. 收(付)方名称在X列配置的列表中
        lian_dong_mask = (
                (df_processed['业务名称'] == '联动代发') &
                (df_processed['收(付)方名称'].isin(lian_dong_payment_names)) &
                (~df_processed['摘要'].str.contains('往来款', na=False)) &  # 新增条件
                (~df_processed['_skip_processing']))  # 跳过处理标记
        # 修改这些记录的摘要内容为"收(付)方名称_劳务费"
        df_processed.loc[lian_dong_mask, '摘要'] = df_processed.loc[lian_dong_mask, '收(付)方名称'] + "_劳务费"

        # 标记借方金额有内容的劳务费记录
        df_processed.loc[lian_dong_mask & (df_processed['借方金额'].notna()) & (df_processed['借方金额'] != 0),
                         '_标记黄色'] = True

    # ===== 新增部分：锦途代发月薪处理 =====
    # 筛选条件：
    # 1. 业务名称为"联动代发"
    # 2. 摘要不以"_劳务费"结尾（即不是劳务费）
    liandong_mask = (df_processed['业务名称'] == '联动代发')
    non_laborfee_mask = (~df_processed['摘要'].str.endswith('_劳务费', na=False))

    combined_mask = liandong_mask & non_laborfee_mask & (~df_processed['_skip_processing'])

    if combined_mask.any():
        # 弹出输入框让用户输入摘要内容
        root = tk.Tk()
        root.withdraw()
        input_summary = simpledialog.askstring(
            "输入摘要内容",
            "请输入锦途代发摘要内容（例如：锦途代发x月xx员工月薪）:",
            initialvalue="锦途代发x月xx员工月薪"
        )

        if input_summary:  # 如果用户输入了内容
            # 处理交易类型包含"退"字的记录
            refund_mask = (combined_mask &
                           df_processed['交易类型'].str.contains('退', na=False))

            # 处理不包含"退"字的记录
            non_refund_mask = (combined_mask &
                               ~df_processed['交易类型'].str.contains('退', na=False))

            # 设置不包含"退"字的记录的摘要
            df_processed.loc[non_refund_mask, '摘要'] = input_summary

            # 设置包含"退"字的记录的摘要（追加"交易失败退款"）
            df_processed.loc[refund_mask, '摘要'] = input_summary + "_交易失败退款"

            # 显示处理结果
            processed_count = non_refund_mask.sum() + refund_mask.sum()
            refund_count = refund_mask.sum()
            messagebox.showinfo(
                "锦途代发处理完成",
                f"成功处理 {processed_count} 条联动代发记录\n"
                f"- 正常处理: {processed_count - refund_count} 条\n"
                f"- 退款处理: {refund_count} 条\n"
                f"摘要内容: '{input_summary}'"
            )
        else:
            messagebox.showinfo("提示", "未输入摘要内容，跳过锦途代发处理")

    # ===== 新增部分：工伤理赔处理 =====
    # 筛选条件：
    # 1. 收(付)方名称包含"保险"
    # 2. 贷方金额有内容
    insurance_mask = (
            (df_processed['收(付)方名称'].str.contains('保险', na=False)) &
            (df_processed['贷方金额'].notna()) &
            (df_processed['贷方金额'] != 0) &
            (~df_processed['_skip_processing'])  # 跳过处理标记
    )

    if insurance_mask.any():
        # 进一步筛选：摘要中不包含"保费"且不包含"退"的记录（这些记录需要修改）
        to_process_mask = (
                insurance_mask &
                (~df_processed['摘要'].str.contains('保费', na=False)) &
                (~df_processed['摘要'].str.contains('退', na=False))
        )

        # 处理这些记录，直接修改摘要为"员工工伤理赔"
        df_processed.loc[to_process_mask, '摘要'] = "xx员工工伤理赔，xx"

        # 统计处理结果
        processed_count = to_process_mask.sum()
        total_insurance_count = insurance_mask.sum()
        skipped_count = total_insurance_count - processed_count
        messagebox.showinfo(
            "工伤理赔处理完成",
            f"共发现 {total_insurance_count} 条保险相关记录\n"
            f"成功处理 {processed_count} 条记录，摘要改为：员工工伤理赔\n"
            f"跳过 {skipped_count} 条（含保费或退款记录，保持原样）"
        )

    # ===== 新增部分：公积金处理 =====
    # 筛选收(付)方名称中包含"公积金"的记录
    gongjijin_mask = (df_processed['收(付)方名称'].str.contains('公积金', na=False) &
                      (~df_processed['_skip_processing']))  # 跳过处理标记

    # 处理这些记录
    if gongjijin_mask.any():
        # 获取D列内容（从第4列开始，索引为3）
        d_column_content = df_processed.iloc[:, 3].astype(str)

        # 构建新的摘要内容
        new_summary = "缴纳" + d_column_content.str[0:7] + "住房公积金"

        # 清除连字符"-"
        new_summary = new_summary.str.replace('-', '')

        # 更新摘要列
        df_processed.loc[gongjijin_mask, '摘要'] = new_summary[gongjijin_mask]

    # ===== 新增部分：根据BB列交易类型筛选并替换摘要为BC列内容 =====
    if transaction_type_mapping:  # 只有当有配置时才处理
        # 筛选交易类型在BB列配置中的记录
        transaction_mask = (df_processed['交易类型'].isin(transaction_type_filters) &
                            (~df_processed['_skip_processing']))  # 跳过处理标记

        # 替换摘要为BC列对应的内容
        df_processed.loc[transaction_mask, '摘要'] = df_processed.loc[transaction_mask, '交易类型'].map(
            transaction_type_mapping)

    # 删除临时列
    df_processed = df_processed.drop('_skip_processing', axis=1)

    # # ===== 标黄处理部分开始 =====
    # # 创建新工作簿（直接使用openpyxl，不通过pandas保存）
    wb = load_workbook(file_path)

    # 处理主数据工作表
    if main_sheet_name in wb.sheetnames:
        ws = wb[main_sheet_name]
        # 清空原有数据（保留格式）
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(main_sheet_name)

    # 黄色填充样式
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 写入列标题
    for col_num, col_name in enumerate(df_processed.columns, 1):
        if col_name != '_标记黄色':  # 跳过临时列
            ws.cell(row=1, column=col_num, value=col_name)

    # 写入数据并设置格式
    for row_num, row_data in enumerate(dataframe_to_rows(df_processed.drop('_标记黄色', axis=1), index=False, header=False),
                                       2):
        for col_num, cell_value in enumerate(row_data, 1):
            ws.cell(row=row_num, column=col_num, value=cell_value)

        # 检查是否需要标记黄色
        if df_processed.iloc[row_num - 2].get('_标记黄色', False):
            for cell in ws[row_num]:
                cell.fill = yellow_fill

    # 生成最终输出文件名
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)
    output_path = os.path.join(dir_name, f"{name}_完整处理结果{ext}")

    # 创建新工作簿（直接使用openpyxl，不通过pandas保存）
    wb = load_workbook(file_path)

    # 处理主数据工作表
    if main_sheet_name in wb.sheetnames:
        ws = wb[main_sheet_name]
        # 清空原有数据（保留格式）
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(main_sheet_name)

    # 写入列标题
    for col_num, col_name in enumerate(df_processed.columns, 1):
        ws.cell(row=1, column=col_num, value=col_name)

    # 获取"收(付)方名称"列的索引（从1开始计数）
    payment_name_col = df_processed.columns.get_loc('收(付)方名称') + 1

    # 获取交易类型列位置
    trans_type_col = df_processed.columns.get_loc('交易类型') + 1

    # 获取业务名称和借方金额列位置
    biz_name_col = df_processed.columns.get_loc('业务名称') + 1
    debit_amt_col = df_processed.columns.get_loc('借方金额') + 1

    # 获取摘要列位置（列索引+1）
    summary_col = df_processed.columns.get_loc('摘要') + 1

    # 黄色填充样式
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 写入数据并设置格式
    for row_num, row_data in enumerate(dataframe_to_rows(df_processed, index=False, header=False), 2):

        # 初始化标志
        is_lian_dong = False
        has_debit = False
        match_found = False

        for col_num, cell_value in enumerate(row_data, 1):
            ws.cell(row=row_num, column=col_num, value=cell_value)

            # 当处理到"收(付)方名称"列时检查是否包含Z列配置中的关键词
            if col_num == payment_name_col:
                cell_value_str = str(cell_value)  # 确保转换为字符串
                # 检查是否包含任意一个关键词
                if any(keyword in cell_value_str for keyword in highlight_names):
                    # 为整行设置黄色填充
                    for cell in ws[row_num]:
                        cell.fill = yellow_fill

            # 检查交易类型是否包含AB列关键词
            if col_num == trans_type_col:
                cell_value_str = str(cell_value).strip()
                if any(keyword in cell_value_str for keyword in transaction_keywords):
                    # 整行填充黄色
                    for cell in ws[row_num]:
                        cell.fill = yellow_fill

            # 只在摘要列检查匹配
            if col_num == summary_col:
                cell_str = str(cell_value).strip()
                # 检查是否完全匹配AD列任一内容
                if cell_str in summary_match_list:
                    match_found = True
            # 整行填充黄色（在列循环结束后处理）
            if match_found:
                for cell in ws[row_num]:
                    cell.fill = yellow_fill

            # 只在收(付)方名称列检查匹配
            if col_num == payment_name_col:
                cell_str = str(cell_value).strip()
                # 检查是否完全匹配AD列任一内容
                if cell_str in gs_gs_wzj:
                    match_found = True
            # 整行填充黄色（在列循环结束后处理）
            if match_found:
                for cell in ws[row_num]:
                    cell.fill = yellow_fill

            # 检查业务名称
            if col_num == biz_name_col:
                is_lian_dong = (str(cell_value).strip() == "联动代发")

            # 检查借方金额
            if col_num == debit_amt_col:
                has_debit = (not pd.isna(cell_value)) and (cell_value != 0)

        # 两个条件同时满足时填充黄色
        if is_lian_dong and has_debit:
            for cell in ws[row_num]:
                cell.fill = yellow_fill

    # 保存结果
    wb.save(output_path)
    print(f"处理完成，结果已保存到: {output_path}")


if __name__ == "__main__":
    try:
        file_path = select_file()
        if file_path:
            process_excel(file_path)
        else:
            print("未选择文件")
    except Exception as e:
        print(f"处理出错: {str(e)}")