"""
Excel汇总处理工具
根据备注2列内容缩写，汇总打款金额，生成供应商汇总表
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import re
import traceback
from typing import Dict, List, Set, Tuple, Any, Optional
import warnings
warnings.filterwarnings('ignore')


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
        self.workbook: Optional[openpyxl.Workbook] = None
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
            self.workbook = load_workbook(file_path, data_only=True)
            self.sheet_names = self.workbook.sheetnames
            return True, f"成功读取文件：{os.path.basename(file_path)}", self.sheet_names
        except Exception as e:
            # 关闭工作簿释放资源
            if hasattr(self, 'workbook') and self.workbook:
                self.workbook.close()
                self.workbook = None
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
    
    def process_data(self) -> Tuple[bool, str]:
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
            
            # 检测"打款金额"列（H列）
            amount_cols = [i for i, col in enumerate(df.columns) if col == '打款金额']
            
            if not amount_cols:
                # 没有列名匹配，尝试使用固定位置（索引7，即第8列）
                if df.shape[1] > 7:
                    amount_col_idx = 7
                else:
                    return False, "数据列数不足，无法确定'打款金额'列位置"
            elif len(amount_cols) >= 2:
                # 有多个"打款金额"列，选择第二个（H列）
                amount_col_idx = amount_cols[1]
            else:
                amount_col_idx = amount_cols[0]
            
            # 应用缩写规则
            df['公司缩写'] = df['备注2'].apply(self._abbreviate_company_name)
            
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
            # 只在"总计"列和"公司"列之间保留一个空列
            final_columns = ['供应商名称'] + unique_companies + ['总计', '', '公司']
            self.output_data = self.output_data.reindex(columns=final_columns)
            
            # 获取H列名称
            #h_col_name = df.columns[amount_col_idx] if amount_col_idx < len(df.columns) else f'第{amount_col_idx + 1}列'
            
            return True, f"处理完成！共{len(self.output_data)}个供应商，{len(unique_companies)}个公司"
            
        except Exception as e:
            import traceback
            error_msg = f"数据处理失败：{str(e)}\n{traceback.format_exc()}"
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
        keywords = ["刘先锋现金卡支付", "个体工商户", "有限公司", "分公司"]
        
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
    
    def save_to_excel(self, save_path: str) -> Tuple[bool, str]:
        """
        保存到Excel文件
        
        Args:
            save_path: 保存路径
            
        Returns:
            (success, message): 是否成功、消息
        """
        try:
            if self.output_data is None:
                return False, "没有数据可保存"
            
            # 使用openpyxl创建新工作簿
            wb = openpyxl.Workbook()
            ws = wb.active
            # 从当前工作表名称提取日期作为工作表标题
            sheet_title = self._extract_date_from_sheet_name(self.current_sheet)
            ws.title = sheet_title
            
            # 写入表头（第1行）
            columns = list(self.output_data.columns)
            for col_idx, col_name in enumerate(columns, start=1):
                ws.cell(row=1, column=col_idx, value=col_name)
            
            # 写入数据（从第2行开始）
            for row_idx, row in enumerate(self.output_data.itertuples(index=False), start=2):
                for col_idx, value in enumerate(row, start=1):
                    # 如果数值为0或空字符串，则显示为空白
                    if isinstance(value, (int, float)) and value == 0:
                        ws.cell(row=row_idx, column=col_idx, value=None)
                    elif pd.isna(value):
                        ws.cell(row=row_idx, column=col_idx, value=None)
                    elif value == '':
                        ws.cell(row=row_idx, column=col_idx, value=None)
                    else:
                        ws.cell(row=row_idx, column=col_idx, value=value)
            
            # 设置样式
            self._apply_excel_styles(ws, self.output_data)
            
            # 保存
            wb.save(save_path)
            return True, f"文件已保存到：{save_path}"
            
        except Exception as e:
            import traceback
            error_msg = f"保存失败：{str(e)}\n{traceback.format_exc()}"
            return False, error_msg
    
    def _apply_excel_styles(self, ws: openpyxl.worksheet.worksheet.Worksheet, df: pd.DataFrame) -> None:
        """
        应用Excel样式
        
        Args:
            ws: 工作表对象
            df: 数据框
        """
        # 设置标题行样式
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for col in range(1, df.shape[1] + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        
        # 设置数据行样式
        for row_idx in range(2, df.shape[0] + 2):
            for col_idx in range(1, df.shape[1] + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
        
        # 自动调整列宽
        for col in range(1, df.shape[1] + 1):
            max_length = 0
            column_letter = get_column_letter(col)
            
            # 检查标题
            cell = ws.cell(row=1, column=col)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
            
            # 检查数据
            for row_idx in range(2, df.shape[0] + 2):
                cell = ws.cell(row=row_idx, column=col)
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            # 设置列宽（最大30，最小10）
            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[column_letter].width = adjusted_width


class SummaryProcessorApp:
    """汇总处理GUI应用程序"""
    
    def __init__(self, root: tk.Tk):
        """
        初始化应用程序
        
        Args:
            root: Tkinter根窗口
        """
        self.root = root
        self.root.title("Excel汇总处理工具")
        self.root.geometry("1100x800")
        self.root.minsize(900, 600)
        
        # 配置窗口可调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        self.processor = SummaryProcessor()
        self.current_sheet = tk.StringVar()
        
        # 配置Treeview样式，减小行高
        self._setup_treeview_style()
        
        self._create_ui()
    
    def _setup_treeview_style(self):
        """配置Treeview样式，优化显示效果"""
        style = ttk.Style()
        
        # 设置Treeview的行高
        style.configure("Treeview", 
                       rowheight=25,  # 设置行高
                       font=('Microsoft YaHei UI', 9))
        
        # 设置标题样式
        style.configure("Treeview.Heading", 
                       font=('Microsoft YaHei UI', 10, 'bold'))
        
        # 设置选中项样式
        style.map("Treeview",
                 background=[('selected', '#0078d7')],
                 foreground=[('selected', 'white')])
    
    def _create_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置主框架的网格权重
        main_frame.columnconfigure(0, weight=1)
        
        # ========== 文件选择区域 ==========
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="12")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        file_frame.columnconfigure(0, weight=1)
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, font=('Microsoft YaHei UI', 9))
        file_entry.grid(row=0, column=0, padx=(0, 10), sticky=(tk.W, tk.E))
        
        browse_btn = ttk.Button(file_frame, text="浏览...", command=self._browse_file, width=12)
        browse_btn.grid(row=0, column=1, padx=(0, 15))
        
        ttk.Label(file_frame, text="工作表:", font=('Microsoft YaHei UI', 9)).grid(row=0, column=2, padx=(0, 5))
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.current_sheet, 
                                       state='readonly', width=25, font=('Microsoft YaHei UI', 9))
        self.sheet_combo.grid(row=0, column=3)
        
        # ========== 操作按钮区域 ==========
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        
        process_btn = ttk.Button(button_frame, text="📊数据处理", command=self._process_data, width=18)
        process_btn.grid(row=0, column=0, padx=(0, 10))
        
        save_btn = ttk.Button(button_frame, text="💾文件保存", command=self._save_file, width=18)
        save_btn.grid(row=0, column=1, padx=(0, 10))
        
        clear_btn = ttk.Button(button_frame, text="🗑️清空", command=self._clear_all, width=18)
        clear_btn.grid(row=0, column=2)
        
        # ========== 数据预览区域 ==========
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览", padding="12")
        preview_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 12))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # 创建表格
        self.tree = ttk.Treeview(preview_frame, selectmode='browse', show='headings')
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 配置Treeview的滚动命令
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 初始化滚动位置到最左边
        self.tree.xview_moveto(0)
        
        # 绑定鼠标滚轮事件
        self.tree.bind("<MouseWheel>", self._on_mousewheel)
        self.tree.bind("<Button-4>", self._on_mousewheel)
        self.tree.bind("<Button-5>", self._on_mousewheel)
        
        # ========== 状态信息栏 ==========
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="12")
        status_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, 
                                                   font=('Consolas', 9))
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重（让数据预览和状态信息自动调整大小）
        main_frame.rowconfigure(2, weight=3)  # 数据预览占3份
        main_frame.rowconfigure(3, weight=1)  # 状态信息占1份
        
        # 配置文件选择区域的网格权重
        file_frame.columnconfigure(0, weight=1)
        
        self._log_status("欢迎使用Excel汇总处理工具！")
        self._log_status("请选择一个Excel文件开始处理。")
    
    def _on_mousewheel(self, event):
        """
        处理鼠标滚轮事件，支持表格滚动
        
        Args:
            event: 滚轮事件
        """
        if event.num == 4 or event.delta > 0:
            self.tree.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.tree.yview_scroll(1, "units")
    
    def _log_status(self, message: str):
        """
        记录状态信息
        
        Args:
            message: 状态消息
        """
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()
    
    def _browse_file(self):
        """浏览并选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self._log_status(f"正在读取文件：{os.path.basename(file_path)}...")
            success, message, sheets = self.processor.load_file(file_path)
            
            if success:
                self.file_path_var.set(file_path)
                self.sheet_combo['values'] = sheets
                if sheets:
                    self.sheet_combo.current(0)
                self._log_status(message)
                self._log_status(f"包含{len(sheets)}个工作表")
            else:
                messagebox.showerror("错误", message)
                self._log_status(f"错误：{message}")
    
    def _process_data(self):
        """执行数据处理"""
        file_path = self.file_path_var.get()
        sheet_name = self.current_sheet.get()
        
        if not file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        if not sheet_name:
            messagebox.showwarning("警告", "请选择工作表")
            return
        
        self._log_status("=" * 50)
        self._log_status("开始数据处理...")
        
        # 加载工作表
        self._log_status(f"1. 加载工作表：{sheet_name}")
        success, message = self.processor.load_sheet(sheet_name)
        
        if not success:
            messagebox.showerror("错误", message)
            self._log_status(f"错误：{message}")
            return
        
        self._log_status(message)
        
        # 处理数据
        self._log_status("2. 检测列...")
        if "备注2" not in self.processor.data_frame.columns:
            messagebox.showerror("错误", "未找到'备注2'列")
            self._log_status("错误：未找到'备注2'列")
            return
        
        self._log_status("   ✓ 找到'备注2'列")
        
        if "供应商名称" not in self.processor.data_frame.columns:
            messagebox.showerror("错误", "未找到'供应商名称'列")
            self._log_status("错误：未找到'供应商名称'列")
            return
        
        self._log_status("   ✓ 找到'供应商名称'列")
        
        if "收款姓名" not in self.processor.data_frame.columns:
            messagebox.showerror("错误", "未找到'收款姓名'列")
            self._log_status("错误：未找到'收款姓名'列")
            return
        
        self._log_status("   ✓ 找到'收款姓名'列")
        
        self._log_status("3. 应用缩写规则...")
        success, message = self.processor.process_data()
        
        if not success:
            messagebox.showerror("错误", message)
            self._log_status(f"错误：{message}")
            return
        
        self._log_status(message)
        
        # 显示数据预览
        self._log_status("4. 生成汇总表...")
        self._show_preview()
        
        self._log_status("5. 数据处理完成！")
        self._log_status("=" * 50)
        
        messagebox.showinfo("成功", "数据处理完成！\n可以点击'文件保存'按钮保存结果。")
    
    def _show_preview(self):
        """显示数据预览"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.processor.output_data is None:
            return
        
        df = self.processor.output_data
        
        # 设置列
        columns = list(df.columns)
        self.tree['columns'] = columns
        
        # 设置列标题和列宽
        for col in columns:
            self.tree.heading(col, text=col)
            
            # 根据列名动态设置列宽（优化宽度以减少横向滚动）
            if col == '供应商名称':
                self.tree.column(col, width=180, anchor=tk.W)
            elif col == '公司':
                self.tree.column(col, width=250, anchor=tk.W)
            elif col == '总计':
                self.tree.column(col, width=100, anchor=tk.E)
            elif col == '':
                self.tree.column(col, width=30, anchor=tk.CENTER)
            else:
                # 公司列
                self.tree.column(col, width=90, anchor=tk.E)
        
        # 显示数据（最多显示50行）
        display_rows = min(50, len(df))
        for row in df.head(display_rows).itertuples(index=False):
            # 处理空值和零值
            processed_row = tuple(
                '' if (isinstance(v, (int, float)) and v == 0) 
                else v if pd.notna(v) 
                else '' 
                for v in row
            )
            self.tree.insert('', tk.END, values=processed_row)
        
        # 强制更新界面
        self.tree.update_idletasks()
        
        # 滚动到最左边
        self.tree.xview_moveto(0)
        
        self._log_status(f"   ✓ 预览显示前{display_rows}行数据")
        self._log_status(f"   ✓ 共{len(df)}行，{len(columns)}列")
    
    def _save_file(self):
        """保存文件"""
        if self.processor.output_data is None:
            messagebox.showwarning("警告", "没有数据可保存，请先进行数据处理")
            return
        
        # 从工作表名称提取日期部分作为文件名
        sheet_name = self.current_sheet.get()
        default_name = self.processor._extract_date_from_sheet_name(sheet_name, include_timestamp=True)
        
        file_path = filedialog.asksaveasfilename(
            title="保存文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=default_name
        )
        
        if file_path:
            self._log_status("正在保存文件...")
            success, message = self.processor.save_to_excel(file_path)
            
            if success:
                self._log_status(message)
                messagebox.showinfo("成功", message)
            else:
                messagebox.showerror("错误", message)
                self._log_status(f"错误：{message}")
    
    def _clear_all(self):
        """清空所有数据"""
        self.file_path_var.set("")
        self.current_sheet.set("")
        self.sheet_combo['values'] = []
        
        # 清空预览
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 重置处理器
        self.processor = SummaryProcessor()
        
        self._log_status("已清空所有数据")


def main():
    """主函数"""
    root = tk.Tk()
    app = SummaryProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
