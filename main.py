# -- coding: utf-8 --
# @Time : 2025-09-22 14:19
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : main.py
# @Software: PyCharm

from cf_duoge_wenj import split_excel_by_column as split_to_files #拆分_拆分成多个表格文件
from cw_kgz_sx import select_excel_file #财务_开关账_收支记账_筛选
from cf_biaoge_li import split_excel_by_column #拆分_拆分在表格里
from Ncengjiwenjj_fzzt import main_zfzwj #N层级文件夹_复制粘贴_只要获取文件
from cf_duoge_wenj_zz import split_excel_by_column_zcfgzb #拆分_拆分两次_在职离职_吴苏霞
from shouxufei_cf import shouxufei_cf_xzy #许子怡的
from cw_gys_qzy import run_excel_merger_qzy #财务，供应商，戚倬悦，就合并Excel
from cf_zdnrfyq_nrcs_wsx import main_cf_zdnr_ygbg #拆分_指定内容放一起_拆分参数_吴苏霞
from cw_jxkh_zhf import PerformanceApp # （绩效考核）数据匹配_稳岗率计算_合并表格_张惠芳
from cw_yzpz_pmh_main import ExcelProcessorGUI # 预支平账_潘墨涵
from hs_huizong_cxm_main import SummaryProcessorApp # 核算_汇总_陈小卯

import tkinter as tk
from tkinter import messagebox, scrolledtext  # 使用 scrolledtext 支持滚动条


class MainToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("浙江锦途 - 处理Excel的Python脚本")
        self.root.geometry("900x800")  # 增加高度以容纳更多内容
        self.root.resizable(False, False)
        self.root.configure(bg="#F5F5F5")##f0f0f0 #F5F5F5

        # 创建界面
        self.create_widgets()

        # 标记功能按钮是否已显示
        self.function_buttons_visible = False

    def create_widgets(self):
        """创建所有UI控件"""
        # 主框架，用于左右布局
        main_frame = tk.Frame(self.root, bg="white")##f0f0f0
        main_frame.pack(pady=20, padx=20, fill="both", expand=False)

        # 左侧：公告栏文本框（宽度20字符，高度8行）
        announce_frame = tk.LabelFrame(main_frame, text="📢 公告栏", font=("微软雅黑", 20, "bold"), bg="white", bd=2)
        announce_frame.grid(row=0, column=0, rowspan=2, padx=(0, 20), sticky="n")

        self.announce_text = scrolledtext.ScrolledText(
            announce_frame,
            width=30,
            height=12,
            font=("微软雅黑", 12),
            bg="white",
            fg="black",
            wrap="word"
        )
        self.announce_text.pack(padx=5, pady=5)
        self.announce_text.insert("1.0", "需要添加需求,\n请在企业微信上搜索：\n王智利\n进行反馈。")
        self.announce_text.config(state="disabled")  # 禁止编辑

        # 右侧：标题
        title_frame = tk.Frame(main_frame, bg="white")
        title_frame.grid(row=0, column=1, sticky="w")

        tk.Label(
            title_frame,
            text="        浙江锦途专用\n处理Excel的Python脚本",
            font=("微软雅黑", 18, "bold"),
            fg="darkblue",
            bg="white",
            justify="left"
        ).pack(anchor="w")

        # 主入口按钮区域
        self.main_button_frame = tk.Frame(self.root, bg="#F5F5F5")
        self.main_button_frame.pack(pady=10)

        # 主入口按钮 - 通用功能
        self.general_button = tk.Button(
            self.main_button_frame,
            text="通用功能",
            font=("微软雅黑", 14, "bold"),
            width=15,
            bg="lightblue",
            fg="black",
            relief="raised",
            bd=3,
            command=self.show_general_functions
        )
        self.general_button.grid(row=0, column=0, padx=10, pady=10)

        # 主入口按钮 - 财务功能
        self.finance_button = tk.Button(
            self.main_button_frame,
            text="财务功能",
            font=("微软雅黑", 14, "bold"),
            width=15,
            bg="lightblue",#lightgreen
            fg="black",
            relief="raised",
            bd=3,
            command=self.show_finance_functions
        )
        self.finance_button.grid(row=0, column=1, padx=10, pady=10)

        # 主入口按钮 - 人事功能
        self.rs_zygn = tk.Button(
            self.main_button_frame,
            text="人事功能",
            font=("微软雅黑", 14, "bold"),
            width=15,
            bg="lightblue",#lightgreen
            fg="black",
            relief="raised",
            bd=3,
            command=self.rs_zygn_dan
        )
        self.rs_zygn.grid(row=0, column=2, padx=10, pady=10)

        # 功能按钮区域框架
        self.function_frame = tk.Frame(self.root, bg="#f0f0f0")
        self.function_frame.pack(pady=10)

        # 创建功能按钮但不显示
        self.create_function_buttons()

        # 退出按钮
        self.quit_button = tk.Button(
            self.root,
            text="退出",
            font=("微软雅黑", 10),
            width=10,
            bg="#f8f9fa",
            fg="red",
            command=self.root.quit
        )
        self.quit_button.pack(pady=20)

    def create_function_buttons(self):
        """创建功能按钮但不显示"""
        # 通用功能框架
        self.general_frame = tk.LabelFrame(self.function_frame, text="通用功能", font=("微软雅黑", 12, "bold"), bg="#f0f0f0",
                                           bd=2)

        # 创建内部框架用于多列布局
        inner_frame = tk.Frame(self.general_frame, bg="#f0f0f0")
        inner_frame.pack(padx=10, pady=10)

        # 第一列
        col1_frame = tk.Frame(inner_frame, bg="#f0f0f0")
        col1_frame.grid(row=0, column=0, padx=10, pady=5)


        # 按钮1 - 拆分_拆成多个Excel文件
        self.btn1 = tk.Button(
            col1_frame,
            text="1、拆分_拆成多个Excel文件",
            font=("微软雅黑", 13),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.run_split_to_files
        )
        self.btn1.pack(pady=10)

        # 按钮3 - N层级文件夹_复制粘贴_只要获取文件
        self.btn4 = tk.Button(
            col1_frame,
            text="2、N层级文件夹中_只获取文件",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.run_split_to_zhqwj
        )
        self.btn4.pack(pady=10)

        # 第二列
        col2_frame = tk.Frame(inner_frame, bg="#f0f0f0")
        col2_frame.grid(row=1, column=0, padx=10, pady=5)

        # 按钮2 - 拆分_按列拆为多个工作表
        self.btn3 = tk.Button(
            col2_frame,
            text="3、拆分_拆成多个工作表",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.run_split_to_sheets
        )
        self.btn3.pack(pady=10)



        # 财务功能框架
        self.finance_frame = tk.LabelFrame(self.function_frame, text="财务功能", font=("微软雅黑", 12, "bold"), bg="#f0f0f0",
                                           bd=2)

        # 创建内部框架用于多列布局
        inner_frame_cwzy = tk.Frame(self.finance_frame, bg="#f0f0f0")
        inner_frame_cwzy.pack(padx=10, pady=10)

        # 第一列
        col1_frame_cwzy = tk.Frame(inner_frame_cwzy, bg="#f0f0f0")
        col1_frame_cwzy.grid(row=0, column=0, padx=10, pady=5)

        # 按钮1 - 财务_收支记账_筛选_开关账
        self.btn_cw_1 = tk.Button(
            col1_frame_cwzy,
            text="1、财务_收支记账_筛选",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.run_financial
        )
        self.btn_cw_1.pack(pady=10)

        # 按钮2 - 供应商_合并表格
        self.btn_cw_2 = tk.Button(
            col1_frame_cwzy,
            text="2、供应商_合并表格",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.cwzy_gys_hbbg
        )
        self.btn_cw_2.pack(pady=10)

        # 按钮3 - 出纳_手续费
        self.btn_cw_3 = tk.Button(
            col1_frame_cwzy,
            text="3、出纳_手续费",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.run_financial_shouxufei_cf_xzy
        )
        self.btn_cw_3.pack(pady=10)

        # 第二列
        col2_frame_cwzy = tk.Frame(inner_frame_cwzy, bg="#f0f0f0")
        col2_frame_cwzy.grid(row=0, column=1, padx=10, pady=5)

        # 按钮4 - 绩效考核
        self.btn_cw_4 = tk.Button(
            col2_frame_cwzy,
            text="4、绩效考核",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.open_cw_jxkh
        )
        self.btn_cw_4.pack(pady=10)

        # 按钮5 - 预支平账
        self.btn_cw_5 = tk.Button(
            col2_frame_cwzy,
            text="5、预支平账",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.open_cw_yzpz
        )
        self.btn_cw_5.pack(pady=10)

        # 按钮6 - 核算
        self.btn_cw_6 = tk.Button(
            col2_frame_cwzy,
            text="6、核算",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.open_hs_hz
        )
        self.btn_cw_6.pack(pady=10)

        # 人事功能框架
        self.rs_zygn_dkj = tk.LabelFrame(self.function_frame, text="人事功能", font=("微软雅黑", 12, "bold"), bg="#f0f0f0",
                                           bd=2)

        # 创建内部框架用于多列布局
        inner_frame_cwzy = tk.Frame(self.general_frame, bg="#f0f0f0")
        inner_frame_cwzy.pack(padx=10, pady=10)

        # 第一列
        col1_frame_rszy = tk.Frame(self.rs_zygn_dkj, bg="#f0f0f0")
        col1_frame_rszy.grid(row=0, column=0, padx=10, pady=5)

        # 按钮1 - 拆分_拆分成多个表格文件_在职离职_吴苏霞
        self.btn_rs_1 = tk.Button(
            col1_frame_rszy,
            text="1、拆分_拆分两次_在职离职",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.rszy_cf_lc
        )
        self.btn_rs_1.pack(pady=10)

        # 按钮2 - 拆分_指定内容放一起_拆分参数_吴苏霞
        self.btn_rs_2 = tk.Button(
            col1_frame_rszy,
            text="2、拆分_指定内容放一起_拆分参数",
            font=("微软雅黑", 12),
            width=25,
            bg="white",
            fg="black",
            relief="groove",
            bd=2,
            anchor="w",
            command=self.rszy_cf_cfcs_wsx
        )
        self.btn_rs_2.pack(pady=10)


    def show_general_functions(self):
        """显示通用功能"""
        # 隐藏所有功能框架
        self.hide_all_functions()
        # 显示通用功能框架
        self.general_frame.grid(row=0, column=0, padx=20, pady=10)
        self.function_buttons_visible = True

    def show_finance_functions(self):
        """显示财务功能"""
        # 隐藏所有功能框架
        self.hide_all_functions()
        # 显示财务功能框架
        self.finance_frame.grid(row=0, column=0, padx=20, pady=10)
        self.function_buttons_visible = True

    def rs_zygn_dan(self):
        """显示人事功能"""
        # 隐藏所有功能框架
        self.hide_all_functions()
        # 显示人事功能框架
        self.rs_zygn_dkj.grid(row=0, column=0, padx=20, pady=10)
        self.function_buttons_visible = True

    def hide_all_functions(self):
        """隐藏所有功能框架"""
        self.general_frame.grid_forget()
        self.finance_frame.grid_forget()
        self.rs_zygn_dkj.grid_forget()
        self.function_buttons_visible = False

    def run_split_to_files(self):
        try:
            split_to_files()
        except Exception as e:
            messagebox.showerror("错误", f"拆分为多个文件时出错: {str(e)}")

    def run_financial(self):
        try:
            select_excel_file()
        except Exception as e:
            messagebox.showerror("错误", f"财务功能出错: {str(e)}")

    def run_split_to_sheets(self):
        try:
            split_excel_by_column()
        except Exception as e:
            messagebox.showerror("错误", f"拆分为多个工作表时出错: {str(e)}")

    def run_split_to_zhqwj(self):
        try:
            main_zfzwj()
        except Exception as e:
            messagebox.showerror("错误", f"代码里的函数不对: {str(e)}")

    def rszy_cf_lc(self):
        try:
            split_excel_by_column_zcfgzb()
        except Exception as e:
            messagebox.showerror("错误", f"代码里的函数不对: {str(e)}")

    def run_financial_shouxufei_cf_xzy(self):
        try:
            shouxufei_cf_xzy()
        except Exception as e:
            messagebox.showerror("错误", f"代码里的函数不对: {str(e)}")

    def cwzy_gys_hbbg(self):
        try:
            run_excel_merger_qzy()
        except Exception as e:
            messagebox.showerror("错误", f"代码里的函数不对: {str(e)}")

    def rszy_cf_cfcs_wsx(self):
        try:
            main_cf_zdnr_ygbg()
        except Exception as e:
            messagebox.showerror("错误", f"代码里的函数不对: {str(e)}")

    def open_cw_jxkh(self):
        """打开绩效考核处理工具"""
        try:
            window = tk.Toplevel(self.root)
            window.title("绩效考核处理工具")
            window.geometry("700x700")
            app = PerformanceApp(window)
        except Exception as e:
            messagebox.showerror("错误", f"打开绩效考核工具失败: {str(e)}")

    def open_cw_yzpz(self):
        """打开预支平账处理工具"""
        try:
            window = tk.Toplevel(self.root)
            app = ExcelProcessorGUI(window)
        except Exception as e:
            messagebox.showerror("错误", f"打开预支平账工具失败: {str(e)}")

    def open_hs_hz(self):
        """打开核算汇总处理工具"""
        try:
            window = tk.Toplevel(self.root)
            app = SummaryProcessorApp(window)
        except Exception as e:
            messagebox.showerror("错误", f"打开核算工具失败: {str(e)}")

# —— 程序启动 ——
if __name__ == "__main__":
    root = tk.Tk()
    app = MainToolApp(root)
    root.mainloop()