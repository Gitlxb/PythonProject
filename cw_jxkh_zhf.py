#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
绩效考核处理工具 - 主界面
整合数据匹配、稳岗率计算、合并表格三大功能
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os

# 导入三个子模块
from cw_sjpp_zhf import App as DataMatchApp
from cw_wgl_main_zhf import ExcelAnalysisApp as StabilityApp
from cw_hbbg_zhf import ExcelMergerApp as MergerApp


class PerformanceApp:
    """绩效考核处理工具主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("绩效考核处理工具")
        self.root.geometry("700x700")
        self.root.resizable(False, False)

        # 创建界面
        self.create_widgets()

    def create_widgets(self):
        """创建界面组件"""
        # 标题
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill=tk.X, pady=(40, 20))

        title_label = ttk.Label(
            title_frame,
            text="绩效考核处理工具",
            font=("微软雅黑", 28, "bold")
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            title_frame,
            text="数据匹配 · 稳岗率计算 · 表格合并",
            font=("微软雅黑", 12),
            foreground="#666666"
        )
        subtitle_label.pack(pady=(5, 0))

        # 分隔线
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, padx=50, pady=20)

        # 按钮框架
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.BOTH, expand=True, padx=100, pady=20)

        # 按钮样式配置
        button_config = {
            'width': 40,
            'height': 2,
            'font': ('微软雅黑', 14, 'bold')
        }

        # 按钮 1：数据匹配
        self.btn_data_match = tk.Button(
            btn_frame,
            text="📊 数据匹配",
            command=self.open_data_match,
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=3,
            **button_config
        )
        self.btn_data_match.pack(pady=20)

        # 按钮 2：稳岗率计算
        self.btn_stability = tk.Button(
            btn_frame,
            text="📈 稳岗率计算",
            command=self.open_stability,
            bg="#2196F3",
            fg="white",
            activebackground="#1976D2",
            activeforeground="white",
            relief=tk.RAISED,
            bd=3,
            **button_config
        )
        self.btn_stability.pack(pady=20)

        # 按钮 3：合并表格
        self.btn_merge = tk.Button(
            btn_frame,
            text="📋 合并表格",
            command=self.open_merge,
            bg="#FF9800",
            fg="white",
            activebackground="#F57C00",
            activeforeground="white",
            relief=tk.RAISED,
            bd=3,
            **button_config
        )
        self.btn_merge.pack(pady=20)

        # 底部说明
        note_frame = ttk.Frame(self.root)
        note_frame.pack(fill=tk.X, pady=(20, 10), padx=50)

        note_text = (
            "使用说明：\n"
            "1. 数据匹配：从原始表自动生成完整版表格，包含公式链和数据关联\n"
            "2. 稳岗率计算：分析项目驻场和项目经理的稳岗率，生成奖励汇总\n"
            "3. 合并表格：将多个Excel文件合并到一个汇总表中，保留原始格式"
        )
        note_label = ttk.Label(
            note_frame,
            text=note_text,
            font=("微软雅黑", 9),
            foreground="#888888",
            justify=tk.LEFT
        )
        note_label.pack(anchor=tk.W)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪 - 请选择要使用的功能")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(10, 5),
            font=('微软雅黑', 9)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def open_data_match(self):
        """打开数据匹配工具"""
        self.status_var.set("正在打开数据匹配工具...")
        try:
            # 创建新窗口
            data_match_window = tk.Toplevel(self.root)
            data_match_window.title("数据匹配 - 完整版表格自动生成工具")
            data_match_window.geometry("820x530")
            
            # 创建数据匹配应用
            app = DataMatchApp(data_match_window)
            
            self.status_var.set("数据匹配工具已打开")
        except Exception as e:
            messagebox.showerror("错误", f"打开数据匹配工具失败：{str(e)}")
            self.status_var.set("打开数据匹配工具失败")

    def open_stability(self):
        """打开稳岗率计算工具"""
        self.status_var.set("正在打开稳岗率计算工具...")
        try:
            # 创建新窗口
            stability_window = tk.Toplevel(self.root)
            stability_window.title("稳岗率计算 - 分析工具")
            stability_window.geometry("600x650")
            
            # 创建稳岗率应用
            app = StabilityApp(stability_window)
            
            self.status_var.set("稳岗率计算工具已打开")
        except Exception as e:
            messagebox.showerror("错误", f"打开稳岗率计算工具失败：{str(e)}")
            self.status_var.set("打开稳岗率计算工具失败")

    def open_merge(self):
        """打开合并表格工具"""
        self.status_var.set("正在打开合并表格工具...")
        try:
            # 创建新窗口
            merge_window = tk.Toplevel(self.root)
            merge_window.title("合并表格 - Excel表格合并工具")
            merge_window.geometry("750x550")
            
            # 创建合并表格应用
            app = MergerApp(merge_window)
            
            self.status_var.set("合并表格工具已打开")
        except Exception as e:
            messagebox.showerror("错误", f"打开合并表格工具失败：{str(e)}")
            self.status_var.set("打开合并表格工具失败")


def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置窗口样式
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    
    # 创建主应用
    app = PerformanceApp(root)
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
