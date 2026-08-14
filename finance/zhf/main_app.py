#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
绩效考核处理工具 - 主界面
左侧导航栏 + Frame切换 方案
整合：数据匹配、稳岗率计算、合并表格、在职月份计算 四大功能
"""

import tkinter as tk
from tkinter import ttk, messagebox


class MainApp:
    """绩效考核处理工具主界面 - 左侧导航+Frame切换"""

    def __init__(self, root):
        self.root = root
        self.root.title("绩效考核处理工具")
        self.root.geometry("900x650")
        self.root.minsize(800, 550)

        # 当前激活的Frame
        self.current_frame = None
        self.frame_instances = {}

        # 导航按钮引用
        self.nav_buttons = {}

        # 创建界面
        self._create_ui()

    def _create_ui(self):
        """创建界面布局"""
        self.root.configure(bg="#2c3e50")

        # 主水平容器
        main_container = tk.Frame(self.root, bg="#2c3e50")
        main_container.pack(fill=tk.BOTH, expand=True)

        # ===== 左侧导航栏（深色背景）=====
        nav_frame = tk.Frame(main_container, bg="#2c3e50", width=180)
        nav_frame.pack(side="left", fill="y")
        nav_frame.pack_propagate(False)

        # 导航标题
        tk.Label(
            nav_frame, text="功能菜单",
            font=("Microsoft YaHei", 14, "bold"), fg="white", bg="#2c3e50",
            pady=20
        ).pack()

        # 分隔线
        ttk.Separator(nav_frame, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)

        # 导航按钮配置
        nav_items = [
            {"key": "sjpp", "text": "📊 数据匹配", "desc": "完整版表格自动生成"},
            {"key": "wgl", "text": "📈 稳岗率计算", "desc": "驻场&项目经理分析"},
            {"key": "hbbg", "text": "📋 合并表格", "desc": "多Excel文件合并"},
            {"key": "tenure", "text": "📅 在职月份计算", "desc": "员工在职月份/天数统计"},
        ]

        btn_container = tk.Frame(nav_frame, bg="#2c3e50")
        btn_container.pack(fill="x", padx=12, pady=10)

        for item in nav_items:
            btn = tk.Button(
                btn_container,
                text=item["text"],
                command=lambda k=item["key"]: self._switch_frame(k),
                font=("Microsoft YaHei", 11),
                fg="white", bg="#34495e",
                activebackground="#3498db", activeforeground="white",
                relief="flat", bd=0,
                anchor="w", width=18, height=2,
                cursor="hand2",
            )
            btn.pack(pady=4, fill="x")
            self.nav_buttons[item["key"]] = btn

            # 描述文字
            tk.Label(
                btn_container, text=item["desc"],
                font=("Microsoft YaHei", 8),
                fg="#95a5a6", bg="#2c3e50",
                anchor="w"
            ).pack(pady=(0, 6), fill="x")

        # 底部信息
        bottom_frame = tk.Frame(nav_frame, bg="#2c3e50")
        bottom_frame.pack(side="bottom", fill="x", pady=15, padx=10)

        tk.Label(
            bottom_frame, text="v2.0 导航版",
            font=("Microsoft YaHei", 9), fg="#7f8c8d", bg="#2c3e50"
        ).pack()

        # ===== 右侧内容区（白色背景）=====
        self.content_frame = tk.Frame(main_container, bg="white")
        self.content_frame.pack(side="right", fill="both", expand=True)

        # ===== 底部状态栏 =====
        status_bar = tk.Frame(self.root, bg="#ecf0f1", height=28)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)

        self.status_var = tk.StringVar(value="就绪 - 请选择左侧功能菜单")
        tk.Label(
            status_bar, textvariable=self.status_var,
            font=("Microsoft YaHei", 9),
            bg="#ecf0f1", fg="#7f8c8d", anchor="w", padx=15
        ).pack(fill="x")

        # 默认显示第一个功能
        self._switch_frame("sjpp")

    def _switch_frame(self, frame_key: str):
        """切换到指定的功能Frame"""
        # 更新导航按钮样式
        for key, btn in self.nav_buttons.items():
            if key == frame_key:
                btn.configure(bg="#3498db", fg="white", font=("Microsoft YaHei", 11, "bold"))
            else:
                btn.configure(bg="#34495e", fg="white", font=("Microsoft YaHei", 11))

        # 隐藏当前Frame
        if self.current_frame is not None:
            self.current_frame.pack_forget()

        # 获取或创建目标Frame实例
        if frame_key not in self.frame_instances:
            self.frame_instances[frame_key] = self._create_frame(frame_key)

        target_frame = self.frame_instances[frame_key]

        # 显示新Frame
        target_frame.pack(fill=tk.BOTH, expand=True)
        self.current_frame = target_frame

        # 更新状态栏
        status_texts = {
            "sjpp": "数据匹配 - 从原始表自动生成完整版表格",
            "wgl": "稳岗率计算 - 分析项目驻场和项目经理的稳岗率",
            "hbbg": "合并表格 - 将多个Excel文件合并到汇总表",
            "tenure": "在职月份计算 - 统计员工在职月份和天数"
        }
        self.status_var.set(status_texts.get(frame_key, ""))

    def _create_frame(self, frame_key: str) -> tk.Frame:
        """根据key创建对应的Frame实例"""
        if frame_key == "sjpp":
            from .sjpp_ui import DataMatchFrame
            return DataMatchFrame(self.content_frame, self.status_var)
        elif frame_key == "wgl":
            from .wgl_ui import StabilityFrame
            return StabilityFrame(self.content_frame, self.status_var)
        elif frame_key == "hbbg":
            from .hbbg_zhf import MergerFrame
            return MergerFrame(self.content_frame, self.status_var)
        elif frame_key == "tenure":
            from .tenure_gui import App as TenureApp
            container = tk.Frame(self.content_frame, bg="white")
            TenureApp(container, status_var=self.status_var)
            return container
        else:
            raise ValueError(f"未知的Frame类型: {frame_key}")


def main():
    """主函数"""
    root = tk.Tk()

    # 设置窗口图标（如果有的话）
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    # 创建主应用
    app = MainApp(root)

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
