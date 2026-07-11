#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
浙江锦途 - 处理Excel的Python脚本
使用CustomTkinter实现的现代化GUI
左侧导航栏 + Frame切换 方案
整合：通用功能、财务功能、人事功能
"""

import customtkinter as ctk
import tkinter as tk
from frame_general import GeneralFrame
from frame_finance import FinanceFrame
from frame_hr import HRFrame

# 设置CustomTkinter外观
ctk.set_appearance_mode("light")  # 可选: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # 可选: "blue", "green", "dark-blue"


class MainApp:
    """浙江锦途主界面 - 左侧导航+Frame切换"""

    def __init__(self, root):
        self.root = root
        self.root.title("浙江锦途 - 处理Excel的Python脚本")
        self.root.geometry("1100x750")
        self.root.minsize(950, 650)

        # 当前激活的Frame
        self.current_frame = None
        self.frame_instances = {}

        # 导航按钮引用
        self.nav_buttons = {}

        # 创建界面
        self._create_ui()

    def _create_ui(self):
        """创建界面布局"""
        # 主水平容器
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        # ===== 左侧导航栏（深色背景）=====
        nav_frame = ctk.CTkFrame(main_container, width=240, fg_color="#2c3e50", corner_radius=0)
        nav_frame.pack(side="left", fill="y")
        nav_frame.pack_propagate(False)

        # 导航标题
        title_label = ctk.CTkLabel(
            nav_frame,
            text="🔧 功能菜单",
            font=("Microsoft YaHei", 20, "bold"),
            text_color="white",
            pady=20
        )
        title_label.pack(pady=(20, 10))

        # 分隔线
        separator = ctk.CTkFrame(nav_frame, height=2, fg_color="#34495e")
        separator.pack(fill="x", padx=15, pady=(0, 15))

        # ===== 分类：通用功能 =====
        category_label = ctk.CTkLabel(
            nav_frame,
            text="— 通用功能 —",
            font=("Microsoft YaHei", 12),
            text_color="#7f8c8d"
        )
        category_label.pack(padx=15, pady=(10, 5), anchor="w")

        general_items = [
            {"key": "general", "text": "📂 通用工具", "desc": "Excel拆分 & 文件获取"},
        ]
        for item in general_items:
            btn = ctk.CTkButton(
                nav_frame,
                text=item["text"],
                command=lambda k=item["key"]: self._switch_frame(k),
                font=("Microsoft YaHei", 14),
                fg_color="#34495e",
                hover_color="#27ae60",
                text_color="white",
                corner_radius=8,
                height=45,
                anchor="w"
            )
            btn.pack(pady=5, padx=15, fill="x")
            self.nav_buttons[item["key"]] = btn

            desc_label = ctk.CTkLabel(
                nav_frame,
                text=item["desc"],
                font=("Microsoft YaHei", 11),
                text_color="#7f8c8d"
            )
            desc_label.pack(padx=(25, 0), pady=(0, 8), anchor="w")

        # ===== 分类：财务功能 =====
        category_label2 = ctk.CTkLabel(
            nav_frame,
            text="— 财务功能 —",
            font=("Microsoft YaHei", 12),
            text_color="#7f8c8d"
        )
        category_label2.pack(padx=15, pady=(20, 5), anchor="w")

        finance_items = [
            {"key": "finance", "text": "💰 财务工具", "desc": "记账/合并/手续费/考核等"},
        ]
        for item in finance_items:
            btn = ctk.CTkButton(
                nav_frame,
                text=item["text"],
                command=lambda k=item["key"]: self._switch_frame(k),
                font=("Microsoft YaHei", 14),
                fg_color="#34495e",
                hover_color="#2980b9",
                text_color="white",
                corner_radius=8,
                height=45,
                anchor="w"
            )
            btn.pack(pady=5, padx=15, fill="x")
            self.nav_buttons[item["key"]] = btn

            desc_label2 = ctk.CTkLabel(
                nav_frame,
                text=item["desc"],
                font=("Microsoft YaHei", 11),
                text_color="#7f8c8d"
            )
            desc_label2.pack(padx=(25, 0), pady=(0, 8), anchor="w")

        # ===== 分类：人事功能 =====
        category_label3 = ctk.CTkLabel(
            nav_frame,
            text="— 人事功能 —",
            font=("Microsoft YaHei", 12),
            text_color="#7f8c8d"
        )
        category_label3.pack(padx=15, pady=(20, 5), anchor="w")

        hr_items = [
            {"key": "hr", "text": "👥 人事工具", "desc": "在职离职拆分/工资表匹配/Word合并"},
        ]
        for item in hr_items:
            btn = ctk.CTkButton(
                nav_frame,
                text=item["text"],
                command=lambda k=item["key"]: self._switch_frame(k),
                font=("Microsoft YaHei", 14),
                fg_color="#34495e",
                hover_color="#e67e22",
                text_color="white",
                corner_radius=8,
                height=45,
                anchor="w"
            )
            btn.pack(pady=5, padx=15, fill="x")
            self.nav_buttons[item["key"]] = btn

            desc_label3 = ctk.CTkLabel(
                nav_frame,
                text=item["desc"],
                font=("Microsoft YaHei", 11),
                text_color="#7f8c8d"
            )
            desc_label3.pack(padx=(25, 0), pady=(0, 8), anchor="w")

        # 底部公告区域
        separator2 = ctk.CTkFrame(nav_frame, height=2, fg_color="#34495e")
        separator2.pack(fill="x", padx=15, pady=(30, 15), side="bottom")

        bottom_frame = ctk.CTkFrame(nav_frame, fg_color="transparent")
        bottom_frame.pack(side="bottom", fill="x", pady=15, padx=15)

        info_label = ctk.CTkLabel(
            bottom_frame,
            text="📢 需要添加需求\n   请在企业微信搜索: 龙喜兵",
            font=("Microsoft YaHei", 11),
            text_color="#7f8c8d",
            justify="left"
        )
        info_label.pack(anchor="w")

        version_label = ctk.CTkLabel(
            bottom_frame,
            text="\nv4.0 CustomTkinter版",
            font=("Microsoft YaHei", 12),
            text_color="#555555"
        )
        version_label.pack(anchor="w")

        # ===== 右侧内容区（白色背景 + 可滚动）=====
        content_outer = ctk.CTkFrame(main_container, fg_color="white", corner_radius=0)
        content_outer.pack(side="right", fill="both", expand=True)

        # 垂直滚动条
        self.content_scrollbar = ctk.CTkScrollbar(content_outer, orientation="vertical")
        self.content_scrollbar.pack(side="right", fill="y")

        # 画布作为可滚动容器
        self.canvas = tk.Canvas(
            content_outer,
            bg="white",
            yscrollcommand=self.content_scrollbar.set,
            highlightthickness=0
        )
        self.canvas.pack(side="left", fill="both", expand=True)
        self.content_scrollbar.configure(command=self.canvas.yview)

        # 内容frame（放入canvas中）
        self.content_frame = tk.Frame(self.canvas, bg="white")
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.content_frame, anchor="nw"
        )

        # 绑定鼠标滚轮事件（支持Windows）
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")

        self.canvas.bind("<Enter>", _bind_mousewheel)
        self.canvas.bind("<Leave>", _unbind_mousewheel)

        # 当内容高度变化时更新滚动区域
        self.content_frame.bind("<Configure>", self._on_content_configure)

        # 当Canvas宽度变化时，同步调整内部frame宽度（自适应布局）
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # ===== 底部状态栏 =====
        status_bar = ctk.CTkFrame(self.root, height=32, fg_color="#ecf0f1", corner_radius=0)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)

        self.status_var = tk.StringVar(value="就绪 - 请选择左侧功能菜单")
        status_label = ctk.CTkLabel(
            status_bar,
            textvariable=self.status_var,
            font=("Microsoft YaHei", 10),
            text_color="#7f8c8d",
            anchor="w"
        )
        status_label.pack(fill="x", padx=20)

        # 默认显示第一个功能
        self._switch_frame("general")

    def _switch_frame(self, frame_key: str):
        """切换到指定的功能Frame"""
        # 更新导航按钮样式
        color_map = {
            "general": "#27ae60",
            "finance": "#2980b9",
            "hr": "#e67e22",
        }
        highlight_color = color_map.get(frame_key, "#3498db")

        for key, btn in self.nav_buttons.items():
            if key == frame_key:
                btn.configure(fg_color=highlight_color)
            else:
                btn.configure(fg_color="#34495e")

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
            "general": "通用工具 - Excel文件拆分与文件获取",
            "finance": "财务工具 - 收支记账、合并表格、绩效考核等",
            "hr": "人事工具 - 在职离职拆分、参数拆分",
        }
        self.status_var.set(status_texts.get(frame_key, ""))

    def _create_frame(self, frame_key: str) -> tk.Frame:
        """根据key创建对应的Frame实例"""
        if frame_key == "general":
            return GeneralFrame(self.content_frame, self.status_var)
        elif frame_key == "finance":
            return FinanceFrame(self.content_frame, self.status_var)
        elif frame_key == "hr":
            return HRFrame(self.content_frame, self.status_var)
        else:
            raise ValueError(f"未知的Frame类型: {frame_key}")

    def _on_content_configure(self, event):
        """内容区域高度变化时更新Canvas滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """Canvas宽度变化时，同步拉伸内部content_frame宽度"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)


def main():
    """主函数"""
    root = ctk.CTk()
    app = MainApp(root)

    root.mainloop()


if __name__ == "__main__":
    main()
