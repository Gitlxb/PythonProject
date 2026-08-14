#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
浙江锦途 - 处理Excel的Python脚本
使用CustomTkinter实现的现代化GUI
左侧导航栏 + Frame切换 方案
整合：通用功能、财务功能、人事功能
"""

# ============================================================
# Step 0: Windows 高 DPI 感知 —— 一切之前，确保字体清晰
# ============================================================
import ctypes
import sys as _sys
if _sys.platform == "win32":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor v2
    except (AttributeError, OSError):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except (AttributeError, OSError):
                pass

# ============================================================
# Step 1: 启动画面 —— 先亮 Splash，再加载重型模块
# ============================================================
import tkinter as tk

_SPLASH_BG = "#2c3e50"
_SPLASH_FG = "#ffffff"
_SPLASH_MUTED = "#7f8c8d"
_SPLASH_W, _SPLASH_H = 500, 340
_BAR_W, _BAR_H = 300, 4

_splash = tk.Tk()
_splash.overrideredirect(True)       # 无边框
_splash.attributes("-topmost", True)  # 置顶

# ---- 尺寸 & 居中 ----
screen_w = _splash.winfo_screenwidth()
screen_h = _splash.winfo_screenheight()
_splash.geometry(
    f"{_SPLASH_W}x{_SPLASH_H}+{(screen_w - _SPLASH_W) // 2}+{(screen_h - _SPLASH_H) // 2}"
)
_splash.configure(bg=_SPLASH_BG)

# ---- 品牌名 ----
tk.Label(
    _splash,
    text="浙江锦途",
    font=("Microsoft YaHei UI", 30, "bold"),
    fg=_SPLASH_FG, bg=_SPLASH_BG
).pack(pady=(56, 4))

# ---- 副标题 ----
tk.Label(
    _splash,
    text="Excel 数据处理工具集",
    font=("Microsoft YaHei UI", 14),
    fg=_SPLASH_MUTED, bg=_SPLASH_BG
).pack()

# ---- 分割线 ----
tk.Frame(_splash, height=1, bg="#3d566e").pack(
    fill="x", padx=80, pady=(20, 14)
)

# ---- 加载提示 ----
_load_var = tk.StringVar(value="正在启动，请稍候 ...")
tk.Label(
    _splash,
    textvariable=_load_var,
    font=("Microsoft YaHei UI", 12),
    fg="#95a5a6", bg=_SPLASH_BG
).pack()

# ---- 版本号 ----
tk.Label(
    _splash,
    text="v4.0",
    font=("Microsoft YaHei UI", 11),
    fg="#555555", bg=_SPLASH_BG
).pack(pady=(14, 0))

# ---- 进度条（Canvas 自绘，背景槽 + 填充矩形）----
_progress_var = tk.DoubleVar(value=0)
_percent_var = tk.StringVar(value="0%")

_progress_canvas = tk.Canvas(
    _splash,
    width=_BAR_W, height=_BAR_H,
    bg=_SPLASH_BG, highlightthickness=0, bd=0
)
_progress_canvas.pack(pady=(22, 4))

# 背景槽（深色）
_progress_canvas.create_rectangle(
    0, 0, _BAR_W, _BAR_H,
    fill="#1a2530", outline=""
)
# 进度填充（绿色，宽度由 _set_progress 动态控制）
_fill_id = _progress_canvas.create_rectangle(
    0, 0, 0, _BAR_H,
    fill="#27ae60", outline=""
)

# 百分比文字（居中，紧贴进度条下方）
tk.Label(
    _splash,
    textvariable=_percent_var,
    font=("Microsoft YaHei UI", 10),
    fg=_SPLASH_MUTED, bg=_SPLASH_BG
).pack()

# ---- 进度更新工具函数 ----
def _set_progress(pct: float, text: str = None):
    """更新进度条和提示文字（0-100）"""
    pct = max(0.0, min(100.0, pct))
    _progress_var.set(pct)
    _progress_canvas.coords(_fill_id, 0, 0, _BAR_W * pct / 100, _BAR_H)
    _percent_var.set(f"{int(pct)}%")
    if text is not None:
        _load_var.set(text)
    _splash.update_idletasks()  # 不阻塞主线程，刷新 idle 任务

# 初始 0%
_set_progress(0, "正在启动，请稍候 ...")
_splash.update()

# ============================================================
# Step 2: 重型导入（Splash 覆盖此阶段，按阶段上报进度）
# ============================================================

_set_progress(15, "正在加载核心模块 ...")
import customtkinter as ctk

_set_progress(45, "正在加载功能模块 ...")
from frame_general import GeneralFrame
from frame_finance import FinanceFrame
from frame_hr import HRFrame
from ui_components import Theme

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

_set_progress(70, "正在准备界面 ...")


# ============================================================
# Step 3: 主界面
# ============================================================

class MainApp:
    """浙江锦途主界面 - 左侧导航+Frame切换"""

    def __init__(self, root):
        self.root = root
        self.root.title("浙江锦途 - 处理Excel的Python脚本")
        self.root.geometry("1180x800")
        self.root.minsize(1000, 680)

        # 当前激活的Frame
        self.current_frame = None
        self.frame_instances = {}

        # 导航引用
        self.nav_buttons = {}   # key -> CTkButton
        self.nav_strips = {}    # key -> tk.Frame (左侧色条)

        # 模块颜色映射
        self.module_colors = {
            "general": Theme.COLOR_ACCENT_GENERAL,
            "finance": Theme.COLOR_ACCENT_FINANCE,
            "hr": Theme.COLOR_ACCENT_HR,
        }

        # 创建界面
        self._create_ui()

    # ============================================================
    # 界面构建
    # ============================================================

    def _create_ui(self):
        """创建界面布局"""
        # 主水平容器
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        # ===== 左侧导航栏 =====
        self._build_nav(main_container)

        # ===== 右侧内容区（CTkScrollableFrame）=====
        self._build_content(main_container)

        # ===== 底部状态栏 =====
        self._build_status_bar()

        # 默认显示第一个功能
        self._switch_frame("general")

    def _build_nav(self, parent):
        """构建左侧导航栏"""
        nav_frame = ctk.CTkFrame(
            parent,
            width=Theme.NAV_WIDTH,
            fg_color=Theme.COLOR_NAV_BG,
            corner_radius=0
        )
        nav_frame.pack(side="left", fill="y")
        nav_frame.pack_propagate(False)

        # ---- 标题 ----
        ctk.CTkLabel(
            nav_frame,
            text="🔧 功能菜单",
            font=Theme.FONT_HEADING,
            text_color=Theme.COLOR_TEXT_ON_DARK,
        ).pack(pady=(Theme.SPACING_LG, Theme.SPACING_SM))

        # ---- 分隔线 ----
        ctk.CTkFrame(
            nav_frame, height=2, fg_color=Theme.COLOR_NAV_DIVIDER
        ).pack(fill="x", padx=Theme.SPACING_MD, pady=(0, Theme.SPACING_MD))

        # ---- 导航项 ----
        nav_items = [
            ("general", "📂 通用工具", "Excel拆分 & 文件获取",     Theme.COLOR_ACCENT_GENERAL),
            ("finance", "💰 财务工具", "记账/合并/手续费/考核等",  Theme.COLOR_ACCENT_FINANCE),
            ("hr",      "👥 人事工具", "在职离职/工资匹配/文档",   Theme.COLOR_ACCENT_HR),
        ]

        for key, text, desc, color in nav_items:
            # 分类标签
            ctk.CTkLabel(
                nav_frame,
                text=desc,
                font=Theme.FONT_CAPTION,
                text_color=Theme.COLOR_TEXT_MUTED
            ).pack(padx=Theme.SPACING_LG, pady=(Theme.SPACING_LG, Theme.SPACING_XS), anchor="w")

            # 按钮容器（包裹色条+按钮）
            btn_wrapper = ctk.CTkFrame(
                nav_frame, fg_color="transparent",
                corner_radius=0, height=40
            )
            btn_wrapper.pack(pady=Theme.SPACING_XS, padx=Theme.SPACING_MD, fill="x")
            btn_wrapper.pack_propagate(False)

            # 左侧色条（默认与背景同色=隐藏，激活时显示模块色）
            strip = tk.Frame(
                btn_wrapper,
                bg=Theme.COLOR_NAV_BG,
                width=3, highlightthickness=0
            )
            strip.pack(side="left", fill="y")
            strip.pack_propagate(False)
            self.nav_strips[key] = strip

            # 按钮
            btn = ctk.CTkButton(
                btn_wrapper,
                text=text,
                command=lambda k=key: self._switch_frame(k),
                font=Theme.FONT_BODY,
                fg_color=Theme.COLOR_NAV_BUTTON,
                hover_color=Theme.COLOR_NAV_BUTTON_ACTIVE,
                text_color=Theme.COLOR_TEXT_ON_DARK,
                corner_radius=Theme.RADIUS_MD,
                height=40,
                anchor="w"
            )
            btn.pack(side="left", fill="both", expand=True)
            self.nav_buttons[key] = btn

        # ---- 底部分隔线（固定底部）----
        ctk.CTkFrame(
            nav_frame, height=2, fg_color=Theme.COLOR_NAV_DIVIDER
        ).pack(fill="x", padx=Theme.SPACING_MD,
               pady=(Theme.SPACING_XL, Theme.SPACING_MD), side="bottom")

        # ---- 底部信息 ----
        bottom = ctk.CTkFrame(nav_frame, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", pady=Theme.SPACING_MD,
                    padx=Theme.SPACING_MD)

        ctk.CTkLabel(
            bottom,
            text="📢 需要添加需求\n   请在企业微信搜索: 龙喜兵",
            font=Theme.FONT_CAPTION,
            text_color=Theme.COLOR_TEXT_MUTED,
            justify="left"
        ).pack(anchor="w")

        ctk.CTkLabel(
            bottom,
            text="\nv4.0 CustomTkinter版",
            font=Theme.FONT_BODY,
            text_color=Theme.COLOR_TEXT_SECONDARY
        ).pack(anchor="w")

    def _build_content(self, parent):
        """构建右侧可滚动内容区"""
        content_outer = ctk.CTkFrame(
            parent,
            fg_color=Theme.COLOR_BG_CONTENT,
            corner_radius=0
        )
        content_outer.pack(side="right", fill="both", expand=True)

        self.content_frame = ctk.CTkScrollableFrame(
            content_outer,
            fg_color=Theme.COLOR_BG_CONTENT,
            corner_radius=0,
            scrollbar_button_color="#cccccc",
            scrollbar_button_hover_color="#aaaaaa"
        )
        self.content_frame.pack(fill="both", expand=True)

    def _build_status_bar(self):
        """构建底部状态栏"""
        status_bar = ctk.CTkFrame(
            self.root,
            height=Theme.STATUS_BAR_HEIGHT,
            fg_color=Theme.COLOR_STATUS_BAR,
            corner_radius=0
        )
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)

        # 左侧状态指示灯
        self.status_dot = tk.Canvas(
            status_bar,
            width=14, height=Theme.STATUS_BAR_HEIGHT,
            bg=Theme.COLOR_STATUS_BAR,
            highlightthickness=0
        )
        self.status_dot.pack(side="left", padx=(Theme.SPACING_LG, Theme.SPACING_SM))
        self._dot_id = self.status_dot.create_oval(
            4, 13, 12, 21, fill=Theme.COLOR_ACCENT_GENERAL, outline=""
        )

        # 状态文本
        self.status_var = tk.StringVar(value="就绪 - 请选择左侧功能菜单")
        ctk.CTkLabel(
            status_bar,
            textvariable=self.status_var,
            font=Theme.FONT_SMALL,
            text_color=Theme.COLOR_TEXT_SECONDARY,
            anchor="w"
        ).pack(side="left", fill="x", expand=True, padx=(0, Theme.SPACING_LG))

    # ============================================================
    # 导航切换
    # ============================================================

    def _switch_frame(self, frame_key: str):
        """切换到指定的功能Frame"""
        color = self.module_colors.get(frame_key, Theme.COLOR_ACCENT_GENERAL)

        for key, btn in self.nav_buttons.items():
            strip = self.nav_strips.get(key)
            if key == frame_key:
                btn.configure(fg_color=Theme.COLOR_NAV_BUTTON_ACTIVE)
                if strip:
                    strip.configure(bg=color)
            else:
                btn.configure(fg_color=Theme.COLOR_NAV_BUTTON)
                if strip:
                    strip.configure(bg=Theme.COLOR_NAV_BG)

        self.status_dot.itemconfig(self._dot_id, fill=color)

        if self.current_frame is not None:
            self.current_frame.pack_forget()

        if frame_key not in self.frame_instances:
            self.frame_instances[frame_key] = self._create_frame(frame_key)

        target_frame = self.frame_instances[frame_key]
        target_frame.pack(fill=tk.BOTH, expand=True)
        self.current_frame = target_frame

        status_texts = {
            "general": "通用工具 - Excel文件拆分与文件获取",
            "finance": "财务工具 - 收支记账、合并表格、绩效考核等",
            "hr": "人事工具 - 在职离职拆分、参数拆分",
        }
        self.status_var.set(status_texts.get(frame_key, ""))

    def _create_frame(self, frame_key: str) -> ctk.CTkFrame:
        if frame_key == "general":
            return GeneralFrame(self.content_frame, self.status_var)
        elif frame_key == "finance":
            return FinanceFrame(self.content_frame, self.status_var)
        elif frame_key == "hr":
            return HRFrame(self.content_frame, self.status_var)
        else:
            raise ValueError(f"未知的Frame类型: {frame_key}")


# ============================================================
# Step 4: 入口 —— 关闭 Splash → 打开主窗口
# ============================================================

def main():
    _set_progress(85, "正在构建主界面 ...")
    root = ctk.CTk()
    app = MainApp(root)

    # ============================================================
    # 关键修复：销毁 splash 前，把主窗口设为 tkinter 的默认 root
    #
    # 原因：splash 用独立的 tk.Tk()，主窗口 ctk.CTk() 是另一个 Tk 实例。
    # _splash.destroy() 时其 Tcl 解释器被销毁，tkinter 会把
    # _default_root 重置为 None。后续若调用 tk.StringVar()（不带 master）
    # 会报 "Too early to create variable: no default root window"。
    # 显式赋值后，主窗口成为唯一的 default root。
    # ============================================================
    tk._default_root = root

    # 进度到 100%，并短暂停留让用户看到"启动完成"
    _set_progress(100, "启动完成")
    _splash.update()
    import time as _time
    _time.sleep(0.4)

    # 关闭 splash，聚焦主窗口
    _splash.destroy()
    root.lift()
    root.focus_force()

    root.mainloop()


if __name__ == "__main__":
    main()
