#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
公共 UI 组件模块
提供统一的 Design Tokens 与可复用组件（FunctionCard）
"""

import customtkinter as ctk
import tkinter as tk


# ============================================================
# Design Tokens — 全局设计变量（唯一来源）
# ============================================================

class Theme:
    """设计令牌：所有颜色 / 字号 / 间距 / 圆角均在此定义"""

    # ---- 主题色 ----
    COLOR_NAV_BG = "#2c3e50"
    COLOR_NAV_BUTTON = "#34495e"
    COLOR_NAV_BUTTON_ACTIVE = "#3d566e"
    COLOR_NAV_DIVIDER = "#3d566e"

    COLOR_ACCENT_GENERAL = "#27ae60"
    COLOR_ACCENT_FINANCE = "#2980b9"
    COLOR_ACCENT_HR = "#e67e22"
    COLOR_ACCENT_PURPLE = "#8e44ad"
    COLOR_ACCENT_TEAL = "#16a085"

    COLOR_BG_CONTENT = "#ffffff"
    COLOR_BG_CARD = "#fafafa"
    COLOR_BG_CARD_HOVER = "#f0f0f0"
    COLOR_BORDER = "#e0e0e0"
    COLOR_STATUS_BAR = "#f0f2f5"

    # ---- 文字色 ----
    COLOR_TEXT_PRIMARY = "#2c3e50"
    COLOR_TEXT_SECONDARY = "#666666"
    COLOR_TEXT_CAPTION = "#888888"
    COLOR_TEXT_MUTED = "#95a5a6"
    COLOR_TEXT_ON_DARK = "#ffffff"

    # ---- 字号层级 ----
    # 调大：标题 20→24，子标题 14→17，正文 12→14，注释 11→13
    # 字体从 Microsoft YaHei 改为 Microsoft YaHei UI（Windows UI 优化版，渲染更清晰）
    FONT_DISPLAY = ("Microsoft YaHei UI", 26, "bold")
    FONT_HEADING = ("Microsoft YaHei UI", 24, "bold")
    FONT_SUBHEADING = ("Microsoft YaHei UI", 17, "bold")
    FONT_BODY = ("Microsoft YaHei UI", 14)
    FONT_BOLD = ("Microsoft YaHei UI", 14, "bold")
    FONT_CAPTION = ("Microsoft YaHei UI", 13)
    FONT_SMALL = ("Microsoft YaHei UI", 12)

    # ---- 间距 ----
    SPACING_XS = 6
    SPACING_SM = 10
    SPACING_MD = 20
    SPACING_LG = 28
    SPACING_XL = 36

    # ---- 圆角 ----
    RADIUS_SM = 4
    RADIUS_MD = 8
    RADIUS_LG = 12

    # ---- 尺寸 ----
    NAV_WIDTH = 250
    STATUS_BAR_HEIGHT = 44
    CARD_GRID_BREAKPOINT = 600
    CARD_BUTTON_HEIGHT = 40


# ============================================================
# 辅助函数
# ============================================================

def darken_color(hex_color: str, factor: float = 0.2) -> str:
    """将 hex 颜色暗化指定比例（0~1）"""
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = max(0, int(r * (1 - factor)))
    g = max(0, int(g * (1 - factor)))
    b = max(0, int(b * (1 - factor)))
    return f"#{r:02x}{g:02x}{b:02x}"


# ============================================================
# FunctionCard 组件
# ============================================================

class FunctionCard(ctk.CTkFrame):
    """
    通用功能卡片
    ┌────┬─────────────────────────────┬────────┐
    │ 色 ▎ 📌 功能标题                  │ ▶ 运行 │
    │ 条 ▎ 功能描述文字...               │        │
    └────┴─────────────────────────────┴────────┘
    """

    def __init__(
        self,
        parent,
        title: str,
        description: str,
        icon: str = "📌",
        color: str = Theme.COLOR_ACCENT_GENERAL,
        command=None,
        **kwargs
    ):
        super().__init__(
            parent,
            fg_color=Theme.COLOR_BG_CARD,
            corner_radius=Theme.RADIUS_LG,
            border_width=1,
            border_color=Theme.COLOR_BORDER,
            **kwargs
        )

        self._color = color

        # ---- 左侧 4px 色条 ----
        self._accent_bar = tk.Frame(self, bg=color, width=4, highlightthickness=0)
        self._accent_bar.pack(side="left", fill="y")
        self._accent_bar.pack_propagate(False)

        # ---- 主内容区 ----
        self._inner = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self._inner.pack(side="left", fill="both", expand=True,
                         padx=Theme.SPACING_MD, pady=Theme.SPACING_MD + 2)

        # 左侧信息
        self._left = ctk.CTkFrame(self._inner, fg_color="transparent", corner_radius=0)
        self._left.pack(side="left", fill="both", expand=True)

        self._title_label = ctk.CTkLabel(
            self._left,
            text=f"{icon}  {title}",
            font=Theme.FONT_SUBHEADING,
            text_color=Theme.COLOR_TEXT_PRIMARY,
            anchor="w",
            justify="left"
        )
        self._title_label.pack(anchor="w", fill="x")

        self._desc_label = ctk.CTkLabel(
            self._left,
            text=description,
            font=Theme.FONT_CAPTION,
            text_color=Theme.COLOR_TEXT_CAPTION,
            anchor="w",
            justify="left"
        )
        self._desc_label.pack(anchor="w", fill="x", pady=(Theme.SPACING_XS, 0))

        # 右侧按钮
        self._btn = ctk.CTkButton(
            self._inner,
            text="▶ 运行",
            font=Theme.FONT_BOLD,
            fg_color=color,
            hover_color=darken_color(color, 0.25),
            text_color=Theme.COLOR_TEXT_ON_DARK,
            corner_radius=Theme.RADIUS_MD,
            width=110,
            height=Theme.CARD_BUTTON_HEIGHT,
            command=self._on_click
        )
        self._btn.pack(side="right", padx=(Theme.SPACING_LG, 0))

        self._command = command

        # ---- Hover 效果 ----
        self._bind_hover()

    def _bind_hover(self):
        """绑定鼠标悬停事件"""
        targets = [
            self, self._inner, self._left,
            self._title_label, self._desc_label, self._accent_bar
        ]
        for t in targets:
            t.bind("<Enter>", self._on_enter, add="+")
            t.bind("<Leave>", self._on_leave, add="+")
            try:
                t.configure(cursor="hand2")
            except Exception:
                pass

    def _on_enter(self, event):
        self.configure(fg_color=Theme.COLOR_BG_CARD_HOVER)

    def _on_leave(self, event):
        self.configure(fg_color=Theme.COLOR_BG_CARD)

    def _on_click(self):
        if self._command:
            self._command()
