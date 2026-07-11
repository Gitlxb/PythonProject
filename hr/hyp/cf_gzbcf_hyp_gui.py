"""
工资表拆分工具 - 统一GUI界面
包含增强版和简化版两个标签页，可在同一窗口内切换使用

使用方式：
  # 独立窗口
  app = UnifiedSplitApp()
  app.run()

  # 嵌入到 Frame（统一工具箱等场景）
  app = UnifiedSplitApp(parent=frame)
"""
import tkinter as tk
from tkinter import ttk

from .cf_gzbcf_hyp import SplitExcelApp as EnhancedApp
from .cf_gzbcf_hyp_simple import SplitExcelApp as SimpleApp


class UnifiedSplitApp:
    """
    统一拆分工具界面，支持独立窗口和 Frame 嵌入两种模式。
    
    参数:
        parent: None → 创建独立 Tk 窗口
                tk.Widget → UI 构建在该 widget 内（嵌入模式）
    """

    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
            self._standalone = True
        else:
            # 嵌入模式：直接使用传入的 parent 作为容器
            self.root = parent
            self._standalone = False

        self._build_ui()

        if self._standalone:
            self.root.title("工资表拆分工具")
            self.root.geometry("650x620")
            self.root.resizable(False, False)

    def _build_ui(self):
        """构建界面"""
        container = self.root  # 独立模式= Tk，嵌入模式= 传入的 Frame

        # 标题（仅在独立模式显示）
        if self._standalone:
            title_frame = tk.Frame(container, bg="#f0f0f0")
            title_frame.pack(fill="x", padx=0, pady=0)
            title_label = tk.Label(
                title_frame,
                text="🧾 工资表拆分工具",
                font=("Microsoft YaHei", 16, "bold"),
                bg="#f0f0f0", fg="#333333"
            )
            title_label.pack(anchor="w", padx=15, pady=10)

        # ---- 切换按钮栏 ----
        btn_bar = tk.Frame(container, bg="#f0f0f0")
        btn_bar.pack(fill="x", padx=10, pady=(0, 5))

        self._btn_enhanced = tk.Button(
            btn_bar, text=" 增强版 ",
            font=("Microsoft YaHei", 11, "bold"),
            bg="#27ae60", fg="white",
            relief="flat", cursor="hand2",
            command=self._show_enhanced
        )
        self._btn_enhanced.pack(side="left", padx=(0, 4))

        self._btn_simple = tk.Button(
            btn_bar, text=" 简化版 ",
            font=("Microsoft YaHei", 11, "bold"),
            bg="#bdc3c7", fg="#333333",
            relief="flat", cursor="hand2",
            command=self._show_simple
        )
        self._btn_simple.pack(side="left", padx=4)

        # 当前使用提示标签
        self._hint_label = tk.Label(
            btn_bar,
            text="当前使用：增强版",
            font=("Microsoft YaHei", 10),
            bg="#f0f0f0", fg="#27ae60"
        )
        self._hint_label.pack(side="left", padx=(15, 0))

        # ---- 内容区 Frame ----
        self._content_frame = tk.Frame(container)
        self._content_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))

        # 增强版 Frame
        self.frame_enhanced = tk.Frame(self._content_frame)
        self.enhanced_app = EnhancedApp(self.frame_enhanced)
        self.frame_enhanced.pack(fill="both", expand=True)

        # 简化版 Frame（初始隐藏）
        self.frame_simple = tk.Frame(self._content_frame)
        self.simple_app = SimpleApp(self.frame_simple)

        # 底部说明（仅在独立模式显示）
        if self._standalone:
            info_frame = tk.Frame(container, bg="#f9f9f9")
            info_frame.pack(fill="x", padx=10, pady=(0, 10))

            info_text = ("💡 增强版：支持社保/公积金公式回填、详细日志\n"
                         "💡 简化版：轻量快速，基础拆分功能")
            info_label = tk.Label(
                info_frame,
                text=info_text,
                font=("Microsoft YaHei", 9),
                bg="#f9f9f9",
                fg="#666666",
                justify="left"
            )
            info_label.pack(anchor="w", padx=10, pady=5)

        self._current = "enhanced"  # 标记当前显示的版本

    def run(self):
        """独立模式入口"""
        if self._standalone:
            self.root.mainloop()

    def _show_enhanced(self):
        """切换到增强版"""
        if self._current == "enhanced":
            return
        self.frame_simple.pack_forget()
        self.frame_enhanced.pack(fill="both", expand=True)
        self._btn_enhanced.config(bg="#27ae60", fg="white")
        self._btn_simple.config(bg="#bdc3c7", fg="#333333")
        self._hint_label.config(text="当前使用：增强版", fg="#27ae60")
        self._current = "enhanced"

    def _show_simple(self):
        """切换到简化版"""
        if self._current == "simple":
            return
        self.frame_enhanced.pack_forget()
        self.frame_simple.pack(fill="both", expand=True)
        self._btn_simple.config(bg="#2980b9", fg="white")
        self._btn_enhanced.config(bg="#bdc3c7", fg="#333333")
        self._hint_label.config(text="当前使用：简化版", fg="#2980b9")
        self._current = "simple"


def main():
    """独立运行入口"""
    app = UnifiedSplitApp()
    app.run()


if __name__ == "__main__":
    main()
