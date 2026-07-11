# -- coding: utf-8 --
# @File : main_integrated.py
# @Description: 四合一整合工具 - 代工费拆分 / 数据比较 / 预支处理 / 报销拆分
# @Note: 各功能核心逻辑文件不做任何修改，仅适配GUI层为Frame嵌入模式

import tkinter as tk
from tkinter import ttk, messagebox


class IntegratedApp:
    """四合一工具主应用"""

    NAV_BUTTONS = [
        {"name": "dgf_cf", "text": "📊 代工费拆分", "desc": "处理代工费表格与发放记录的自动拆分"},
        {"name": "bijiao", "text": "🔍 数据比较", "desc": "工人预支申请表 vs 预支明细表 对比"},
        {"name": "yuzhi", "text": "💰 预支处理", "desc": "手续费计算、拆分工作表并格式化"},
        {"name": "bxcf", "text": "📋 报销拆分", "desc": "Excel合并单元格拆分、按列拆分统计"},
    ]

    def __init__(self, parent=None):
        if parent and (isinstance(parent, tk.Tk) or isinstance(parent, tk.Toplevel)):
            self.root = tk.Toplevel(parent)
            self.root.transient(parent)
            self.root.grab_set()
        else:
            self.root = tk.Tk()
        self.root.title("浙江锦途 - 财务工具")
        self.root.geometry("1150x780")
        self.root.minsize(1000, 650)
        self.root.configure(bg="#F5F5F5")

        # 当前显示的Frame引用
        self.current_frame = None
        self.frames = {}

        self._build_ui()
        self._switch_frame("dgf_cf")  # 默认显示第一个

        if parent:
            self.root.lift()
            self.root.focus_force()

    def _build_ui(self):
        # ===== 左侧导航栏 =====
        self.nav_frame = tk.Frame(self.root, bg="#2c3e50", width=200)
        self.nav_frame.pack(side="left", fill="y")
        self.nav_frame.pack_propagate(False)

        # 标题
        tk.Label(
            self.nav_frame, text="财务工具箱",
            font=("Microsoft YaHei", 16, "bold"), fg="white",
            bg="#2c3e50", pady=20
        ).pack()

        # 分隔线
        ttk.Separator(self.nav_frame, orient="horizontal").pack(fill="x", padx=10)

        # 导航按钮容器
        btn_container = tk.Frame(self.nav_frame, bg="#2c3e50")
        btn_container.pack(fill="x", padx=10, pady=15)

        self.nav_btn_refs = {}
        for item in self.NAV_BUTTONS:
            btn = tk.Button(
                btn_container,
                text=item["text"],
                font=("Microsoft YaHei", 11),
                fg="white", bg="#34495e",
                activebackground="#3498db", activeforeground="white",
                relief="flat", bd=0,
                anchor="w", width=18, height=2,
                cursor="hand2",
                command=lambda name=item["name"]: self._switch_frame(name)
            )
            btn.pack(pady=4, fill="x")
            self.nav_btn_refs[item["name"]] = btn

        # 底部信息
        bottom_frame = tk.Frame(self.nav_frame, bg="#2c3e50")
        bottom_frame.pack(side="bottom", fill="x", pady=15, padx=10)

        tk.Label(
            bottom_frame, text="v1.0 | 四合一版本",
            font=("Microsoft YaHei", 9), fg="#95a5a6",
            bg="#2c3e50"
        ).pack()

        # ===== 右侧内容区 =====
        self.content_area = tk.Frame(self.root, bg="white")
        self.content_area.pack(side="right", fill="both", expand=True)

        # ===== 状态栏 =====
        status_bar = tk.Frame(self.root, bg="#ecf0f1", height=28)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)

        self.status_label = tk.Label(
            status_bar, text="就绪", font=("Microsoft YaHei", 9),
            bg="#ecf0f1", fg="#7f8c8d", anchor="w", padx=15
        )
        self.status_label.pack(fill="x")

    def _switch_frame(self, frame_name):
        """切换右侧内容区显示的功能Frame"""
        # 更新导航按钮样式
        for name, btn in self.nav_btn_refs.items():
            if name == frame_name:
                btn.configure(bg="#3498db")
            else:
                btn.configure(bg="#34495e")

        # 销毁当前Frame
        if self.current_frame is not None:
            self.current_frame.destroy()
            self.current_frame = None

        # 延迟加载/创建新Frame
        try:
            frame_cls = self._get_frame_class(frame_name)
            self.current_frame = frame_cls(self.content_area, self)
            self.current_frame.pack(fill="both", expand=True)
            desc = next((item["desc"] for item in self.NAV_BUTTONS if item["name"] == frame_name), "")
            self.status_label.config(text=f"当前: {desc}")
        except Exception as e:
            messagebox.showerror("加载错误", f"加载 [{frame_name}] 失败:\n{str(e)}")

    def _get_frame_class(self, frame_name):
        """根据名称返回对应的Frame类"""
        if frame_name == "dgf_cf":
            from .dgf_cf_frame import DgfCfFrame
            return DgfCfFrame
        elif frame_name == "bijiao":
            from .bijiao_ui import BijiaoFrame
            return BijiaoFrame
        elif frame_name == "yuzhi":
            from .yuzhi_frame import YuzhiFrame
            return YuzhiFrame
        elif frame_name == "bxcf":
            from .bxcf_frame import BxcfFrame
            return BxcfFrame
        else:
            raise ValueError(f"未知功能: {frame_name}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = IntegratedApp()
    app.run()
