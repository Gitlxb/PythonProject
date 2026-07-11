# -*- coding: utf-8 -*-
"""
人事文档处理 — 统一界面（左侧导航 + 右侧 Frame 切换）

将「工资表拆分」和「Word 文档合并」整合为一个统一工具箱窗口。
使用方式：
    # 独立窗口
    app = UnifiedHRTools()
    app.run()

    # 嵌入到 Toplevel
    app = UnifiedHRTools(parent=window)
"""
import tkinter as tk
from tkinter import ttk

from .cf_gzbcf_hyp_gui import UnifiedSplitApp
from .word_gui_hyp import WordMergeApp


class UnifiedHRTools:
    """人事文档处理统一工具箱"""

    # 导航项定义
    NAV_ITEMS = [
        {'id': 'salary_split', 'icon': '🧾', 'label': '工资表拆分',
         'desc': '按部门列将工资表拆分为独立Excel文件'},
        {'id': 'word_merge', 'icon': '📝', 'label': 'Word 文档合并',
         'desc': '智能按目录层级合并，保留完整格式'},
    ]

    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
            self._standalone = True
        else:
            self.root = parent
            self._standalone = False

        self._current = None       # 当前选中的导航项 ID
        self._content_widgets = {} # 已创建的内容 widget {id: widget}

        self._build_ui()

        if self._standalone:
            self.root.title("📦 人事文档处理工具箱")
            self.root.geometry("1100x750")

    def _build_ui(self):
        container = self.root

        # ---- 左侧导航栏 ----
        self.nav_frame = tk.Frame(container, bg='#2C3E50', width=180)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        # 导航标题
        nav_title = tk.Label(
            self.nav_frame,
            text="📂 功能导航",
            font=("Microsoft YaHei", 13, "bold"),
            bg='#2C3E50', fg='#ECF0F1',
            anchor='w'
        )
        nav_title.pack(fill=tk.X, padx=16, pady=(20, 12))

        # 分隔线
        sep = ttk.Separator(self.nav_frame, orient='horizontal')
        sep.pack(fill=tk.X, padx=12, pady=(0, 8))

        # 导航按钮
        self._nav_buttons = {}
        for item in self.NAV_ITEMS:
            btn_frame = tk.Frame(self.nav_frame, bg='#2C3E50',
                                 cursor='hand2')
            btn_frame.pack(fill=tk.X, padx=8, pady=2)

            btn_inner = tk.Frame(btn_frame, bg='#34495E', padx=4, pady=10)
            btn_inner.pack(fill=tk.X)
            btn_inner.bind('<Button-1>', lambda e, iid=item['id']: self._switch_to(iid))
            btn_frame.bind('<Button-1>', lambda e, iid=item['id']: self._switch_to(iid))

            icon_label = tk.Label(
                btn_inner, text=item['icon'],
                font=("Segoe UI Emoji", 14),
                bg='#34495E', fg='#ECF0F1'
            )
            icon_label.pack(side=tk.LEFT, padx=(6, 10))

            text_label = tk.Label(
                btn_inner, text=item['label'],
                font=("Microsoft YaHei", 11),
                bg='#34495E', fg='#ECF0F1', anchor='w'
            )
            text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # 绑定 hover 效果
            for w in [btn_frame, btn_inner, icon_label, text_label]:
                w.bind('<Enter>', lambda e, b=btn_inner: b.configure(bg='#3D566E'))
                w.bind('<Leave>', lambda e, b=btn_inner: b.configure(bg='#34495E'))
                w.bind('<Button-1>', lambda e, iid=item['id']: self._switch_to(iid))

            self._nav_buttons[item['id']] = {
                'frame': btn_frame,
                'inner': btn_inner,
                'text': text_label,
            }

        # 底部提示
        hint = tk.Label(
            self.nav_frame,
            text="选择左侧功能\n开始操作",
            font=("Microsoft YaHei", 9),
            bg='#2C3E50', fg='#7F8C8D',
            justify='center'
        )
        hint.pack(side=tk.BOTTOM, padx=0, pady=15)

        # ---- 右侧内容区 ----
        self.content_area = tk.Frame(container, bg='#F5F6FA')
        self.content_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 欢迎页面（默认）
        self._show_welcome()

    def _show_welcome(self):
        """显示欢迎/占位页"""
        for w in self.content_area.winfo_children():
            w.destroy()

        welcome = tk.Frame(self.content_area, bg='#F5F6FA')
        welcome.place(relx=0.5, rely=0.5, anchor='center')

        tk.Label(
            welcome, text="👈 请从左侧导航选择一个工具",
            font=("Microsoft YaHei", 16),
            bg='#F5F6FA', fg='#95A5A6'
        ).pack(pady=20)

        tk.Label(
            welcome,
            text="\n".join(f"  {item['icon']} {item['label']} — {item['desc']}"
                          for item in self.NAV_ITEMS),
            font=("Microsoft YaHei", 11),
            bg='#F5F6FA', fg='#BDC3C7',
            justify='left'
        ).pack()

    def _switch_to(self, item_id):
        """切换到指定功能"""
        if self._current == item_id:
            return

        # 更新导航高亮
        for nid, widgets in self._nav_buttons.items():
            color = '#E67E22' if nid == item_id else '#ECF0F1'
            inner_color = '#E67E22' if nid == item_id else '#34495E'
            bg_color = '#3D566E' if nid == item_id else '#34495E'
            widgets['inner'].configure(bg=bg_color)
            widgets['text'].configure(fg=color)
            for child in widgets['inner'].winfo_children():
                if isinstance(child, tk.Label) and 'icon' not in str(child.cget('text')):
                    pass  # skip non-icon labels
                child.configure(bg=bg_color)

        # 清空内容区并加载新组件
        for w in self.content_area.winfo_children():
            w.destroy()

        self._current = item_id

        if item_id == 'salary_split':
            self._load_salary_split()
        elif item_id == 'word_merge':
            self._load_word_merge()

    def _load_salary_split(self):
        """加载工资表拆分工具（嵌入模式）"""
        try:
            app = UnifiedSplitApp(parent=self.content_area)
            self._content_widgets['salary_split'] = app
        except Exception as e:
            self._show_error("工资表拆分工具", str(e))

    def _load_word_merge(self):
        """加载 Word 合并工具（嵌入模式）"""
        try:
            app = WordMergeApp(master=self.content_area)
            self._content_widgets['word_merge'] = app
        except Exception as e:
            self._show_error("Word 文档合并工具", str(e))

    def _show_error(self, tool_name, error_msg):
        """显示错误信息页"""
        err_frame = tk.Frame(self.content_area, bg='#F5F6FA')
        err_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)

        tk.Label(
            err_frame, text=f"⚠️ {tool_name} 加载失败",
            font=("Microsoft YaHei", 14, "bold"),
            bg='#F5F6FA', fg='#E74C3C'
        ).pack(pady=(30, 10))

        tk.Label(
            err_frame, text=error_msg,
            font=("Microsoft YaHei", 10),
            bg='#F5F6FA', fg='#C0392B',
            wraplength=600, justify='left'
        ).pack(pady=(0, 10))

    def run(self):
        """独立模式入口"""
        if self._standalone:
            self.root.mainloop()


def main():
    app = UnifiedHRTools()
    app.run()


if __name__ == '__main__':
    main()
