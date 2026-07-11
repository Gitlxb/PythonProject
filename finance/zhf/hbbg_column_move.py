# -*- coding: utf-8 -*-
"""
通用合并模式 — 列移动设置模块

功能：
  1) FileChipSelector: 标签式文件选择器（芯片 + ✕ 删除 + 自动换行）
  2) ColumnMovePanel: UI 组件（目标文件选择 + 移动列 + 右移格数）
  3) 多规则管理：同一文件可添加多条规则，自动按列索引倒序排列
  4) apply_column_moves(): 核心算法 — 在源文件加载后、复制前执行列右移

使用方式：
    from hbbg_column_move import ColumnMovePanel, apply_column_moves

    # 在 MergerFrame 中：
    self.col_move_panel = ColumnMovePanel(parent_frame)
    rules = self.col_move_panel.get_rules()  # 获取所有规则
    apply_columnmoves(worksheet, rules)       # 执行列移动
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl.utils import get_column_letter, column_index_from_string
from copy import copy


# ============================================================
#  数据结构
# ============================================================

class ColumnMoveRule:
    """单条列移动规则"""

    def __init__(self, target_files, column, offset):
        """
        参数:
            target_files: list[str] — 目标文件名列表（basename，如 ["2025年09月流水明细.xlsx"]）
            column: str — 要移动的列字母（如 "A", "J"）
            offset: int — 右移格数（正整数）
        """
        self.target_files = target_files  # list of basename
        self.column = column.upper()      # "A", "J" ...
        self.offset = max(0, int(offset))  # 右移格数

    @property
    def col_index(self):
        """列索引（1-based），用于排序"""
        return column_index_from_string(self.column)

    def __repr__(self):
        files = ",".join(self.target_files[:2])
        if len(self.target_files) > 2:
            files += f"...(+{len(self.target_files)-2})"
        return f"Move({files}, {self.column}+{self.offset})"

    def matches_file(self, basename):
        """判断此规则是否适用于指定文件"""
        return basename in self.target_files


# ============================================================
#  核心算法：列右移执行
# ============================================================

def apply_column_moves(ws, rules):
    """
    对工作表 ws 执行多条列移动规则。

    关键策略：按 col_index **降序** 排列后依次执行，
    避免先移左列导致右列索引漂移的问题。

    移动方式：在目标列位置插入 N 个空列（原位留空），
    openpyxl insert_cols 会将原列及之后的列全部右移。
    """
    if not rules:
        return 0

    # 按列索引降序排列（先移高列号，再移低列号，避免索引漂移）
    sorted_rules = sorted(rules, key=lambda r: r.col_index, reverse=True)

    total_inserted = 0
    for rule in sorted_rules:
        col_idx = rule.col_index
        offset = rule.offset
        if offset <= 0:
            continue

        # 连续调用 insert_cols(offset 次) = 右移 offset 格
        for _ in range(offset):
            ws.insert_cols(col_idx)

        total_inserted += offset
        print(f"  [列移动] {rule.column} 列 → 右移 {offset} 格"
              f"(目标文件: {','.join(rule.target_files[:2])})")

    return total_inserted


def get_rules_for_file(all_rules, basename):
    """获取适用于指定文件的所有列移动规则（已排序）"""
    matched = [r for r in all_rules if r.matches_file(basename)]
    return sorted(matched, key=lambda r: r.col_index, reverse=True)


# ============================================================
#  UI 组件：标签式文件选择器
# ============================================================

CHIP_BG = "#D6EAF8"          # 芯片背景色（浅蓝）
CHIP_BG_HOVER = "#AED6F1"    # 鼠标悬停时
CHIP_FG = "#1A5276"          # 芯片文字色
CHIP_BORDER = "#85C1E9"      # 芯片边框
CHIP_CLOSE_BG = "#E8F4FD"    # ✕ 按钮背景
CHIP_CLOSE_FG = "#999999"    # ✕ 按钮默认色
CHIP_CLOSE_FG_HOVER = "#E74C3C"  # ✕ 悬停变红
MAX_CHIPS_PER_ROW = 3        # 每行最多芯片数（自动换行）


class FileChip(tk.Frame):
    """单个文件标签芯片：文件名 + ✕ 关闭按钮"""

    def __init__(self, parent, filename, on_remove, **kwargs):
        super().__init__(parent, bg=CHIP_BG, relief="solid", bd=1, **kwargs)
        self._filename = filename
        self._on_remove = on_remove

        # 绑定整行背景色，防止 pack 内部缝隙露出父背景
        self.configure(highlightbackground=CHIP_BORDER, highlightthickness=1)

        # 文件名标签
        self._label = tk.Label(
            self, text=filename,
            bg=CHIP_BG, fg=CHIP_FG,
            font=("Microsoft YaHei UI", 8),
            padx=1, pady=1
        )
        self._label.pack(side=tk.LEFT, padx=(6, 0), pady=2)

        # ✕ 关闭按钮
        self._close_btn = tk.Label(
            self, text="✕",
            bg=CHIP_BG, fg=CHIP_CLOSE_FG,
            font=("Microsoft YaHei UI", 9, "bold"),
            padx=2, pady=1, cursor="hand2"
        )
        self._close_btn.pack(side=tk.LEFT, padx=(0, 4), pady=2)

        # 绑定事件
        self._close_btn.bind("<Button-1>", lambda e: self._on_remove(self._filename))
        self._close_btn.bind("<Enter>", self._on_close_enter)
        self._close_btn.bind("<Leave>", self._on_close_leave)

    def _on_close_enter(self, event):
        self._close_btn.configure(fg=CHIP_CLOSE_FG_HOVER)

    def _on_close_leave(self, event):
        self._close_btn.configure(fg=CHIP_CLOSE_FG)

    @property
    def filename(self):
        return self._filename


class FileChipSelector(tk.Frame):
    """
    标签式文件选择器 — 以可视化的芯片标签展示已选文件。

    特性：
      - 芯片 + ✕ 关闭按钮，一目了然
      - 自动换行（每行固定 N 个芯片）
      - "＋ 添加文件" 按钮弹出下拉菜单
      - 支持 set_available() / get_selected() / clear()
    """

    def __init__(self, parent, available_callback=None, **kwargs):
        """
        参数:
            parent: 父容器
            available_callback: Callable[[], List[str]]
                返回当前可选文件名的回调
        """
        super().__init__(parent, **kwargs)
        self._available_cb = available_callback
        self._chips = {}          # basename -> FileChip
        self._selected_order = [] # 保持选中顺序

        self._build_ui()

    def _build_ui(self):
        """构建界面：芯片展示区 + 添加按钮（点击弹出多选对话框）"""
        # 芯片容器（用 grid 布局实现自动换行）
        self._chip_container = tk.Frame(self, bg=self.cget("bg") or "#F0F0F0")
        self._chip_container.pack(fill=tk.X, expand=True)

        # 空状态提示（初始显示）
        self._empty_hint = tk.Label(
            self._chip_container,
            text="点击下方「＋ 添加文件」选择目标文件",
            font=("Microsoft YaHei UI", 8),
            fg="#AAAAAA"
        )
        self._empty_hint.grid(row=0, column=0, sticky="w", pady=4)

        # "＋ 添加文件" 按钮 → 点击弹出多选对话框
        self._add_btn = tk.Button(
            self, text="＋ 添加文件",
            font=("Microsoft YaHei UI", 8),
            bg="#EBF5FB", fg="#2980B9",
            activebackground="#D6EAF8",
            cursor="hand2", relief="flat",
            padx=8, pady=2,
            command=self._open_multi_select_dialog
        )
        self._add_btn.pack(anchor=tk.W, pady=(4, 0))

    # ---- 公共接口 ----

    def get_selected(self):
        """返回已选文件的 basename 列表（按选择顺序）"""
        return list(self._selected_order)

    def set_available(self, all_files):
        """
        更新可选文件列表，清理已不存在于 all_files 中的芯片。

        参数:
            all_files: List[str] — 所有可选文件的 basename
        """
        # 移除不再可用的芯片
        to_remove = [f for f in self._selected_order if f not in all_files]
        for f in to_remove:
            self._remove_chip(f, rebuild=False)

        # 重建网格布局
        self._rebuild_grid()

    def clear(self):
        """清除所有已选文件"""
        for f in list(self._selected_order):
            self._remove_chip(f, rebuild=False)
        self._rebuild_grid()

    # ---- 内部方法 ----

    def _add_chips_batch(self, basenames):
        """批量添加多个文件芯片"""
        added = False
        for bn in basenames:
            if bn not in self._chips and bn not in self._selected_order:
                if self._empty_hint.winfo_exists():
                    self._empty_hint.grid_forget()
                chip = FileChip(self._chip_container, bn, self._remove_chip)
                self._chips[bn] = chip
                self._selected_order.append(bn)
                added = True
        if added:
            self._rebuild_grid()
        return added

    def _open_multi_select_dialog(self):
        """弹出多选对话框，用户可一次性勾选多个目标文件"""
        # 获取当前可选文件列表
        available = []
        if self._available_cb:
            available = self._available_cb()

        # 过滤掉已选的
        unselected = [f for f in available if f not in self._selected_order]

        if not unselected:
            messagebox.showinfo("提示", "没有更多可选文件了", parent=self.winfo_toplevel())
            return

        # ---- 弹出 Toplevel 对话框 ----
        dlg = tk.Toplevel(self)
        dlg.title("选择目标文件")
        dlg.resizable(True, True)
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        # 居中显示
        dlg.update_idletasks()
        pw = 400; ph = 320
        sx = dlg.winfo_screenwidth(); sy = dlg.winfo_screenheight()
        dlg.geometry(f"{pw}x{ph}+{(sx-pw)//2}+{(sy-ph)//2}")

        main_frame = ttk.Frame(dlg, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="勾选需要添加的目标文件（支持 Ctrl+A 全选）：",
                  font=("Microsoft YaHei UI", 9)).pack(anchor=tk.W)

        # 全选/反选工具栏
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill=tk.X, pady=(4, 0))

        self._select_all_var = tk.BooleanVar(value=False)
        cb_all = ttk.Checkbutton(
            toolbar, text="全选", variable=self._select_all_var,
            command=lambda: self._toggle_all_checks(unselected)
        )
        cb_all.pack(side=tk.LEFT)

        ttk.Label(toolbar, text=f"（共 {len(unselected)} 个可选）",
                  foreground="#888888").pack(side=tk.LEFT, padx=(8, 0))

        # 文件列表 + 滚动条
        list_container = ttk.Frame(main_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        scrollbar = tk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        lb = tk.Listbox(
            list_container,
            selectmode=tk.EXTENDED,   # 支持 Shift/Ctrl 多选 + Ctrl+A
            font=("Microsoft YaHei UI", 9),
            yscrollcommand=scrollbar.set,
            height=12
        )
        lb.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=lb.yview)

        # 填充未选文件名
        for fn in unselected:
            lb.insert(tk.END, fn)

        # 绑定空格键切换选中状态
        def on_space(event):
            sel = lb.curselection()
            if sel:
                idx = sel[0]
                lb.selection_set(idx) if idx not in lb.curselection() else None
        lb.bind("<space>", on_space)

        # 确定按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(8, 0))

        result = [None]  # 用 list 包装以便在闭包中写入结果

        def on_ok():
            selected_indices = lb.curselection()
            chosen = [unselected[i] for i in selected_indices]
            if chosen:
                self._add_chips_batch(chosen)
            result[0] = chosen
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        tk.Button(btn_frame, text="✓ 确定",
                  bg="#2196F3", fg="white", font=("Microsoft YaHei UI", 9),
                  cursor="hand2", relief="flat", command=on_ok).pack(side=tk.RIGHT, padx=(4, 0))
        tk.Button(btn_frame, text="取消",
                  bg="#E0E0E0", fg="#333333", font=("Microsoft YaHei UI", 9),
                  cursor="hand2", relief="flat", command=on_cancel).pack(side=tk.RIGHT)

        # 绑定回车键确认、ESC 取消
        lb.bind("<Return>", lambda e: on_ok())
        dlg.bind("<Escape>", lambda e: on_cancel())

        # 双击直接添加单个
        def on_double_click(event):
            on_ok()
        lb.bind("<Double-Button-1>", on_double_click)

        dlg.wait_window()

    def _add_chip(self, basename):
        """添加一个文件芯片"""
        if basename in self._chips:
            return

        # 隐藏空提示
        if self._empty_hint.winfo_exists():
            self._empty_hint.grid_forget()

        chip = FileChip(self._chip_container, basename, self._remove_chip)
        self._chips[basename] = chip
        self._selected_order.append(basename)

        self._rebuild_grid()

    def _remove_chip(self, basename, rebuild=True):
        """移除一个文件芯片"""
        if basename not in self._chips:
            return

        chip = self._chips.pop(basename)
        chip.destroy()
        if basename in self._selected_order:
            self._selected_order.remove(basename)

        if rebuild:
            self._rebuild_grid()

        # 显示空提示（如果没有芯片了）
        if not self._chips and self._empty_hint.winfo_exists():
            self._empty_hint.grid(row=0, column=0, sticky="w", pady=4)

    def _toggle_all_checks(self, file_list):
        """全选/取消全选（由 Checkbutton 触发）"""
        # 此功能在 Listbox 模式下通过 Ctrl+A 实现，此处保留接口兼容
        pass

    def _rebuild_grid(self):
        """
        重建芯片的网格布局 — 自动换行。
        每行最多 MAX_CHIPS_PER_ROW 个，超出则换行。
        """
        # 清空所有 grid 中的 widget（不移除，只取消 grid 管理）
        for w in self._chip_container.grid_slaves():
            w.grid_forget()

        # 按选中顺序排列
        for idx, basename in enumerate(self._selected_order):
            chip = self._chips.get(basename)
            if not chip:
                continue
            row = idx // MAX_CHIPS_PER_ROW
            col = idx % MAX_CHIPS_PER_ROW
            chip.grid(row=row, column=col, sticky="w", padx=2, pady=2)

        # 更新空提示状态
        if not self._chips:
            self._empty_hint.grid(row=0, column=0, sticky="w", pady=4)
        elif self._empty_hint.winfo_exists():
            self._empty_hint.grid_forget()


# ============================================================
#  UI 组件：列移动设置面板
# ============================================================

COLUMN_LETTERS = [get_column_letter(i) for i in range(1, 53)]  # A~AZ
RULES_MAX_VISIBLE_HEIGHT = 130  # 规则区域最大可见高度（px）— 紧凑以留空间给操作按钮


class ScrollableRulesFrame(tk.Frame):
    """
    可滚动的规则容器：Canvas + 内部 Frame + Scrollbar。
    当内容超出最大高度时自动显示滚动条。
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        # 滚动条
        self._scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
        self._scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Canvas
        self._canvas = tk.Canvas(
            self, bg="#F9F9F9",
            yscrollcommand=self._scrollbar.set,
            highlightthickness=0
        )
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._scrollbar.configure(command=self._canvas.yview)

        # Canvas 内部 Frame（真正的规则容器）
        self.inner = tk.Frame(self._canvas, bg="#F9F9F9")
        self._canvas_window = self._canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )

        # 绑定事件：内容变化时更新滚动区域
        self.inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # 鼠标滚轮支持
        self._bind_mousewheel()

    def _on_inner_configure(self, event):
        """内部 Frame 大小变化时，更新 Canvas 的 scrollregion"""
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """Canvas 宽度变化时，调整内部 Frame 宽度"""
        self._canvas.itemconfig(self._canvas_window, width=event.width)

    def _bind_mousewheel(self):
        """绑定鼠标滚轮——仅在鼠标位于 Canvas 上方时滚动"""
        def _on_mousewheel(event):
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_enter(event):
            self._canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _on_leave(event):
            self._canvas.unbind_all("<MouseWheel>")

        self._canvas.bind("<Enter>", _on_enter)
        self._canvas.bind("<Leave>", _on_leave)


class ColumnMovePanel(ttk.LabelFrame):
    """
    「列移动设置」UI 面板 — 可选功能，嵌入到通用合并模式选项区中。

    外部接口:
        .get_rules() -> List[ColumnMoveRule]
        .set_available_files(files: List[str]) — 更新可选文件列表
    """

    def __init__(self, parent, available_files_callback=None):
        """
        参数:
            parent: 父容器 widget
            available_files_callback: Callable[[], List[str]]
                返回当前已选择的文件列表（basename），用于目标文件选择
        """
        super().__init__(parent, text="  列移动设置（可选）  ", padding=10)
        self._available_cb = available_files_callback
        self._rules = []          # List[ColumnMoveRule]
        self._rule_widgets = []   # 每条规则的 UI frame 列表

        self._build_ui()

    def _build_ui(self):
        # 提示文字
        hint = ttk.Label(
            self,
            text="💡 可对指定文件的指定列进行整体右移（插入空列），原位置留空",
            font=("Microsoft YaHei UI", 9), foreground="#888888"
        )
        hint.pack(anchor=tk.W, pady=(0, 8))

        # ---- 可滚动的规则列表容器（最大高度 RULES_MAX_VISIBLE_HEIGHT）----
        self._scroll_frame = ScrollableRulesFrame(
            self, bg="#F9F9F9", height=RULES_MAX_VISIBLE_HEIGHT
        )
        self._scroll_frame.pack(fill=tk.X, pady=(0, 6))
        self._scroll_frame.pack_propagate(False)  # 固定高度

        # 规则实际容器（在 Canvas 内部）
        self._rules_container = self._scroll_frame.inner

        # 添加按钮
        add_btn = tk.Button(
            self, text="＋ 添加列移动规则",
            font=("Microsoft YaHei UI", 10),
            bg="#E3F2FD", fg="#1565C0",
            activebackground="#BBDEFB",
            cursor="hand2", relief="flat",
            command=self._add_rule
        )
        add_btn.pack(anchor=tk.W)

        # 初始显示一条空提示
        self._empty_label = tk.Label(
            self._rules_container,
            text="  （暂无规则）",
            font=("Microsoft YaHei UI", 9),
            fg="#BBBBBB", bg="#F9F9F9"
        )
        self._empty_label.pack(pady=8)

    def _add_rule(self):
        """添加一条新的列移动规则 UI 卡片"""
        # 隐藏空提示
        if self._empty_label.winfo_exists():
            self._empty_label.destroy()

        idx = len(self._rule_widgets)

        card = ttk.Frame(self._rules_container)
        card.pack(fill=tk.X, pady=4)

        # --- 规则编号 + 删除按钮 ---
        header = ttk.Frame(card)
        header.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(header, text=f"规则 {idx + 1}",
                  font=("Microsoft YaHei UI", 9, "bold"),
                  foreground="#E65100").pack(side=tk.LEFT)

        del_btn = tk.Button(
            header, text="✕ 删除此规则",
            font=("Microsoft YaHei UI", 8),
            bg="#FFEBEE", fg="#C62828",
            activebackground="#FFCDD2",
            cursor="hand2", relief="flat",
            command=lambda c=card, i=idx: self._remove_rule(c, i)
        )
        del_btn.pack(side=tk.RIGHT)

        # --- 目标文件（使用标签式选择器）---
        row1 = ttk.Frame(card)
        row1.pack(fill=tk.X, pady=2)

        ttk.Label(row1, text="目标文件：", width=10,
                  font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, anchor=tk.N)

        # 芯片式文件选择器
        chip_selector = FileChipSelector(
            row1,
            available_callback=self._available_cb,
            bg="#F9F9F9"
        )
        chip_selector.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        # --- 移动列 ---
        row2 = ttk.Frame(card)
        row2.pack(fill=tk.X, pady=2)

        ttk.Label(row2, text="移动列：", width=10,
                  font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

        col_var = tk.StringVar(value="A")
        col_combo = ttk.Combobox(
            row2, textvariable=col_var,
            values=COLUMN_LETTERS, width=6,
            state="readonly", font=("Microsoft YaHei UI", 9)
        )
        col_combo.pack(side=tk.LEFT, padx=(0, 12))

        # --- 右移格数 ---
        ttk.Label(row2, text="右移格数：", font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

        offset_var = tk.StringVar(value="1")
        offset_spin = tk.Spinbox(
            row2, textvariable=offset_var,
            from_=1, to=50, width=5,
            font=("Microsoft YaHei UI", 9)
        )
        offset_spin.pack(side=tk.LEFT)

        # 保存引用
        widget_info = {
            'card': card,
            'chip_selector': chip_selector,
            'col_var': col_var,
            'offset_var': offset_var,
            'idx': idx,
        }
        self._rule_widgets.append(widget_info)

    def _remove_rule(self, card, idx):
        """删除一条规则"""
        card.destroy()
        if idx < len(self._rule_widgets):
            del self._rule_widgets[idx]
            # 更新后续规则编号
            for w in self._rule_widgets[idx:]:
                w['idx'] -= 1

        if not self._rule_widgets:
            self._empty_label = tk.Label(
                self._rules_container,
                text="  （暂无规则）",
                font=("Microsoft YaHei UI", 9),
                fg="#BBBBBB", bg="#F9F9F9"
            )
            self._empty_label.pack(pady=8)

    def get_rules(self):
        """
        收集所有有效规则并返回。

        返回: List[ColumnMoveRule]
            同一文件的多条规则会自动按 col_index 倒序排列（由 apply_column_moves 后续处理）
        """
        rules = []
        for info in self._rule_widgets:
            # 从芯片选择器获取已选文件
            chip_sel = info['chip_selector']
            files = chip_sel.get_selected()

            col = info['col_var'].get().strip().upper()
            try:
                offset = int(info['offset_var'].get())
            except ValueError:
                offset = 1

            if not files or not col or offset <= 0:
                continue

            # 验证列字母合法性
            if col not in COLUMN_LETTERS:
                continue

            rule = ColumnMoveRule(target_files=files[:], column=col, offset=offset)
            rules.append(rule)

        self._rules = rules
        return rules

    def set_available_files(self, files):
        """更新所有规则中的可选文件列表"""
        basenames = [os.path.basename(f) for f in files]
        for info in self._rule_widgets:
            chip_sel = info['chip_selector']
            chip_sel.set_available(basenames)

    def clear_all(self):
        """清除所有规则"""
        for info in self._rule_widgets[:]:
            self._remove_rule(info['card'], info['idx'])
        self._rules.clear()
