# -*- coding: utf-8 -*-
"""
弹窗对话框模块
包含：打款类型/月份选择弹窗、核对条件选择弹窗
"""

import tkinter as tk
from tkinter import ttk, messagebox


def show_selection_dialog(parent, payment_types, months, amount_columns):
    """
    显示打款类型、月份和打款金额列选择对话框

    Args:
        parent: 父窗口
        payment_types: 打款类型列表
        months: 月份列表
        amount_columns: 金额列列表 [(列索引, 列字母), ...]

    Returns:
        (selected_types, selected_months, selected_amount_col_idx)
        用户取消时返回 (None, None, None)
    """
    dialog = tk.Toplevel(parent)
    dialog.title("选择数据")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.resizable(False, False)

    # 主容器
    main_container = tk.Frame(dialog)
    main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

    # 上方区域：左侧打款类型 + 右侧月份
    top_frame = tk.Frame(main_container)
    top_frame.pack(fill=tk.BOTH, expand=True)

    # 左侧：打款类型（多选，默认不选中）
    type_vars = {}
    if payment_types:
        left_frame = ttk.LabelFrame(top_frame, text="打款类型")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        for pt in payment_types:
            var = tk.BooleanVar(value=False)
            type_vars[pt] = var
            ttk.Checkbutton(left_frame, text=pt, variable=var).pack(anchor=tk.W, pady=3, padx=8)

    # 右侧：月份（多选，默认不选中）
    month_vars = {}
    if months:
        right_frame = ttk.LabelFrame(top_frame, text="月份")
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        for m in months:
            var = tk.BooleanVar(value=False)
            month_vars[m] = var
            ttk.Checkbutton(right_frame, text=m, variable=var).pack(anchor=tk.W, pady=3, padx=8)

    # 中间区域：打款金额列选择（有多列时才显示）
    amount_var = tk.IntVar(value=-1)
    if len(amount_columns) >= 2:
        amount_frame = ttk.LabelFrame(main_container, text="打款金额列")
        amount_frame.pack(fill=tk.X, pady=(10, 0))
        for idx, col_letter in amount_columns:
            ttk.Radiobutton(
                amount_frame,
                text=f"打款金额（{col_letter}列）",
                variable=amount_var,
                value=idx
            ).pack(anchor=tk.W, pady=2, padx=8)

    # 按钮区域 - 底部居中
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=(10, 15))

    result = {"types": None, "months": None, "amount_col_idx": None}

    def on_confirm():
        selected_types = [t for t, v in type_vars.items() if v.get()] if payment_types else []
        selected_months = [m for m, v in month_vars.items() if v.get()] if months else []

        if payment_types and not selected_types:
            messagebox.showwarning("警告", "请至少选择一个打款类型", parent=dialog)
            return
        if months and not selected_months:
            messagebox.showwarning("警告", "请至少选择一个月份", parent=dialog)
            return
        if len(amount_columns) >= 2 and amount_var.get() == -1:
            messagebox.showwarning("警告", "请选择一个打款金额列", parent=dialog)
            return

        result["types"] = selected_types if payment_types else None
        result["months"] = selected_months if months else None
        if len(amount_columns) >= 2:
            result["amount_col_idx"] = amount_var.get()
        elif len(amount_columns) == 1:
            result["amount_col_idx"] = amount_columns[0][0]
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    ttk.Button(btn_frame, text="确定", command=on_confirm, width=10).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

    # 自适应大小并居中
    dialog.update_idletasks()
    width = dialog.winfo_reqwidth()
    height = dialog.winfo_reqheight()
    _center_dialog(dialog, parent, width, height)

    parent.wait_window(dialog)
    return result["types"], result["months"], result["amount_col_idx"]


def show_check_condition_dialog(parent, months, month_person_map):
    """
    显示月份（单选）和介绍人（多选）选择弹窗，月份与介绍人联动

    Args:
        parent: 父窗口
        months: 月份列表
        month_person_map: {月份: [介绍人列表]}

    Returns:
        (selected_month, selected_persons)
        用户取消时返回 (None, None)
    """
    dialog = tk.Toplevel(parent)
    dialog.title("核对条件")
    dialog.grab_set()
    dialog.resizable(True, True)
    dialog.geometry("900x650")

    # 主容器
    main_container = tk.Frame(dialog)
    main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
    main_container.columnconfigure(0, weight=0)
    main_container.columnconfigure(1, weight=1)
    main_container.rowconfigure(0, weight=1)

    # ===== 左侧：月份（单选） =====
    left_frame = ttk.LabelFrame(main_container, text="月份（单选）")
    left_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W), padx=(0, 10))

    month_var = tk.StringVar(value=months[0] if months else "")

    # ===== 右侧：介绍人（多选，多列网格+滚动+查询+全选/全不选） =====
    all_persons = sorted(set(p for plist in month_person_map.values() for p in plist))
    right_frame = ttk.LabelFrame(main_container, text="介绍人（多选）")
    right_frame.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.W, tk.E))
    right_frame.rowconfigure(1, weight=1)
    right_frame.columnconfigure(0, weight=1)

    # 工具栏：查询 + 全选 + 全不选
    toolbar = tk.Frame(right_frame)
    toolbar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

    search_var = tk.StringVar()
    search_entry = ttk.Entry(toolbar, textvariable=search_var, width=18)
    search_entry.pack(side=tk.LEFT, padx=(0, 5))
    ttk.Button(toolbar, text="查询", width=6).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Button(toolbar, text="全选", width=6).pack(side=tk.LEFT, padx=(0, 5))
    ttk.Button(toolbar, text="全不选", width=6).pack(side=tk.LEFT)

    # Canvas + 滚动条
    canvas = tk.Canvas(right_frame, highlightthickness=0)
    v_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
    h_scroll = ttk.Scrollbar(right_frame, orient="horizontal", command=canvas.xview)
    canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

    v_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
    h_scroll.grid(row=2, column=0, sticky=(tk.W, tk.E))
    canvas.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))

    # Canvas 内 Frame
    inner_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

    # 介绍人：每列20个，横向排列
    person_vars = {}
    checkbuttons = {}
    items_per_col = 20
    for idx, p in enumerate(all_persons):
        var = tk.BooleanVar(value=False)
        person_vars[p] = var
        col = idx // items_per_col
        row = idx % items_per_col
        cb = ttk.Checkbutton(inner_frame, text=p, variable=var)
        cb.grid(row=row, column=col, sticky=tk.W, pady=2, padx=8)
        checkbuttons[p] = cb

    # 更新 Canvas 滚动区域
    inner_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    # 联动过滤：根据当前月份 + 搜索条件，显示/隐藏并重新排列介绍人
    def filter_persons():
        current_month = month_var.get()
        valid_persons = set(month_person_map.get(current_month, []))
        query = search_var.get().strip().lower()
        visible_idx = 0
        for p, cb in checkbuttons.items():
            if p not in valid_persons:
                cb.grid_remove()
                continue
            if not query or query in p.lower():
                cb.grid()
                col = visible_idx // items_per_col
                row = visible_idx % items_per_col
                cb.grid_configure(row=row, column=col)
                visible_idx += 1
            else:
                cb.grid_remove()
        inner_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
        right_frame.configure(text=f"介绍人（多选，{current_month}，显示{visible_idx}/{len(valid_persons)}个）")

    # 月份切换回调
    def on_month_change():
        filter_persons()

    # 创建月份单选按钮
    for m in months:
        ttk.Radiobutton(left_frame, text=m, variable=month_var, value=m,
                        command=on_month_change).pack(anchor=tk.W, pady=3, padx=10)

    # 全选 / 全不选
    def select_all():
        current_month = month_var.get()
        valid_persons = set(month_person_map.get(current_month, []))
        query = search_var.get().strip().lower()
        for p, var in person_vars.items():
            if p in valid_persons and (not query or query in p.lower()):
                var.set(True)

    def deselect_all():
        current_month = month_var.get()
        valid_persons = set(month_person_map.get(current_month, []))
        query = search_var.get().strip().lower()
        for p, var in person_vars.items():
            if p in valid_persons and (not query or query in p.lower()):
                var.set(False)

    # 配置工具栏按钮命令
    for widget in toolbar.winfo_children():
        if isinstance(widget, ttk.Button):
            text = widget.cget("text")
            if text == "查询":
                widget.configure(command=filter_persons)
            elif text == "全选":
                widget.configure(command=select_all)
            elif text == "全不选":
                widget.configure(command=deselect_all)

    search_entry.bind("<Return>", lambda e: filter_persons())

    # 鼠标滚轮绑定
    def _on_mousewheel(event):
        canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _bind_wheel(event=None):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _unbind_wheel(event=None):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _bind_wheel)
    canvas.bind("<Leave>", _unbind_wheel)

    # 初始化：显示默认月份的介绍人
    if months:
        on_month_change()

    # ===== 底部按钮（居中） =====
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=(10, 15))

    result = {"month": None, "persons": None}

    def on_confirm():
        selected_month = month_var.get()
        valid_persons = set(month_person_map.get(selected_month, []))
        selected_persons = [p for p, v in person_vars.items() if v.get() and p in valid_persons]

        if not selected_month:
            messagebox.showwarning("警告", "请选择一个月份", parent=dialog)
            return
        if not selected_persons:
            messagebox.showwarning("警告", "请至少选择一个介绍人", parent=dialog)
            return

        result["month"] = selected_month
        result["persons"] = selected_persons
        _unbind_wheel()
        dialog.destroy()

    def on_cancel():
        _unbind_wheel()
        dialog.destroy()

    ttk.Button(btn_frame, text="确定", command=on_confirm, width=12).pack(side=tk.LEFT, padx=6)
    ttk.Button(btn_frame, text="取消", command=on_cancel, width=12).pack(side=tk.LEFT, padx=6)

    # 居中显示
    dialog.update_idletasks()
    _center_dialog(dialog, parent, dialog.winfo_reqwidth(), dialog.winfo_reqheight())

    parent.wait_window(dialog)
    return result["month"], result["persons"]


def _center_dialog(dialog, parent, width, height):
    """将对话框居中显示在主窗口上"""
    parent.update_idletasks()
    root_x = parent.winfo_x()
    root_y = parent.winfo_y()
    root_width = parent.winfo_width()
    root_height = parent.winfo_height()

    x = root_x + (root_width - width) // 2
    y = root_y + (root_height - height) // 2

    # 确保窗口不会超出屏幕
    screen_width = parent.winfo_screenwidth()
    screen_height = parent.winfo_screenheight()
    x = max(0, min(x, screen_width - width))
    y = max(0, min(y, screen_height - height))

    dialog.geometry(f"+{x}+{y}")
