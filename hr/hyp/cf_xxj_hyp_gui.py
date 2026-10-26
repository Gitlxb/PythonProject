# -*- coding: utf-8 -*-
"""
孝心金拆分 — 界面层（tkinter）

支持两种使用方式：
    独立窗口：XXJSplitApp().run()
    嵌入框架：XXJSplitApp(parent=frame)        # 用于「人事文档处理工具箱」

界面流程：
    选择 Excel -> 选择工作表（自动读第 2 行表头）-> 选择拆分依据列
    -> 预览区列出识别到的分组（同名分组用行范围区分）
    -> 【直接拆分】每个分组输出一个文件 / 【组合拆分】勾选若干分组合并成一个文件
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from .cf_xxj_hyp import XXJSplitter
except ImportError:
    from cf_xxj_hyp import XXJSplitter


class XXJGroupDialog(tk.Toplevel):
    """组合拆分弹窗：左侧多选分组并添加到右侧组合，一个组合输出一个文件"""

    def __init__(self, parent, segments):
        super().__init__(parent)
        self.title("组合拆分")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # 修复 Win+D 最小化后 Alt+Tab 无法恢复、界面卡死的问题：
        # grab_set() 在窗口最小化后仍占用输入，恢复时主窗口可见但输入被重定向到不可见的对话框。
        self._grab_active = True
        self.bind("<Unmap>", self._on_unmap)
        self.bind("<Map>", self._on_map)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        self._all_segments = list(segments)
        self.groups = []      # [[SplitSegment, ...], ...]
        self.result = None    # 确认后返回 groups

        self._build_ui()
        self._fit_to_content()

    def _fit_to_content(self):
        """按内容实际所需尺寸定窗口大小（留少量余量），避免按钮区被挤出可视范围"""
        self.update_idletasks()
        width = max(self.winfo_reqwidth() + 16, 740)
        height = min(max(self.winfo_reqheight() + 20, 480), 680)
        self.minsize(660, 480)
        self.geometry(f"{width}x{height}")

    def _on_unmap(self, event):
        if self._grab_active:
            try:
                self.grab_release()
            except tk.TclError:
                pass
            self._grab_active = False

    def _on_map(self, event):
        if not self._grab_active:
            try:
                self.grab_set()
            except tk.TclError:
                pass
            self._grab_active = True
        try:
            self.lift()
            self.focus_force()
        except tk.TclError:
            pass

    def _build_ui(self):
        # 底部按钮区最先 pack（side="bottom"）：pack 按调用顺序分配空间，
        # 若放在最后 pack，窗口高度不足时它会被挤出可视范围（按钮看不见）。
        bottom = tk.Frame(self)
        bottom.pack(side="bottom", fill="x", padx=15, pady=(8, 12))
        tk.Label(bottom, text="提示：未参与组合的分组将各自单独拆分",
                 font=("Microsoft YaHei", 9), fg="#888888").pack(side="left")
        tk.Button(bottom, text="取消", width=10, font=("Microsoft YaHei", 10),
                  bg="#95a5a6", fg="white", relief="flat", cursor="hand2",
                  command=self.destroy).pack(side="right")
        tk.Button(bottom, text="确认", width=10, font=("Microsoft YaHei", 10),
                  bg="#2980b9", fg="white", relief="flat", cursor="hand2",
                  command=self._confirm).pack(side="right", padx=(6, 0))

        tk.Label(
            self,
            text="在左侧多选分组，点击『添加到新组合』把选中的分组并成一个文件。\n"
                 "同名分组（如两段「浙江锦途」）用行范围区分；未参与组合的分组将各自单独输出。",
            font=("Microsoft YaHei", 9), fg="#555555", justify="left"
        ).pack(side="top", anchor="w", padx=15, pady=(12, 8))

        mid = tk.Frame(self)
        mid.pack(side="top", fill="both", expand=True, padx=15)

        # ---- 左：可选分组 ----
        left = tk.Frame(mid)
        left.pack(side="left", fill="both", expand=True)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        tk.Label(left, text="可选分组（多选，可双击）:", font=("Microsoft YaHei", 10, "bold")
                 ).grid(row=0, column=0, columnspan=2, sticky="w")
        self.list_available = tk.Listbox(
            left, selectmode="multiple", height=14,
            font=("Microsoft YaHei", 10), exportselection=False
        )
        self.list_available.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        sb_left = ttk.Scrollbar(left, orient="vertical", command=self.list_available.yview)
        sb_left.grid(row=1, column=1, sticky="ns", pady=(4, 6))
        self.list_available.configure(yscrollcommand=sb_left.set)
        self.list_available.bind("<Double-Button-1>", lambda _e: self._add_group())
        tk.Button(
            left, text="添加到新组合 >>", font=("Microsoft YaHei", 10),
            bg="#27ae60", fg="white", relief="flat", cursor="hand2",
            command=self._add_group
        ).grid(row=2, column=0, columnspan=2, sticky="ew")

        # ---- 右：已创建的组合 ----
        right = tk.Frame(mid)
        right.pack(side="right", fill="both", expand=True, padx=(12, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        tk.Label(right, text="已创建的组合:", font=("Microsoft YaHei", 10, "bold")
                 ).grid(row=0, column=0, columnspan=2, sticky="w")
        self.list_groups = tk.Listbox(
            right, height=14, font=("Microsoft YaHei", 10), exportselection=False
        )
        self.list_groups.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        sb_right = ttk.Scrollbar(right, orient="vertical", command=self.list_groups.yview)
        sb_right.grid(row=1, column=1, sticky="ns", pady=(4, 6))
        self.list_groups.configure(yscrollcommand=sb_right.set)
        tk.Button(
            right, text="<< 删除选中组合", font=("Microsoft YaHei", 10),
            bg="#e74c3c", fg="white", relief="flat", cursor="hand2",
            command=self._remove_group
        ).grid(row=2, column=0, columnspan=2, sticky="ew")

        self._refresh_available()
        self._refresh_groups()

    def _available(self):
        grouped = set()
        for g in self.groups:
            for seg in g:
                grouped.add(id(seg))
        return [s for s in self._all_segments if id(s) not in grouped]

    def _refresh_available(self):
        self.list_available.delete(0, "end")
        for seg in self._available():
            self.list_available.insert("end", seg.label)

    def _refresh_groups(self):
        self.list_groups.delete(0, "end")
        for i, g in enumerate(self.groups, 1):
            self.list_groups.insert("end", f"组合{i}: " + "、".join(s.label for s in g))

    def _add_group(self):
        sel = self.list_available.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先在左侧多选要组合的分组。", parent=self)
            return
        avail = self._available()
        self.groups.append([avail[i] for i in sel])
        self._refresh_available()
        self._refresh_groups()

    def _remove_group(self):
        sel = self.list_groups.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先在右侧选中要删除的组合。", parent=self)
            return
        idx = sel[0]
        if 0 <= idx < len(self.groups):
            del self.groups[idx]
            self._refresh_available()
            self._refresh_groups()

    def _confirm(self):
        self.result = self.groups
        self.destroy()


class XXJSplitApp:
    """孝心金拆分工具界面

    参数:
        parent: None      -> 创建独立 Tk 窗口
                tk.Widget -> UI 构建在该 widget 内（嵌入模式）
    """

    COLUMNS = ("name", "range", "count", "amount")
    # 列宽比例：拆分值 / 源文件行范围 / 行数 / 金额合计
    # 用比例而非固定宽度，避免富余宽度被某一列全部吃掉导致列间距失衡
    _COL_RATIO = (("name", 0.30), ("range", 0.26), ("count", 0.14), ("amount", 0.30))
    # 列宽上限：超宽窗口下不让某一列摊得过大（金额/行数这类短内容列收窄）
    _COL_CAP = {"name": 460, "range": 320, "count": 160, "amount": 300}

    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
            self._is_top_level = True
        else:
            self.root = parent.winfo_toplevel()
            self._is_top_level = False

        self.parent = parent
        # 弹窗归属：嵌入模式下指向嵌入容器，避免 messagebox/filedialog 把主窗口顶到前台
        self._dialog_parent = parent if parent is not None else self.root

        self.splitter = XXJSplitter(log=self._log)
        self.filepath = ""
        self.output_dir = ""
        self._tree_width = 0        # 列宽重算的防抖基准
        self._sash_done = False     # 分栏初始位置是否已设定

        self._build_ui()

        if self._is_top_level:
            self.root.title("孝心金拆分工具")
            self.root.geometry("840x720")
            self.root.minsize(760, 600)

    # ------------------------------------------------------------ UI
    def _build_ui(self):
        padx, pady = 10, 5
        container = self.parent if self.parent is not None else self.root
        container.columnconfigure(1, weight=1)
        container.rowconfigure(4, weight=1)

        tk.Label(container, text="Excel文件:").grid(row=0, column=0, sticky="e", padx=padx, pady=pady)
        self.entry_file = tk.Entry(container, width=48)
        self.entry_file.grid(row=0, column=1, sticky="ew", padx=padx, pady=pady)
        tk.Button(container, text="浏览", command=self._choose_file).grid(row=0, column=2, padx=padx, pady=pady)

        tk.Label(container, text="工作表:").grid(row=1, column=0, sticky="e", padx=padx, pady=pady)
        self.combo_sheet = ttk.Combobox(container, width=45, state="readonly")
        self.combo_sheet.grid(row=1, column=1, sticky="ew", padx=padx, pady=pady)
        self.combo_sheet.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        tk.Label(container, text="拆分依据:").grid(row=2, column=0, sticky="e", padx=padx, pady=pady)
        self.combo_col = ttk.Combobox(container, width=45, state="readonly")
        self.combo_col.grid(row=2, column=1, sticky="ew", padx=padx, pady=pady)
        self.combo_col.bind("<<ComboboxSelected>>", self._on_col_selected)

        btn_frame = tk.Frame(container)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10)
        tk.Button(btn_frame, text="直接拆分", width=15, font=("Microsoft YaHei", 10, "bold"),
                  bg="#27ae60", fg="white", relief="flat", cursor="hand2",
                  command=self._start_split).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="组合拆分", width=15, font=("Microsoft YaHei", 10, "bold"),
                  bg="#2980b9", fg="white", relief="flat", cursor="hand2",
                  command=self._choose_groups).pack(side=tk.LEFT, padx=5)
        self.btn_open_dir = tk.Button(btn_frame, text="打开输出目录", width=15,
                                      command=self._open_output_dir, state="disabled")
        self.btn_open_dir.pack(side=tk.LEFT, padx=5)

        # ---- 下方：可拖拽分割的「分组预览 / 处理日志」上下双区 ----
        # 用 PanedWindow 而不是各占一行的 grid：两个区域都想要空间时，
        # 固定行高会让底部区域被挤出可视范围，PanedWindow 由权重分配且可手动拖动。
        paned = ttk.PanedWindow(container, orient="vertical")
        paned.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=padx, pady=(0, pady))
        self.paned = paned

        style = ttk.Style()
        style.configure("XXJ.Treeview", rowheight=23, font=("Microsoft YaHei", 9))
        style.configure("XXJ.Treeview.Heading", font=("Microsoft YaHei", 9, "bold"))

        # 上区：分组预览
        preview = tk.Frame(paned)
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(1, weight=1)

        pv_head = tk.Frame(preview, bg="#EAF0F6")
        pv_head.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(pv_head, text="分组预览", font=("Microsoft YaHei", 9, "bold"),
                 bg="#EAF0F6", fg="#2C5282").pack(side="left", padx=(8, 8), pady=3)
        self.lbl_preview = tk.Label(pv_head, text="", font=("Microsoft YaHei", 9),
                                    bg="#EAF0F6", fg="#4A5568")
        self.lbl_preview.pack(side="left", pady=3)

        self.tree = ttk.Treeview(preview, columns=self.COLUMNS, show="headings",
                                 height=5, style="XXJ.Treeview")
        for cid, title, anchor in (
            ("name", "拆分值", "w"),
            ("range", "源文件行范围", "center"),
            ("count", "行数", "center"),
            ("amount", "金额合计", "e"),
        ):
            self.tree.heading(cid, text=title)
            self.tree.column(cid, width=160, anchor=anchor, stretch=False, minwidth=56)
        self.tree.grid(row=1, column=0, sticky="nsew")
        tree_vsb = ttk.Scrollbar(preview, orient="vertical", command=self.tree.yview)
        tree_vsb.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_vsb.set)
        self.tree.tag_configure("odd", background="#F6F8FB")
        self.tree.bind("<Configure>", self._on_tree_configure)

        paned.add(preview, weight=3)

        # 下区：处理日志
        logs = tk.Frame(paned)
        logs.columnconfigure(0, weight=1)
        logs.rowconfigure(1, weight=1)

        log_head = tk.Frame(logs, bg="#EAF0F6")
        log_head.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(log_head, text="处理日志", font=("Microsoft YaHei", 9, "bold"),
                 bg="#EAF0F6", fg="#2C5282").pack(side="left", padx=(8, 8), pady=3)
        self.lbl_log_hint = tk.Label(log_head, text="", font=("Microsoft YaHei", 9),
                                     bg="#EAF0F6", fg="#4A5568")
        self.lbl_log_hint.pack(side="left", pady=3)

        self.text_log = tk.Text(logs, height=5, wrap="none", state="disabled",
                                font=("Microsoft YaHei", 9), relief="flat",
                                bg="#FBFCFD", padx=6, pady=4)
        self.text_log.grid(row=1, column=0, sticky="nsew")
        log_vsb = ttk.Scrollbar(logs, orient="vertical", command=self.text_log.yview)
        log_vsb.grid(row=1, column=1, sticky="ns")
        log_hsb = ttk.Scrollbar(logs, orient="horizontal", command=self.text_log.xview)
        log_hsb.grid(row=2, column=0, sticky="ew")
        self.text_log.configure(yscrollcommand=log_vsb.set, xscrollcommand=log_hsb.set)

        paned.add(logs, weight=2)
        # 保证两区各有最低可用高度（部分 Tk 版本不支持 minsize，静默跳过）
        for pane, min_h in ((preview, 130), (logs, 150)):
            try:
                paned.paneconfigure(pane, minsize=min_h)
            except tk.TclError:
                pass
        paned.bind("<Configure>", self._init_sash, add="+")
        self.root.after(300, self._init_sash)

    def _init_sash(self, _event=None):
        """首次布局完成后把上下分栏定在 60%，保证预览与日志都有足够高度"""
        if self._sash_done:
            return
        try:
            height = self.paned.winfo_height()
            if height > 120:
                self.paned.sashpos(0, int(height * 0.60))
                self._sash_done = True
        except tk.TclError:
            pass

    def _on_tree_configure(self, event):
        """列宽按比例跟随控件宽度，富余宽度不再集中到某一列"""
        self._apply_col_widths(event.width)

    def _apply_col_widths(self, width):
        if width < 240 or abs(width - self._tree_width) < 12:
            return
        self._tree_width = width
        usable = width - 6
        for cid, ratio in self._COL_RATIO:
            self.tree.column(cid, width=max(56, min(int(usable * ratio), self._COL_CAP[cid])))

    def _log(self, msg: str):
        self.text_log.config(state="normal")
        self.text_log.insert(tk.END, msg + "\n")
        self.text_log.see(tk.END)
        self.text_log.config(state="disabled")
        try:
            (self.parent or self.root).update_idletasks()
        except tk.TclError:
            pass

    # ------------------------------------------------------------ 交互
    def _choose_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xlsm")], parent=self._dialog_parent
        )
        if not path:
            return
        self.filepath = path
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, path)
        self.btn_open_dir.config(state="disabled")
        self.output_dir = ""
        self._clear_preview()
        self._clear_log()

        try:
            self.splitter.load(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载文件：\n{e}", parent=self._dialog_parent)
            return

        self.combo_sheet["values"] = self.splitter.wb.sheetnames
        if self.splitter.wb.sheetnames:
            self.combo_sheet.current(self.splitter.wb.sheetnames.index(self.splitter.sheet_name))
        self._refresh_headers()

    def _on_sheet_selected(self, _event=None):
        name = self.combo_sheet.get()
        if not name or not self.splitter.filepath:
            return
        try:
            self.splitter.load(self.splitter.filepath, name)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载工作表：\n{e}", parent=self._dialog_parent)
            return
        self._clear_preview()
        self._clear_log()
        self._refresh_headers()

    def _refresh_headers(self):
        """用第 2 行表头填充拆分依据下拉，默认选中第一列（运营中心）"""
        values = [f"{idx}. {name}" for idx, name in self.splitter.headers]
        self.combo_col["values"] = values
        if values:
            self.combo_col.current(0)
            self._on_col_selected()

    def _on_col_selected(self, _event=None):
        text = self.combo_col.get()
        if not text or not self.splitter.ws:
            return
        col_idx = int(text.split(".")[0])
        try:
            self.splitter.prepare(col_idx)
        except Exception as e:
            messagebox.showerror("错误", f"解析拆分列失败：\n{e}", parent=self._dialog_parent)
            return
        self._refresh_preview()

    def _refresh_preview(self):
        self._clear_preview()
        rows = self.splitter.preview_rows()
        for i, r in enumerate(rows):
            amount = r["amount"]
            amount_txt = f"{amount:,}" if isinstance(amount, (int, float)) else ""
            self.tree.insert(
                "", "end",
                values=(r["name"], r["range"], r["count"], amount_txt),
                tags=("odd",) if i % 2 else (),
            )
        total_rows = sum(r["count"] for r in rows)
        text = f"共 {len(rows)} 个分组 / {total_rows} 行"
        if self.splitter.amount_col_idx:
            total_amount = sum(r["amount"] for r in rows if isinstance(r["amount"], (int, float)))
            text += f" / 金额合计 {total_amount:,}"
        self.lbl_preview.config(text=text)
        self.lbl_log_hint.config(text="")
        self._apply_col_widths(self.tree.winfo_width())

    def _clear_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.lbl_preview.config(text="")
        self.lbl_log_hint.config(text="")

    # ------------------------------------------------------------ 拆分
    def _start_split(self):
        if not self._ensure_ready():
            return
        self._run_split(groups=None)

    def _choose_groups(self):
        if not self._ensure_ready():
            return
        dialog = XXJGroupDialog(self._dialog_parent, self.splitter.segments)
        self.root.wait_window(dialog)
        if dialog.result is None:
            return
        if not dialog.result:
            messagebox.showinfo("提示", "未创建任何组合，将按「直接拆分」处理。", parent=self._dialog_parent)
            self._run_split(groups=None)
            return
        for i, grp in enumerate(dialog.result, 1):
            self._log(f"组合{i}：" + "、".join(s.label for s in grp))
        self._run_split(groups=dialog.result)

    def _ensure_ready(self) -> bool:
        if not self.splitter.ws:
            messagebox.showwarning("提示", "请先选择 Excel 文件。", parent=self._dialog_parent)
            return False
        if not self.combo_col.get():
            messagebox.showwarning("提示", "请先选择拆分依据列。", parent=self._dialog_parent)
            return False
        if not self.splitter.segments:
            messagebox.showwarning("提示", "未识别到可拆分的数据分组。", parent=self._dialog_parent)
            return False
        return True

    def _run_split(self, groups):
        self._clear_log()
        self.lbl_log_hint.config(text="拆分中…")
        try:
            self.root.config(cursor="watch")
            self.root.update_idletasks()
            results, out_dir = self.splitter.split(groups=groups)
            self.output_dir = out_dir
            self.btn_open_dir.config(state="normal")
            self.lbl_log_hint.config(text=f"已生成 {len(results)} 个文件")
            messagebox.showinfo(
                "完成",
                f"拆分完成，共生成 {len(results)} 个文件。\n\n输出目录：\n{out_dir}",
                parent=self._dialog_parent,
            )
        except Exception as e:
            self.lbl_log_hint.config(text="拆分失败")
            messagebox.showerror("拆分失败", str(e), parent=self._dialog_parent)
        finally:
            try:
                self.root.config(cursor="")
            except tk.TclError:
                pass

    def _clear_log(self):
        self.text_log.config(state="normal")
        self.text_log.delete("1.0", tk.END)
        self.text_log.config(state="disabled")

    def _open_output_dir(self):
        if self.output_dir and os.path.isdir(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showwarning("提示", "输出目录不存在。", parent=self._dialog_parent)

    def run(self):
        if self._is_top_level:
            self.root.mainloop()


def main():
    XXJSplitApp().run()


if __name__ == "__main__":
    main()
