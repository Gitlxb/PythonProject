# -*- coding: utf-8 -*-
"""
Word 文档合并工具 — 分类层级 GUI（v4）

Workflow（6步）：
  1. 选择评标办法文档（可多选）
  2. 选择搜索文件夹
  3. 自动提取分类关键词 → 分组就地编辑
  4. 一键搜索（精准→模糊两级，增量搜索）
  5. 确认要合并的文件（分组勾选 + 匹配方式标注）
  6. 选择输出路径 → 开始合并

v4 新特性：
  - 左右分栏布局（操作区 | 日志区）
  - 第3步关键词分组就地可编辑
  - 增量搜索（仅重搜修改过的分类）
  - 模式指示器（自动显示 0-X分 / 主观 / 满分值 检测模式）
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from .word_core_hyp2 import (
    extract_categories_from_documents,
    categorized_search,
    merge_documents_categorized,
)


class WordMergeApp:
    """Word 文档分类合并工具 GUI"""

    def __init__(self, master=None):
        self._standalone = (master is None)
        if self._standalone:
            self.root = tk.Tk()
            self.root.title("Word 文档合并工具（分类版）")
            self.root.geometry("1100x750")
        else:
            self.root = master

        self.root.configure(bg="#F5F5F5")

        # ---- 状态变量 ----
        self.standard_docs = []              # 评标办法文档路径列表
        self.folder_path_var = tk.StringVar() # 搜索文件夹路径
        self.categories = []                 # [{parent, type, keywords, matched_files, ...}, ...]
        self._result_entries = []            # [{type, ...}] 结果列表数据
        self.output_path_var = tk.StringVar()
        self._category_texts = []            # [(parent, Text_widget), ...] 第3步编辑面板
        self._last_keywords_snapshot = {}    # {cat_idx: frozenset(keywords)} 增量搜索快照
        self.detected_modes = set()          # 跨文档累计检测到的模式 {"0-X分", "主观", "满分值"}

        self._build_ui()

    @property
    def _parent_window(self):
        """返回顶层父窗口，用于文件对话框和消息框的 parent 参数。"""
        return self.root.winfo_toplevel()


    def _build_ui(self):
        # ---- 主分栏容器 ----
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # ===== 左侧：操作区 =====
        left_panel = ttk.Frame(paned)
        paned.add(left_panel, weight=1)

        # ---- 第1步 ----
        step1 = ttk.LabelFrame(left_panel, text=" 第1步：选择评标办法文档（可多选） ", padding="8")
        step1.pack(fill=tk.X, pady=(0, 6))
        step1.columnconfigure(1, weight=1)
        ttk.Label(step1, text="评标办法文档：").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        self.standard_docs_label = ttk.Label(step1, text="（未选择）", foreground="#888888")
        self.standard_docs_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        ttk.Button(step1, text="浏览", command=self._browse_standard_docs, width=8).grid(row=0, column=2)

        # ---- 第2步 ----
        step2 = ttk.LabelFrame(left_panel, text=" 第2步：选择搜索文件夹 ", padding="8")
        step2.pack(fill=tk.X, pady=(0, 6))
        step2.columnconfigure(1, weight=1)
        ttk.Label(step2, text="搜索文件夹：").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        ttk.Entry(step2, textvariable=self.folder_path_var).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 8))
        ttk.Button(step2, text="选择文件夹", command=self._browse_folder, width=12).grid(row=0, column=2)

        # ---- 第3步：分类关键词编辑面板 ----
        self.step3 = ttk.LabelFrame(left_panel, text=" 第3步：提取分类关键词（可编辑） ", padding="8")
        self.step3.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        self.step3.columnconfigure(1, weight=1)

        # 按钮行
        btn_row = ttk.Frame(self.step3)
        btn_row.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))
        ttk.Button(btn_row, text="提取关键词", command=self._extract_keywords, width=12).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(btn_row, text="提取后可在下方分类面板中直接编辑关键词（每行一个）",
                  font=("Microsoft YaHei UI", 8), foreground="#888888").pack(side=tk.LEFT)

        # 模式指示标签
        self.mode_indicator_var = tk.StringVar(value="")
        self.mode_indicator = ttk.Label(
            btn_row,
            textvariable=self.mode_indicator_var,
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#1976D2")
        self.mode_indicator.pack(side=tk.LEFT, padx=(12, 0))

        # 分类编辑滚动区域
        self.kw_canvas_frame = ttk.Frame(self.step3)
        self.kw_canvas_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.step3.rowconfigure(1, weight=1)

        # 创建 canvas + scrollbar
        self.kw_canvas = tk.Canvas(self.kw_canvas_frame, height=120, bg="#FAFAFA", highlightthickness=0)
        kw_scroll = ttk.Scrollbar(self.kw_canvas_frame, orient=tk.VERTICAL, command=self.kw_canvas.yview)
        self.kw_inner = ttk.Frame(self.kw_canvas)

        self.kw_canvas.configure(yscrollcommand=kw_scroll.set)
        kw_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.kw_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._kw_window_id = self.kw_canvas.create_window((0, 0), window=self.kw_inner, anchor="nw",
                                                          tags="kw_inner")
        self.kw_canvas.bind("<Configure>", self._on_kw_canvas_configure)
        self.kw_inner.bind("<Configure>", self._on_kw_inner_configure)

        # 占位提示
        self._kw_placeholder = ttk.Label(self.kw_inner, text="（请先点击「提取关键词」）",
                                         font=("Microsoft YaHei UI", 9), foreground="#AAAAAA")
        self._kw_placeholder.pack(pady=14)

        # ---- 第4步 ----
        step4 = ttk.LabelFrame(left_panel, text=" 第4步：搜索文件（精准→模糊 两级） ", padding="8")
        step4.pack(fill=tk.X, pady=(0, 6))
        step4.columnconfigure(1, weight=1)
        ttk.Button(step4, text="搜索文件", command=self._do_search, width=12).grid(row=0, column=0, padx=(0, 10))
        ttk.Label(step4, text="先精准匹配文件名，未命中则用4字分块模糊匹配",
                  font=("Microsoft YaHei UI", 8), foreground="#888888").grid(row=0, column=1, sticky=tk.W)

        # ---- 第5步 ----
        result_frame = ttk.LabelFrame(left_panel, text=" 第5步：搜索结果（单击切换勾选，组标题不可选） ", padding="6")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(1, weight=1)

        self.result_stats_var = tk.StringVar(value="（请先提取关键词并点击搜索）")
        ttk.Label(result_frame, textvariable=self.result_stats_var,
                  font=("Microsoft YaHei UI", 9), foreground="#666666").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 4))

        list_container = ttk.Frame(result_frame)
        list_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_container.columnconfigure(0, weight=1)
        list_container.rowconfigure(0, weight=1)

        self.result_listbox = tk.Listbox(
            list_container, selectmode=tk.EXTENDED,
            font=("Microsoft YaHei UI", 9), exportselection=False, height=6)
        self.result_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_scroll = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.result_listbox.yview)
        result_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.result_listbox.configure(yscrollcommand=result_scroll.set)

        result_btn_bar = ttk.Frame(result_frame)
        result_btn_bar.grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        ttk.Button(result_btn_bar, text="全选", command=self._select_all, width=8).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(result_btn_bar, text="取消全选", command=self._deselect_all, width=10).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(result_btn_bar, text="反选", command=self._toggle_all, width=8).pack(side=tk.LEFT)

        # ---- 第6步 + 开始合并 ----
        step6 = ttk.LabelFrame(left_panel, text=" 第6步：输出 + 合并 ", padding="8")
        step6.pack(fill=tk.X, pady=(0, 4))
        step6.columnconfigure(1, weight=1)
        ttk.Label(step6, text="输出路径：").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        ttk.Entry(step6, textvariable=self.output_path_var).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 8))
        ttk.Button(step6, text="浏览", command=self._browse_output, width=8).grid(row=0, column=2, pady=(0, 4))

        # 按钮行：开始合并 + 打开输出目录
        btn_row2 = ttk.Frame(step6)
        btn_row2.grid(row=1, column=0, columnspan=3, pady=(4, 0))
        self.merge_btn = tk.Button(
            btn_row2, text="▶ 开始合并", command=self._start_merge,
            font=("Microsoft YaHei UI", 11, "bold"), bg="#2196F3", fg="white",
            activebackground="#1976D2", width=14, height=2, cursor="hand2")
        self.merge_btn.pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(
            btn_row2, text="打开输出目录", command=self._open_output_dir,
            font=("Microsoft YaHei UI", 11), bg="#E0E0E0", fg="#333333",
            activebackground="#CCCCCC", width=14, height=2, cursor="hand2"
        ).pack(side=tk.LEFT)

        # ===== 右侧：操作记录（日志）=====
        right_panel = ttk.Frame(paned)
        paned.add(right_panel, weight=1)

        log_frame = ttk.LabelFrame(right_panel, text=" 操作记录 ", padding="6")
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame, wrap=tk.WORD,
            font=("Consolas", 9), bg="#F8F8F8",
            fg="#333333", insertbackground="#333333")
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=log_scroll.set)

    # ========================================================
    #  Canvas 滚动同步
    # ========================================================

    def _on_kw_canvas_configure(self, event):
        self.kw_canvas.itemconfig(self._kw_window_id, width=event.width)

    def _on_kw_inner_configure(self, event):
        self.kw_canvas.configure(scrollregion=self.kw_canvas.bbox("all"))

    # ========================================================
    #  日志
    # ========================================================

    def _log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # ========================================================
    #  浏览
    # ========================================================

    def _browse_standard_docs(self):
        paths = filedialog.askopenfilenames(parent=self._parent_window, 
            title="选择评标办法文档（可多选）",
            filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")])
        if paths:
            self.standard_docs = list(paths)
            n = len(self.standard_docs)
            self.standard_docs_label.config(
                text=os.path.basename(self.standard_docs[0]) if n == 1 else f"已选择 {n} 个文档",
                foreground="#000000")
            self._log(f"✓ 已选择 {n} 个评标办法文档")
            for p in self.standard_docs:
                self._log(f"    {os.path.basename(p)}")

    def _browse_folder(self):
        path = filedialog.askdirectory(parent=self._parent_window, title="选择搜索文件夹")
        if path:
            self.folder_path_var.set(path)
            self._log(f"✓ 已选择搜索文件夹: {path}")

    def _browse_output(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        initial_dir = desktop if os.path.exists(desktop) else None
        path = filedialog.asksaveasfilename(parent=self._parent_window, 
            title="选择输出文件路径", defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            initialfile="合并结果.docx", initialdir=initial_dir)
        if path:
            self.output_path_var.set(path)
            self._log(f"✓ 输出路径: {path}")

    @staticmethod
    def _increment_filename(base_path):
        """自动递增命名：合并结果.docx → 合并结果1.docx → 合并结果2.docx ..."""
        if not os.path.exists(base_path):
            return base_path
        stem, ext = os.path.splitext(base_path)
        i = 1
        while True:
            candidate = f"{stem}{i}{ext}"
            if not os.path.exists(candidate):
                return candidate
            i += 1

    def _open_output_dir(self):
        """打开输出文档所在的目录（如果已设置路径）。"""
        import subprocess
        path = self.output_path_var.get().strip()
        if not path:
            messagebox.showinfo("提示", "请先选择输出文件路径（浏览）或完成合并", parent=self._parent_window)
            return
        directory = os.path.dirname(path)
        if os.path.isdir(directory):
            subprocess.Popen(['explorer', directory])
            self._log(f"✓ 已打开目录: {directory}")
        else:
            messagebox.showwarning("提示", f"目录不存在：{directory}", parent=self._parent_window)

    # ========================================================
    #  提取关键词 + 构建分类编辑面板
    # ========================================================

    def _extract_keywords(self):
        if not self.standard_docs:
            messagebox.showwarning("提示", "请先选择评标办法文档（第1步）", parent=self._parent_window)
            return
        folder = self.folder_path_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择搜索文件夹（第2步）", parent=self._parent_window)
            return

        self._log("=" * 50)
        self._log("开始提取分类关键词...")

        success, result = extract_categories_from_documents(
            self.standard_docs, log_callback=self._log)

        if not success:
            messagebox.showerror("提取失败", result, parent=self._parent_window)
            return

        self.categories = result

        # ---- 检测模式并更新指示器 ----
        self.detected_modes = set()
        for cat in self.categories:
            t = cat.get("type", "")
            if t == "score":
                self.detected_modes.add("0-X分")
            elif t in ("subjective_split", "subjective_direct"):
                self.detected_modes.add("主观")
            elif t == "maxscore":
                self.detected_modes.add("满分值")
        self._update_mode_indicator()

        # ---- 构建分类编辑面板 ----
        self._build_category_editors()

        # 日志汇总
        all_kws = []
        for cat in self.categories:
            all_kws.extend(cat["keywords"])
        self._log(f"\n✓ 共提取 {len(self.categories)} 个类别, {len(all_kws)} 个关键词")
        for cat in self.categories:
            self._log(f"    [{cat['type']}] {cat['parent']}: {len(cat['keywords'])} 个关键词")

    def _build_category_editors(self):
        """在 kw_inner 中为每个分类创建编辑面板。"""
        # 清空旧内容
        for w in self.kw_inner.winfo_children():
            w.destroy()
        self._category_texts.clear()

        if not self.categories:
            self._kw_placeholder = ttk.Label(self.kw_inner, text="（未提取到分类关键词）",
                                             font=("Microsoft YaHei UI", 9), foreground="#AAAAAA")
            self._kw_placeholder.pack(pady=14)
            return

        for ci, cat in enumerate(self.categories):
            parent = cat["parent"]
            keywords = cat["keywords"]
            # 分类标题
            cat_frame = ttk.LabelFrame(self.kw_inner,
                                        text=f" {parent}（{len(keywords)} 个关键词） ",
                                        padding="4")
            cat_frame.pack(fill=tk.X, pady=(0, 4), padx=2)

            # 可编辑 Text 组件
            kw_lines = "\n".join(keywords)
            kw_text = tk.Text(cat_frame, height=max(2, min(len(keywords) + 1, 8)),
                              font=("Microsoft YaHei UI", 9), wrap=tk.WORD,
                              bg="#FFFFFF", relief=tk.SUNKEN, borderwidth=1)
            kw_text.insert("1.0", kw_lines)
            kw_text.pack(fill=tk.X, padx=2, pady=2)
            self._category_texts.append((parent, kw_text))

    # ========================================================
    #  模式指示器
    # ========================================================

    def _update_mode_indicator(self):
        """根据 detected_modes 更新步骤3的模式指示标签。"""
        if not self.detected_modes:
            self.mode_indicator_var.set("")
            return
        if len(self.detected_modes) == 1:
            mode = list(self.detected_modes)[0]
            colors = {"0-X分": "#2E7D32", "主观": "#E65100", "满分值": "#1976D2"}
            self.mode_indicator.configure(foreground=colors.get(mode, "#1976D2"))
            self.mode_indicator_var.set(f"● {mode}模式")
        else:
            self.mode_indicator.configure(foreground="#7B1FA2")
            modes = "、".join(sorted(self.detected_modes))
            self.mode_indicator_var.set(f"● 混合: {modes}")

    # ========================================================
    #  增量搜索
    # ========================================================

    def _sync_keywords_from_editors(self):
        """从编辑面板读取用户修改后的关键词，写回 self.categories。"""
        for ci, (parent, kw_text) in enumerate(self._category_texts):
            raw = kw_text.get("1.0", "end-1c")
            kws = [line.strip() for line in raw.split("\n") if line.strip()]
            if ci < len(self.categories):
                self.categories[ci]["keywords"] = kws

    def _do_search(self):
        self._sync_keywords_from_editors()

        folder = self.folder_path_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择搜索文件夹（第2步）", parent=self._parent_window)
            return
        if not self.categories:
            messagebox.showwarning("提示", "请先提取关键词（第3步）", parent=self._parent_window)
            return

        # ---- 判断增量搜索还是全量搜索 ----
        # 全量条件：首次搜索（无结果） 或 快照不存在
        first_run = (not self._result_entries)

        modified_indices = []
        unchanged_indices = []

        for ci, cat in enumerate(self.categories):
            current_kws = frozenset(cat["keywords"])
            last_kws = self._last_keywords_snapshot.get(ci)

            if first_run:
                # 首次搜索：全量
                self._last_keywords_snapshot[ci] = current_kws
                modified_indices.append(ci)
            elif last_kws is None:
                # 新类别
                self._last_keywords_snapshot[ci] = current_kws
                modified_indices.append(ci)
            elif current_kws != last_kws:
                # 关键词被修改
                self._last_keywords_snapshot[ci] = current_kws
                modified_indices.append(ci)
            else:
                unchanged_indices.append(ci)

        if modified_indices:
            self._log("=" * 50)
            if first_run:
                self._log("开始全量搜索...（所有类别）")
            else:
                self._log("开始增量搜索...（仅重搜修改过的类别）")
                for ci in unchanged_indices:
                    cat = self.categories[ci]
                    self._log(f"  [保留] {cat['parent']}: "
                              f"{len(cat.get('matched_files', []))} 个文件（关键词未变）")

            # 只搜索修改过的分类
            modified_cats = [self.categories[ci] for ci in modified_indices]
            categorized_search(modified_cats, folder, log_callback=self._log)
        else:
            self._log("关键词未变化，使用上次搜索结果")

        # ---- 构建结果列表 ----
        self._result_entries = []
        total_checked = 0
        total_files = 0

        for cat_idx, cat in enumerate(self.categories):
            files = cat.get("matched_files", [])
            parent = cat["parent"]
            self._result_entries.append({
                "type": "header",
                "text": f"[{parent}] — 匹配 {len(files)} 个文件",
                "category_idx": cat_idx,
            })

            for f_idx, fp in enumerate(files):
                fname = os.path.basename(fp)
                match_type = self._resolve_match_type(cat, fname)
                self._result_entries.append({
                    "type": "file",
                    "category_idx": cat_idx,
                    "file_idx": f_idx,
                    "path": fp,
                    "basename": fname,
                    "checked": True,
                    "match_type": match_type,
                })
                total_checked += 1
                total_files += 1

        self._refresh_result_list()

        # 绑定点击事件
        self.result_listbox.bind("<ButtonRelease-1>", self._on_click)

        if total_files == 0:
            messagebox.showwarning("搜索结果", "未找到匹配的文件", parent=self._parent_window)
        else:
            change_tag = "增量" if not first_run and unchanged_indices else "全量"
            self._log(f"\n✓ 搜索完成（{change_tag}），共匹配 {total_files} 个文件"
                      f"（{len(self.categories)} 个类别）")

    # ========================================================
    #  结果列表
    # ========================================================

    def _refresh_result_list(self):
        top_visible = self.result_listbox.yview()[0]
        self.result_listbox.delete(0, tk.END)

        checked_count = 0
        total_count = 0

        for entry in self._result_entries:
            if entry["type"] == "header":
                self.result_listbox.insert(tk.END, entry["text"])
                self.result_listbox.itemconfig(tk.END, fg="#888888", bg="#F0F0F0")
            elif entry["type"] == "file":
                mark = "☑" if entry["checked"] else "☐"
                mt = entry.get("match_type", "")
                tag_str = f"  [{mt}]" if mt else ""
                self.result_listbox.insert(tk.END, f"  {mark}  {entry['basename']}{tag_str}")
                if mt == "精准":
                    self.result_listbox.itemconfig(tk.END, fg="#2E7D32")
                elif mt == "模糊":
                    self.result_listbox.itemconfig(tk.END, fg="#E65100")
                total_count += 1
                if entry["checked"]:
                    checked_count += 1

        self.result_listbox.yview_moveto(top_visible)
        self.result_stats_var.set(
            f"共 {len(self.categories)} 个类别, "
            f"{total_count} 个匹配文件, "
            f"已勾选 {checked_count} 个（单击文件行切换勾选）")

    def _resolve_match_type(self, cat, filename):
        details = cat.get("search_details", [])
        has_precision = False
        has_fuzzy = False
        for kw, fname, method in details:
            if fname == filename:
                if method == "精准":
                    has_precision = True
                elif method == "模糊":
                    has_fuzzy = True
        if has_precision:
            return "精准"
        if has_fuzzy:
            return "模糊"
        return ""

    def _on_click(self, event):
        idx = self.result_listbox.curselection()
        if not idx:
            idx = self.result_listbox.nearest(event.y)
            if idx < 0:
                return
            idx = (idx,)
        i = idx[0]
        if i >= len(self._result_entries):
            return
        entry = self._result_entries[i]
        if entry["type"] != "file":
            return
        entry["checked"] = not entry["checked"]
        self._refresh_result_list()

    def _select_all(self):
        for e in self._result_entries:
            if e["type"] == "file":
                e["checked"] = True
        self._refresh_result_list()

    def _deselect_all(self):
        for e in self._result_entries:
            if e["type"] == "file":
                e["checked"] = False
        self._refresh_result_list()

    def _toggle_all(self):
        for e in self._result_entries:
            if e["type"] == "file":
                e["checked"] = not e["checked"]
        self._refresh_result_list()

    # ========================================================
    #  开始合并
    # ========================================================

    def _start_merge(self):
        output_path = self.output_path_var.get().strip()

        if not self.standard_docs:
            messagebox.showwarning("警告", "请先选择评标办法文档（第1步）", parent=self._parent_window)
            return
        if not self.folder_path_var.get() or not os.path.isdir(self.folder_path_var.get()):
            messagebox.showwarning("警告", "请先选择搜索文件夹（第2步）", parent=self._parent_window)
            return
        if not self._result_entries:
            messagebox.showwarning("警告", "请先搜索文件（第4步）", parent=self._parent_window)
            return

        # 收集勾选的文件（按分类）
        checked_by_cat = {}
        for entry in self._result_entries:
            if entry["type"] == "file" and entry["checked"]:
                ci = entry["category_idx"]
                checked_by_cat.setdefault(ci, []).append(entry["path"])

        if not checked_by_cat:
            messagebox.showwarning("警告", "请至少勾选一个要合并的文件", parent=self._parent_window)
            return

        merge_categories = []
        for ci, files in checked_by_cat.items():
            cat = self.categories[ci]
            merge_categories.append({
                "parent": cat["parent"],
                "type": cat["type"],
                "keywords": cat["keywords"],
                "matched_files": files,
            })

        if not output_path:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            base = os.path.join(desktop, "合并结果.docx") if os.path.exists(desktop) else "合并结果.docx"
            output_path = self._increment_filename(base)
            self.output_path_var.set(output_path)

        if not output_path.endswith(".docx"):
            output_path += ".docx"
            self.output_path_var.set(output_path)

        total_files = sum(len(c["matched_files"]) for c in merge_categories)
        self._log("=" * 50)
        self._log(f"开始分类合并: {len(merge_categories)} 个类别, {total_files} 个文件")
        for c in merge_categories:
            self._log(f"  {c['parent']}: {len(c['matched_files'])} 个文件")
        self._log("=" * 50)

        success, msg = merge_documents_categorized(
            merge_categories, output_path, log_callback=self._log)

        if success:
            self._log(f"\n✓ 合并成功！已保存到: {output_path}")
            messagebox.showinfo("完成", f"合并完成！\n\n保存路径：{output_path}", parent=self._parent_window)
        else:
            self._log(f"\n✗ 合并失败: {msg}")
            messagebox.showerror("错误", msg, parent=self._parent_window)

    # ========================================================
    #  独立运行入口
    # ========================================================

    def run(self):
        if self._standalone:
            self.root.mainloop()


if __name__ == "__main__":
    app = WordMergeApp()
    app.run()
