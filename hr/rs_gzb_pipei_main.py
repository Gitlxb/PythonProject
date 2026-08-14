"""
GUI 界面模块：工资表匹配工具
tkinter 实现，与核心逻辑完全分离
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl

from hr.rs_gzb_pipei import (
    get_sheet_names,
    read_preview_rows,
    process_mode_a,
    process_mode_b_compare,
    apply_highlight_to_cells,
)


def _make_modal(dialog, toplevel):
    """把自定义 Toplevel 配置为可正确恢复的模态对话框。

    修复 Win+D 最小化后 Alt+Tab 无法回到操作界面、界面卡死（只能任务管理器
    结束任务）的问题：根因是 grab_set() 在窗口最小化后仍占用输入，恢复时主
    窗口可见但输入被重定向到不可见的对话框。这里在最小化时释放 grab、恢复时
    重建 grab 并置顶。
    """
    dialog.transient(toplevel)
    dialog.grab_set()
    state = {"grabbed": True}

    def on_unmap(event):
        if state["grabbed"]:
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            state["grabbed"] = False

    def on_map(event):
        if not state["grabbed"]:
            try:
                dialog.grab_set()
            except tk.TclError:
                pass
            state["grabbed"] = True
        try:
            dialog.lift()
            dialog.focus_force()
        except tk.TclError:
            pass

    dialog.bind("<Unmap>", on_unmap)
    dialog.bind("<Map>", on_map)
    return dialog


class SalaryMatchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("工资表匹配工具")
        self.root.geometry("750x620")
        self.root.minsize(700, 550)
        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main_frame,
            text="工资表匹配工具",
            font=("Microsoft YaHei", 18, "bold"),
        ).pack(pady=(0, 15))

        ttk.Label(
            main_frame,
            text="支持两种匹配模式：甲方与内部跨文件匹配 / 内部多工作簿互相匹配",
            font=("Microsoft YaHei", 10),
            foreground="gray",
        ).pack(pady=(0, 10))

        # 模式选择
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)

        ttk.Button(
            btn_frame,
            text="模式 A：甲方匹配内部",
            command=self.run_mode_a,
            width=28,
        ).pack(side=tk.LEFT, padx=12)

        ttk.Button(
            btn_frame,
            text="模式 B：内部互相匹配",
            command=self.run_mode_b,
            width=28,
        ).pack(side=tk.LEFT, padx=12)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = tk.Text(
            log_frame, height=20, wrap=tk.WORD, font=("Consolas", 10)
        )
        vsb = ttk.Scrollbar(
            log_frame, orient=tk.VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=vsb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # -------------------- 公共辅助方法 --------------------

    def _save_workbook(self, wb_out, current_file, default_suffix="_匹配结果.xlsx"):
        """保存 workbook 的通用方法"""
        if wb_out is None:
            return False

        self.log("\n所有匹配完成，准备保存...")
        default_name = (
            os.path.splitext(os.path.basename(current_file))[0]
            + default_suffix
        )
        save_path = self.ask_save_file(default_name)
        if save_path:
            try:
                wb_out.save(save_path)
                self.log(f"已保存：{save_path}")
                messagebox.showinfo(
                    "完成", f"处理完成！\n已保存至：{save_path}", parent=self.root
                )
                return True
            except Exception as e:
                messagebox.showerror("保存错误", str(e), parent=self.root)
                self.log(f"[保存错误] {e}")
                return False
        else:
            self.log("已取消保存")
            return False

    def _check_and_save_previous(self, wb_out, current_file, new_file, default_suffix="_匹配结果.xlsx"):
        """检查文件是否变更，如果变更则保存上一轮结果"""
        if current_file is not None and current_file != new_file:
            self.log("\n文件变更，保存上一轮结果...")
            default_name = (
                os.path.splitext(os.path.basename(current_file))[0]
                + default_suffix
            )
            save_path = self.ask_save_file(default_name)
            if save_path:
                try:
                    wb_out.save(save_path)
                    self.log(f"已保存：{save_path}")
                except Exception as e:
                    self.log(f"[保存错误] {e}")
            else:
                self.log("已取消保存，上一轮结果未保存")
            wb_out.close()
            return None  # 返回 None 表示需要重新加载 workbook
        return wb_out  # 返回原有的 wb_out

    # -------------------- 字段映射弹窗 --------------------

    def ask_field_mapping(self, headers1, headers2, title1="工作簿1", title2="工作簿2"):
        """弹窗让用户手动建立字段映射（按顺序）
        左右双Listbox单击选中 + 添加映射 + 智能匹配 + 上下调整顺序
        """
        dialog = tk.Toplevel(self.root)
        dialog.title("字段映射")
        dialog.geometry("850x650")
        _make_modal(dialog, self.root)
        dialog.resizable(False, False)

        # 顶部标题
        ttk.Label(
            dialog,
            text="请按顺序建立字段映射关系（按键值匹配后逐字段对比数据 + 公式）",
            font=("Microsoft YaHei", 10),
        ).pack(pady=6)

        # ===== 上方：左右字段列表 + 操作按钮 =====
        top_frame = ttk.Frame(dialog, padding=6)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        # 左侧字段列表
        left_frame = ttk.LabelFrame(top_frame, text=title1, padding=6)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)

        lb1 = tk.Listbox(
            left_frame,
            font=("Microsoft YaHei", 10),
            height=12,
            selectmode=tk.SINGLE,
            exportselection=False,
        )
        lb1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for h in headers1:
            lb1.insert(tk.END, h)
        sb1 = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=lb1.yview)
        lb1.config(yscrollcommand=sb1.set)
        sb1.pack(side=tk.RIGHT, fill=tk.Y)

        # 中间操作按钮
        mid_frame = ttk.Frame(top_frame, padding=6)
        mid_frame.pack(side=tk.LEFT, fill=tk.Y, padx=6)

        ttk.Label(mid_frame, text="单击选中左右字段\n再点击添加", font=("Microsoft YaHei", 9), foreground="gray").pack(pady=(20, 10))

        mappings = []

        def add_mapping():
            sel1 = lb1.curselection()
            sel2 = lb2.curselection()
            if not sel1 or not sel2:
                messagebox.showwarning("提示", "请先从左侧和右侧各单击选择一个字段", parent=dialog)
                return
            f1 = lb1.get(sel1[0])
            f2 = lb2.get(sel2[0])
            # 检查是否已存在
            for a, b in mappings:
                if a == f1 and b == f2:
                    messagebox.showwarning("提示", "该映射已存在", parent=dialog)
                    return
            mappings.append((f1, f2))
            listbox.insert(tk.END, f"{f1}   <-->   {f2}")
            # 取消选中，方便继续选下一对
            lb1.selection_clear(0, tk.END)
            lb2.selection_clear(0, tk.END)

        ttk.Button(mid_frame, text="添加映射", command=add_mapping, width=14).pack(pady=8)

        def auto_match():
            """智能匹配：自动将两边名称相同的字段建立映射"""
            matched = 0
            for h1 in headers1:
                if h1 in headers2:
                    # 检查是否已存在
                    exists = any(a == h1 and b == h1 for a, b in mappings)
                    if not exists:
                        mappings.append((h1, h1))
                        listbox.insert(tk.END, f"{h1}   <-->   {h1}")
                        matched += 1
            if matched:
                self.log(f"智能匹配：自动建立 {matched} 组映射")
            else:
                messagebox.showinfo("提示", "未找到名称相同的字段，请手动添加", parent=dialog)

        ttk.Button(mid_frame, text="智能匹配", command=auto_match, width=14).pack(pady=8)

        # 右侧字段列表
        right_frame = ttk.LabelFrame(top_frame, text=title2, padding=6)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)

        lb2 = tk.Listbox(
            right_frame,
            font=("Microsoft YaHei", 10),
            height=12,
            selectmode=tk.SINGLE,
            exportselection=False,
        )
        lb2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for h in headers2:
            lb2.insert(tk.END, h)
        sb2 = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=lb2.yview)
        lb2.config(yscrollcommand=sb2.set)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)

        # ===== 下方：已添加映射列表 + 操作 =====
        bottom_frame = ttk.Frame(dialog, padding=6)
        bottom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        list_frame = ttk.LabelFrame(
            bottom_frame, text="已添加的映射（按此顺序逐字段对比数据）", padding=6
        )
        list_frame.pack(fill=tk.BOTH, expand=True)

        listbox = tk.Listbox(list_frame, font=("Microsoft YaHei", 10), height=8)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_list = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=sb_list.set)
        sb_list.pack(side=tk.RIGHT, fill=tk.Y)

        # 映射操作按钮
        op_frame = ttk.Frame(bottom_frame)
        op_frame.pack(fill=tk.X, pady=6)

        def move_up():
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx == 0:
                return
            mappings[idx], mappings[idx - 1] = mappings[idx - 1], mappings[idx]
            # 刷新显示
            listbox.delete(0, tk.END)
            for a, b in mappings:
                listbox.insert(tk.END, f"{a}   <-->   {b}")
            listbox.selection_set(idx - 1)

        def move_down():
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx >= len(mappings) - 1:
                return
            mappings[idx], mappings[idx + 1] = mappings[idx + 1], mappings[idx]
            listbox.delete(0, tk.END)
            for a, b in mappings:
                listbox.insert(tk.END, f"{a}   <-->   {b}")
            listbox.selection_set(idx + 1)

        def remove_mapping():
            sel = listbox.curselection()
            if sel:
                idx = sel[0]
                listbox.delete(idx)
                mappings.pop(idx)

        def clear_all():
            if mappings and messagebox.askyesno("确认", "确定清空所有已添加的映射？", parent=dialog):
                listbox.delete(0, tk.END)
                mappings.clear()

        ttk.Button(op_frame, text="↑ 上移", command=move_up, width=12).pack(side=tk.LEFT, padx=4)
        ttk.Button(op_frame, text="↓ 下移", command=move_down, width=12).pack(side=tk.LEFT, padx=4)
        ttk.Button(op_frame, text="删除选中", command=remove_mapping, width=12).pack(side=tk.LEFT, padx=4)
        ttk.Button(op_frame, text="清空全部", command=clear_all, width=12).pack(side=tk.LEFT, padx=4)

        # ===== 底部确定/取消 =====
        result = [None]

        def on_ok():
            if not mappings:
                messagebox.showwarning("提示", "请至少添加一组字段映射", parent=dialog)
                return
            result[0] = list(mappings)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(side=tk.BOTTOM, pady=12)
        ttk.Button(btn_frame, text="确定", command=on_ok, width=12).pack(
            side=tk.LEFT, padx=10
        )
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=12).pack(
            side=tk.LEFT, padx=10
        )

        self.root.wait_window(dialog)
        return result[0]

    def ask_save_file(self, default_name="匹配结果.xlsx"):
        return filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            parent=self.root,
        )

    # -------------------- 公共辅助方法 --------------------

    def _auto_pick_key_field(self, var, headers):
        """智能推荐关键列：优先'姓名'/'名字'/'名称'，否则取第一项"""
        picked = ""
        for h in headers:
            if "姓名" in h or "名字" in h or "名称" in h:
                picked = h
                break
        if not picked and headers:
            picked = headers[0]
        var.set(picked)

    def _render_header_options(self, parent_frame, header_var, preview,
                               on_headers_changed=None, max_display_cols=15):
        """在指定 frame 中渲染表头行单选按钮，并通知当前表头字段"""
        for widget in parent_frame.winfo_children():
            widget.destroy()

        if not preview:
            if on_headers_changed:
                on_headers_changed([])
            return

        ttk.Label(parent_frame, text="表头行:").pack(anchor=tk.W)
        header_var.set(preview[0][0])

        def _notify(data):
            headers = [str(v).strip() for v in data if v is not None] if data else []
            if on_headers_changed:
                on_headers_changed(headers)

        for row_idx, row_data in preview:
            display_vals = []
            for i, val in enumerate(row_data[:max_display_cols]):
                if val is not None:
                    display_vals.append(f"列{i + 1}:[{val}]")
            display_text = "    ".join(display_vals) if display_vals else "(空行)"

            rb_frame = ttk.Frame(parent_frame)
            rb_frame.pack(fill=tk.X, pady=2)

            def make_cmd(data):
                def cmd():
                    _notify(data)
                return cmd

            ttk.Radiobutton(
                rb_frame, text=f"第 {row_idx} 行", variable=header_var, value=row_idx,
                command=make_cmd(row_data)
            ).pack(side=tk.LEFT)
            ttk.Label(
                rb_frame, text=display_text, font=("Consolas", 9), foreground="gray"
            ).pack(side=tk.LEFT, padx=10)

        # 初始化通知（默认第一行作为表头）
        _notify(preview[0][1])

    def _log_compare_result(self, total_info, compare_rows, diff_count, unmatched_keys):
        """统一输出匹配结果、目标表差异、合计行参与情况、未匹配人员等日志"""
        base_extra = "（含合计）" if total_info['base_has_total'] else ""
        self.log(
            f"对比完成：基准表 {total_info['base_data_rows']} 人{base_extra}"
            f"，共匹配 {compare_rows} 人，发现 {diff_count} 个差异单元格"
        )
        t_extra = "（含合计）" if total_info['target_has_total'] else ""
        self.log(f"  目标表数据行：{total_info['target_data_rows']} 人{t_extra}")

        # 目标表多出/缺少
        extra = total_info.get("extra_target_keys", [])
        if extra:
            self.log(f"  目标表中有 {len(extra)} 人未参与匹配：{', '.join(extra)}")
        elif total_info["base_data_rows"] > total_info["target_data_rows"]:
            diff = total_info["base_data_rows"] - total_info["target_data_rows"]
            self.log(f"  目标表缺少 {diff} 人未匹配")

        # 合计行参与情况
        if total_info["base_has_total"] and total_info["target_has_total"]:
            if total_info["participated"]:
                self.log(f"  【含合计行匹配】两边数据行数相等，合计行参与对比")
            else:
                self.log(f"  【合计行未参与】基准 {total_info['base_data_rows']} 行 vs"
                         f" 目标 {total_info['target_data_rows']} 行，数据行数不等，合计行不参与")
        elif total_info["base_has_total"]:
            self.log(f"  【合计行未参与】仅基准表有合计行，目标表无")
        elif total_info["target_has_total"]:
            self.log(f"  【合计行未参与】仅目标表有合计行，基准表无")

        # 未匹配人员
        if unmatched_keys:
            self.log(f"  未匹配人员（{len(unmatched_keys)} 人）：{', '.join(unmatched_keys)}")

    # -------------------- 综合配置弹窗 --------------------

    def _build_source_frame(self, parent, title, master, on_headers_changed=None):
        """构建一个Excel数据源配置面板，包含文件选择、工作簿选择、表头行选择。
        :param on_headers_changed: 可选回调，参数为当前headers列表（在切换工作簿/表头行时调用）
        """
        frame = ttk.LabelFrame(parent, text=title, padding=10)

        # 文件选择
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(file_frame, text="Excel文件:", width=10).pack(side=tk.LEFT)
        file_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=file_var, state="readonly", width=45)
        file_entry.pack(side=tk.LEFT, padx=5)

        sheet_var = tk.StringVar()
        header_var = tk.IntVar(value=1)

        # 工作簿选择
        sheet_frame = ttk.Frame(frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        ttk.Label(sheet_frame, text="工作簿:", width=10).pack(side=tk.LEFT)
        sheet_combo = ttk.Combobox(sheet_frame, textvariable=sheet_var, state="readonly", width=40, font=("Microsoft YaHei", 10))
        sheet_combo.pack(side=tk.LEFT, padx=5)

        # 表头行选择区域
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, pady=5)

        def on_sheet_selected(*args):
            path = file_var.get()
            sheet = sheet_var.get()
            if not path or not sheet:
                return
            try:
                preview = read_preview_rows(path, sheet, max_rows=3)
            except Exception as e:
                messagebox.showerror("错误", str(e), parent=master)
                return

            self._render_header_options(
                header_frame, header_var, preview,
                on_headers_changed=on_headers_changed
            )

        sheet_combo.bind("<<ComboboxSelected>>", on_sheet_selected)

        def on_browse():
            path = filedialog.askopenfilename(
                title=f"选择{title}的Excel文件",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
                parent=master,
            )
            if path:
                file_var.set(path)
                try:
                    sheets = get_sheet_names(path)
                except Exception as e:
                    messagebox.showerror("错误", str(e), parent=master)
                    return

                sheet_combo["values"] = sheets
                if len(sheets) == 1:
                    sheet_var.set(sheets[0])
                    on_sheet_selected()
                else:
                    sheet_var.set("")
                    for widget in header_frame.winfo_children():
                        widget.destroy()
                    if on_headers_changed:
                        on_headers_changed([])

        browse_btn = ttk.Button(file_frame, text="浏览...", command=on_browse, width=10)
        browse_btn.pack(side=tk.LEFT, padx=5)

        return {
            "frame": frame,
            "file_var": file_var,
            "sheet_var": sheet_var,
            "header_var": header_var,
        }

    def ask_mode_a_config(self):
        """模式A综合配置弹窗：一次性选择甲方和内部的文件、工作簿、表头行、关键列"""
        dialog = tk.Toplevel(self.root)
        dialog.title("模式A：配置甲方与内部工资表")
        dialog.geometry("850x800")
        _make_modal(dialog, self.root)
        dialog.resizable(False, False)

        ttk.Label(
            dialog,
            text="请配置甲方工资表和内部工资表的信息：",
            font=("Microsoft YaHei", 11),
        ).pack(pady=10)

        # 关键列变量
        jia_key_var = tk.StringVar()
        nei_key_var = tk.StringVar()

        def on_jia_headers(headers):
            jia_key_combo["values"] = headers
            self._auto_pick_key_field(jia_key_var, headers)

        def on_nei_headers(headers):
            nei_key_combo["values"] = headers
            self._auto_pick_key_field(nei_key_var, headers)

        # 甲方配置区
        jia = self._build_source_frame(dialog, "甲方工资表", dialog, on_headers_changed=on_jia_headers)
        jia["frame"].pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        # 内部配置区
        nei = self._build_source_frame(dialog, "内部工资表", dialog, on_headers_changed=on_nei_headers)
        nei["frame"].pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        # 关键列选择区
        key_frame = ttk.LabelFrame(dialog, text="关键列选择（用于行匹配，如\"姓名\"）", padding=10)
        key_frame.pack(fill=tk.X, padx=15, pady=8)

        kf1 = ttk.Frame(key_frame)
        kf1.pack(fill=tk.X, pady=4)
        ttk.Label(kf1, text="甲方关键列:", width=12).pack(side=tk.LEFT)
        jia_key_combo = ttk.Combobox(
            kf1, textvariable=jia_key_var, state="readonly",
            width=30, font=("Microsoft YaHei", 10)
        )
        jia_key_combo.pack(side=tk.LEFT, padx=5)

        kf2 = ttk.Frame(key_frame)
        kf2.pack(fill=tk.X, pady=4)
        ttk.Label(kf2, text="内部关键列:", width=12).pack(side=tk.LEFT)
        nei_key_combo = ttk.Combobox(
            kf2, textvariable=nei_key_var, state="readonly",
            width=30, font=("Microsoft YaHei", 10)
        )
        nei_key_combo.pack(side=tk.LEFT, padx=5)

        result = [None]

        def on_ok():
            jia_file = jia["file_var"].get()
            jia_sheet = jia["sheet_var"].get()
            jia_header = jia["header_var"].get()
            nei_file = nei["file_var"].get()
            nei_sheet = nei["sheet_var"].get()
            nei_header = nei["header_var"].get()
            jia_key = jia_key_var.get()
            nei_key = nei_key_var.get()

            if not jia_file:
                messagebox.showwarning("提示", "请选择甲方工资表文件", parent=dialog)
                return
            if not jia_sheet:
                messagebox.showwarning("提示", "请选择甲方工作簿", parent=dialog)
                return
            if not nei_file:
                messagebox.showwarning("提示", "请选择内部工资表文件", parent=dialog)
                return
            if not nei_sheet:
                messagebox.showwarning("提示", "请选择内部工作簿", parent=dialog)
                return
            if not jia_key:
                messagebox.showwarning("提示", "请选择甲方关键列", parent=dialog)
                return
            if not nei_key:
                messagebox.showwarning("提示", "请选择内部关键列", parent=dialog)
                return

            result[0] = (jia_file, jia_sheet, jia_header, nei_file, nei_sheet, nei_header, jia_key, nei_key)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(side=tk.BOTTOM, pady=15)
        ttk.Button(btn_frame, text="确定", command=on_ok, width=12).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=12).pack(side=tk.LEFT, padx=10)

        self.root.wait_window(dialog)
        return result[0]

    def ask_mode_b_config(self):
        """模式B综合配置弹窗：一次性选择Excel文件、两个工作簿、表头行、关键列"""
        dialog = tk.Toplevel(self.root)
        dialog.title("模式B：配置内部工作簿互相匹配")
        dialog.geometry("850x680")
        _make_modal(dialog, self.root)
        dialog.resizable(False, False)

        ttk.Label(
            dialog,
            text="请选择Excel文件及两个工作簿的信息：",
            font=("Microsoft YaHei", 11),
        ).pack(pady=6)

        # 文件选择
        file_frame = ttk.Frame(dialog)
        file_frame.pack(fill=tk.X, padx=15, pady=4)
        ttk.Label(file_frame, text="Excel文件:", width=10).pack(side=tk.LEFT)
        file_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=file_var, state="readonly", width=50)
        file_entry.pack(side=tk.LEFT, padx=5)

        sheet_a_var = tk.StringVar()
        header_a_var = tk.IntVar(value=1)
        sheet_b_var = tk.StringVar()
        header_b_var = tk.IntVar(value=1)

        # 关键列变量
        key_a_var = tk.StringVar()
        key_b_var = tk.StringVar()

        def on_a_headers(headers):
            key_a_combo["values"] = headers
            self._auto_pick_key_field(key_a_var, headers)

        def on_b_headers(headers):
            key_b_combo["values"] = headers
            self._auto_pick_key_field(key_b_var, headers)

        def _build_wb_section(parent, title, sheet_var_ref, header_var_ref):
            """构建固定高度的工作簿配置区域，返回 (外框frame, combo控件, header_frame)"""
            outer = ttk.Frame(parent, height=170)
            outer.pack(fill=tk.X, padx=15, pady=3)
            outer.pack_propagate(False)

            border = ttk.LabelFrame(outer, text=title, padding=6)
            border.pack(fill=tk.BOTH, expand=True)

            # 工作簿下拉
            top = ttk.Frame(border)
            top.pack(fill=tk.X, pady=3)
            ttk.Label(top, text="工作簿:", width=8).pack(side=tk.LEFT)
            combo = ttk.Combobox(
                top, textvariable=sheet_var_ref, state="readonly", width=45, font=("Microsoft YaHei", 10)
            )
            combo.pack(side=tk.LEFT, padx=5)

            # 表头行区域
            header_frame = ttk.Frame(border)
            header_frame.pack(fill=tk.X, pady=2)

            return outer, combo, header_frame

        # ========== 工作簿A ==========
        outer_a, combo_a, header_a_frame = _build_wb_section(
            dialog, "工作簿 A", sheet_a_var, header_a_var
        )

        def on_sheet_a_selected(*args):
            path = file_var.get()
            sheet = sheet_a_var.get()
            if not path or not sheet:
                on_a_headers([])
                return
            try:
                preview = read_preview_rows(path, sheet, max_rows=3)
            except Exception as e:
                messagebox.showerror("错误", str(e), parent=dialog)
                return
            self._render_header_options(
                header_a_frame, header_a_var, preview,
                on_headers_changed=on_a_headers
            )

        combo_a.bind("<<ComboboxSelected>>", on_sheet_a_selected)

        # ========== 工作簿B ==========
        outer_b, combo_b, header_b_frame = _build_wb_section(
            dialog, "工作簿 B", sheet_b_var, header_b_var
        )

        def on_sheet_b_selected(*args):
            path = file_var.get()
            sheet = sheet_b_var.get()
            if not path or not sheet:
                on_b_headers([])
                return
            try:
                preview = read_preview_rows(path, sheet, max_rows=3)
            except Exception as e:
                messagebox.showerror("错误", str(e), parent=dialog)
                return
            self._render_header_options(
                header_b_frame, header_b_var, preview,
                on_headers_changed=on_b_headers
            )

        combo_b.bind("<<ComboboxSelected>>", on_sheet_b_selected)

        def on_browse():
            path = filedialog.askopenfilename(
                title="选择内部工资表Excel文件",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
                parent=dialog,
            )
            if path:
                file_var.set(path)
                try:
                    sheets = get_sheet_names(path)
                except Exception as e:
                    messagebox.showerror("错误", str(e), parent=dialog)
                    return

                combo_a["values"] = sheets
                combo_b["values"] = sheets

                if len(sheets) == 1:
                    sheet_a_var.set(sheets[0])
                    sheet_b_var.set(sheets[0])
                    on_sheet_a_selected()
                    on_sheet_b_selected()
                else:
                    sheet_a_var.set("")
                    sheet_b_var.set("")
                    for widget in header_a_frame.winfo_children():
                        widget.destroy()
                    for widget in header_b_frame.winfo_children():
                        widget.destroy()
                    on_a_headers([])
                    on_b_headers([])

        ttk.Button(file_frame, text="浏览...", command=on_browse, width=10).pack(side=tk.LEFT, padx=5)

        # 关键列选择区
        key_frame = ttk.LabelFrame(dialog, text="关键列选择（用于行匹配，如\"姓名\"）", padding=10)
        key_frame.pack(fill=tk.X, padx=15, pady=8)

        kf1 = ttk.Frame(key_frame)
        kf1.pack(fill=tk.X, pady=4)
        ttk.Label(kf1, text="工作簿A关键列:", width=14).pack(side=tk.LEFT)
        key_a_combo = ttk.Combobox(
            kf1, textvariable=key_a_var, state="readonly",
            width=30, font=("Microsoft YaHei", 10)
        )
        key_a_combo.pack(side=tk.LEFT, padx=5)

        kf2 = ttk.Frame(key_frame)
        kf2.pack(fill=tk.X, pady=4)
        ttk.Label(kf2, text="工作簿B关键列:", width=14).pack(side=tk.LEFT)
        key_b_combo = ttk.Combobox(
            kf2, textvariable=key_b_var, state="readonly",
            width=30, font=("Microsoft YaHei", 10)
        )
        key_b_combo.pack(side=tk.LEFT, padx=5)

        # 按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        result = [None]

        def on_ok():
            path = file_var.get()
            sheet_a = sheet_a_var.get()
            header_a = header_a_var.get()
            sheet_b = sheet_b_var.get()
            header_b = header_b_var.get()
            key_a = key_a_var.get()
            key_b = key_b_var.get()

            if not path:
                messagebox.showwarning("提示", "请选择Excel文件", parent=dialog)
                return
            if not sheet_a:
                messagebox.showwarning("提示", "请选择工作簿A", parent=dialog)
                return
            if not sheet_b:
                messagebox.showwarning("提示", "请选择工作簿B", parent=dialog)
                return
            if not key_a:
                messagebox.showwarning("提示", "请选择工作簿A关键列", parent=dialog)
                return
            if not key_b:
                messagebox.showwarning("提示", "请选择工作簿B关键列", parent=dialog)
                return

            result[0] = (path, sheet_a, header_a, sheet_b, header_b, key_a, key_b)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        ttk.Button(btn_frame, text="确定", command=on_ok, width=12).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=12).pack(side=tk.LEFT, padx=10)

        self.root.wait_window(dialog)
        return result[0]

    # -------------------- 模式 A --------------------

    def run_mode_a(self):
        self.clear_log()
        self.log("===== 模式 A：甲方工资表匹配内部工资表 =====")

        round_num = 0
        wb_out = None
        current_nei_file = None

        try:
            while True:
                round_num += 1
                self.log(f"\n--- 第 {round_num} 轮匹配 ---")

                # 弹出综合配置窗口（一次性选择甲方和内部的文件、工作簿、表头行）
                config = self.ask_mode_a_config()
                if not config:
                    if round_num == 1:
                        return
                    break

                jia_file, jia_sheet, jia_header, nei_file, nei_sheet, nei_header, jia_key, nei_key = config

                self.log(f"甲方文件：{jia_file}")
                self.log(f"  工作簿：{jia_sheet}")
                self.log(f"  表头行：第 {jia_header} 行")
                self.log(f"  关键列：{jia_key}")

                preview = read_preview_rows(jia_file, jia_sheet, max_rows=3)
                jia_headers = []
                if jia_header <= len(preview):
                    jia_headers = [
                        str(v).strip()
                        for v in preview[jia_header - 1][1]
                        if v is not None
                    ]
                self.log(f"  表头字段：{jia_headers}")

                self.log(f"内部文件：{nei_file}")
                self.log(f"  工作簿：{nei_sheet}")
                self.log(f"  表头行：第 {nei_header} 行")
                self.log(f"  关键列：{nei_key}")

                preview2 = read_preview_rows(nei_file, nei_sheet, max_rows=3)
                nei_headers = []
                if nei_header <= len(preview2):
                    nei_headers = [
                        str(v).strip()
                        for v in preview2[nei_header - 1][1]
                        if v is not None
                    ]
                self.log(f"  表头字段：{nei_headers}")

                # 若内部文件变更，先保存上一轮结果
                wb_out = self._check_and_save_previous(wb_out, current_nei_file, nei_file)

                if wb_out is None:
                    current_nei_file = nei_file

                # 字段映射
                field_mapping = self.ask_field_mapping(
                    jia_headers, nei_headers, "甲方工作簿", "内部工作簿"
                )
                if not field_mapping:
                    self.log("已取消字段映射")
                    break

                self.log(f"字段映射：{field_mapping}")
                self.log(f"关键列：甲方「{jia_key}」 ↔ 内部「{nei_key}」")

                # 执行处理
                self.log("正在对比...")
                try:
                    wb_out, compare_rows, diff_count, unmatched_keys, total_info = process_mode_a(
                        jia_file,
                        jia_sheet,
                        jia_header,
                        nei_file,
                        nei_sheet,
                        nei_header,
                        jia_key,
                        nei_key,
                        field_mapping,
                        wb_out=wb_out,
                    )
                except Exception as e:
                    messagebox.showerror("处理错误", str(e), parent=self.root)
                    self.log(f"[错误] {e}")
                    break

                # 输出匹配结果日志
                self._log_compare_result(total_info, compare_rows, diff_count, unmatched_keys)

                self.log(f"已对工作簿【{nei_sheet}】应用标红")

                # 询问是否继续
                if not messagebox.askyesno(
                    "继续匹配", "是否继续匹配下一对（甲方 vs 内部）？", parent=self.root
                ):
                    break

            # 全部完成后保存
            if wb_out is not None:
                self._save_workbook(wb_out, current_nei_file)
                wb_out = None

        except Exception as e:
            if wb_out is not None:
                wb_out.close()
            messagebox.showerror("错误", str(e), parent=self.root)
            self.log(f"[错误] {e}")

    # -------------------- 模式 B --------------------

    def run_mode_b(self):
        self.clear_log()
        self.log("===== 模式 B：内部工资表互相匹配 =====")

        round_num = 0
        wb_out = None
        current_file = None

        try:
            while True:
                round_num += 1
                self.log(f"\n--- 第 {round_num} 轮匹配 ---")

                # 弹出综合配置窗口（一次性选择文件、工作簿A、工作簿B及表头行）
                config = self.ask_mode_b_config()
                if not config:
                    if round_num == 1:
                        return
                    break

                file_path, sheet_a, header_a, sheet_b, header_b, key_a, key_b = config

                self.log(f"文件：{file_path}")
                self.log(f"  工作簿 A：{sheet_a}，表头行：{header_a}，关键列：{key_a}")
                self.log(f"  工作簿 B：{sheet_b}，表头行：{header_b}，关键列：{key_b}")

                # 如果文件变更，保存上一轮
                wb_out = self._check_and_save_previous(wb_out, current_file, file_path)

                # 加载输出workbook
                if wb_out is None:
                    try:
                        wb_out = openpyxl.load_workbook(file_path)
                        current_file = file_path
                    except Exception as e:
                        messagebox.showerror("错误", f"无法打开文件：{e}", parent=self.root)
                        break

                preview_a = read_preview_rows(file_path, sheet_a, max_rows=3)
                headers_a = []
                if header_a <= len(preview_a):
                    headers_a = [
                        str(v).strip()
                        for v in preview_a[header_a - 1][1]
                        if v is not None
                    ]

                preview_b = read_preview_rows(file_path, sheet_b, max_rows=3)
                headers_b = []
                if header_b <= len(preview_b):
                    headers_b = [
                        str(v).strip()
                        for v in preview_b[header_b - 1][1]
                        if v is not None
                    ]

                # 字段映射
                field_mapping = self.ask_field_mapping(
                    headers_a,
                    headers_b,
                    f"工作簿：{sheet_a}",
                    f"工作簿：{sheet_b}",
                )
                if not field_mapping:
                    self.log("已取消字段映射")
                    break

                self.log(f"字段映射：{field_mapping}")
                self.log(f"关键列：A「{key_a}」 ↔ B「{key_b}」")
                self.log("正在对比...")

                # 执行对比
                try:
                    highlight_cells, compare_rows, diff_count, unmatched_keys, total_info = process_mode_b_compare(
                        file_path,
                        sheet_a,
                        header_a,
                        sheet_b,
                        header_b,
                        key_a,
                        key_b,
                        field_mapping,
                    )
                except Exception as e:
                    messagebox.showerror("处理错误", str(e), parent=self.root)
                    self.log(f"[错误] {e}")
                    break

                # 输出匹配结果日志
                self._log_compare_result(total_info, compare_rows, diff_count, unmatched_keys)

                # 应用标红到 wb_out 的工作簿 B
                apply_highlight_to_cells(wb_out, sheet_b, highlight_cells)
                self.log(f"已对工作簿【{sheet_b}】应用标红")

                # 询问是否继续
                if not messagebox.askyesno(
                    "继续匹配", "是否继续匹配另一对工作簿？", parent=self.root
                ):
                    break

            # 全部完成后保存
            if wb_out is not None:
                self._save_workbook(wb_out, current_file)
                wb_out = None

        except Exception as e:
            if wb_out is not None:
                wb_out.close()
            messagebox.showerror("错误", str(e), parent=self.root)
            self.log(f"[错误] {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SalaryMatchApp(root)
    root.mainloop()
