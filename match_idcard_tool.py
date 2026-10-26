#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
身份证号码匹配工具 —— 多字段组合匹配

功能：
  1. 用户分别选择「数据源表」(第一个Excel) 与「匹配表」(第二个Excel)；
  2. 以 (运营中心, 运营项目, 员工姓名) 为复合键，
     在数据源表中查找对应的身份证号码；
  3. 将身份证号码回填到匹配表的"身份证号码"列；
  4. 未匹配的行身份证号码列留空；
  5. 输出新 Excel，保留匹配表的所有列和行顺序。

字段映射：
  匹配表          -> 数据源表
  运营中心        -> 所属运营中心
  运营项目        -> 运营项目
  员工姓名        -> 姓名
  身份证号码(输出) <- 身份证号码
"""

import os
import re
import sys
import subprocess
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    _TkBase = tk.Tk
except ImportError:
    tk = None
    _TkBase = object

try:
    import openpyxl
except ImportError:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        import openpyxl
    except Exception:
        openpyxl = None


# ---------- 字段定义 ----------
SOURCE_FIELDS = ["姓名", "所属运营中心", "运营项目", "身份证号码"]
MATCH_FIELDS = ["运营中心", "运营项目", "员工姓名", "身份证号码"]

SOURCE_KEY_COLS = ["所属运营中心", "运营项目", "姓名"]
MATCH_KEY_COLS = ["运营中心", "运营项目", "员工姓名"]

OUTPUT_FIELD = "身份证号码"


# ---------- 工具函数 ----------
def normalize_header(h):
    if h is None:
        return ""
    return re.sub(r"\s+", "", str(h))


def norm_val(v):
    if v is None:
        return ""
    return re.sub(r"\s+", "", str(v))


# ---------- 读取数据源表 ----------
def load_source(path):
    """
    读取数据源表，构建复合键 -> 身份证号码的索引字典。
    返回 (index_dict, total_rows)
    """
    if openpyxl is None:
        raise RuntimeError("缺少依赖 openpyxl，请先执行: pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"文件为空: {os.path.basename(path)}")

    header = [normalize_header(h) for h in rows[0]]
    missing = [f for f in SOURCE_FIELDS if f not in header]
    if missing:
        raise ValueError(f"数据源表缺少字段: {', '.join(missing)}")

    col_idx = {f: header.index(f) for f in SOURCE_FIELDS}
    name_idx = col_idx["姓名"]
    center_idx = col_idx["所属运营中心"]
    project_idx = col_idx["运营项目"]
    idcard_idx = col_idx["身份证号码"]

    index = {}
    total = 0
    for r in rows[1:]:
        total += 1
        key = (norm_val(r[center_idx]), norm_val(r[project_idx]), norm_val(r[name_idx]))
        if key == ("", "", ""):
            continue
        idcard = r[idcard_idx]
        if idcard is not None and str(idcard).strip() != "":
            index[key] = str(idcard).strip()

    return index, total


# ---------- 读取匹配表 ----------
def load_match_table(path):
    """
    读取匹配表，返回 (headers_list, rows_list)。
    rows_list 中每行是一个 list，与 headers_list 对齐。
    """
    if openpyxl is None:
        raise RuntimeError("缺少依赖 openpyxl，请先执行: pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"文件为空: {os.path.basename(path)}")

    raw_headers = [str(h) if h is not None else "" for h in rows[0]]
    header_norm = [normalize_header(h) for h in raw_headers]

    missing = [f for f in MATCH_FIELDS if f not in header_norm]
    if missing:
        raise ValueError(f"匹配表缺少字段: {', '.join(missing)}")

    center_idx = header_norm.index("运营中心")
    project_idx = header_norm.index("运营项目")
    name_idx = header_norm.index("员工姓名")
    idcard_idx = header_norm.index("身份证号码")

    data_rows = []
    for r in rows[1:]:
        row = list(r)
        if len(row) < len(raw_headers):
            row.extend([""] * (len(raw_headers) - len(row)))
        data_rows.append(row)

    meta = {
        "headers": raw_headers,
        "center_idx": center_idx,
        "project_idx": project_idx,
        "name_idx": name_idx,
        "idcard_idx": idcard_idx,
    }
    return data_rows, meta


# ---------- 核心匹配 ----------
def process(source_index, match_rows, meta):
    """
    用匹配表每一行的复合键，在数据源索引中查找身份证号码。
    匹配成功：填入身份证号码列；失败：留空。
    """
    center_idx = meta["center_idx"]
    project_idx = meta["project_idx"]
    name_idx = meta["name_idx"]
    idcard_idx = meta["idcard_idx"]

    matched = 0
    unmatched = 0

    for row in match_rows:
        key = (
            norm_val(row[center_idx]),
            norm_val(row[project_idx]),
            norm_val(row[name_idx]),
        )
        if key == ("", "", ""):
            row[idcard_idx] = ""
            unmatched += 1
            continue

        idcard = source_index.get(key)
        if idcard:
            row[idcard_idx] = idcard
            matched += 1
        else:
            row[idcard_idx] = ""
            unmatched += 1

    return matched, unmatched


# ---------- 写出结果 ----------
def write_output(path, headers, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "匹配结果"
    ws.append(headers)
    for row in rows:
        ws.append(row)

    for i in range(1, len(headers) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 18

    wb.save(path)


# ---------- GUI ----------
class MatchIdCardApp(_TkBase):
    def __init__(self):
        super().__init__()
        self.title("身份证号码匹配工具")
        self.geometry("880x340")
        self.minsize(720, 300)

        self._source_full = ""
        self._match_full = ""
        self.source_path = tk.StringVar(value="（未选择文件）")
        self.match_path = tk.StringVar(value="（未选择文件）")
        self.status_var = tk.StringVar(value="● 就绪")
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=0, minsize=200)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=0, minsize=120)

        tk.Label(self, text="身份证号码匹配工具",
                 font=("Microsoft YaHei", 13, "bold"),
                 fg="#1a237e").grid(row=0, column=0, columnspan=3,
                                    sticky="w", padx=18, pady=(14, 0))
        tk.Label(self,
                 text="以 (运营中心 + 运营项目 + 员工姓名) 为复合键匹配身份证号码",
                 fg="#666666").grid(row=1, column=0, columnspan=3,
                                    sticky="w", padx=18, pady=(0, 10))

        group = tk.LabelFrame(self, text="  文件选择  ",
                              font=("Microsoft YaHei", 9),
                              fg="#1a237e", bd=1, relief="groove",
                              labelanchor="nw")
        group.grid(row=2, column=0, columnspan=3,
                   sticky="ew", padx=18, pady=4)
        group.columnconfigure(1, weight=1)

        def add_row(r, label_text, path_var, choose_cmd):
            tk.Label(group, text=label_text, anchor="w"
                     ).grid(row=r, column=0, sticky="w",
                            padx=(12, 6), pady=10)
            tk.Entry(group, textvariable=path_var, state="readonly",
                     readonlybackground="#fafafa",
                     font=("Consolas", 10)
                     ).grid(row=r, column=1, sticky="ew", padx=4, pady=10)
            tk.Button(group, text="选择文件…", width=12,
                      command=choose_cmd
                      ).grid(row=r, column=2, padx=(4, 12), pady=10)

        add_row(0, "数据源表（第一个Excel）",
                self.source_path, self.select_source)
        add_row(1, "匹配表（第二个Excel）",
                self.match_path, self.select_match)

        action = tk.Frame(self)
        action.grid(row=3, column=0, columnspan=3,
                    sticky="ew", padx=18, pady=(14, 6))
        action.columnconfigure(0, weight=1)
        tk.Button(action, text="开 始 处 理",
                  font=("Microsoft YaHei", 11, "bold"),
                  bg="#2e7d32", fg="white",
                  activebackground="#1b5e20", activeforeground="white",
                  relief="flat", cursor="hand2",
                  height=2, width=18,
                  command=self.run).grid(row=0, column=0, sticky="e")

        tk.Label(self, textvariable=self.status_var, anchor="w",
                 fg="#555555", font=("Microsoft YaHei", 9)
                 ).grid(row=4, column=0, columnspan=3,
                        sticky="ew", padx=18, pady=(6, 0))

        tk.Label(self,
                 text="说明：匹配键为 (运营中心 + 运营项目 + 员工姓名)；"
                      "未匹配行身份证号码列留空。",
                 anchor="w", fg="#888888"
                 ).grid(row=5, column=0, columnspan=3,
                        sticky="ew", padx=18, pady=(0, 12))

    def select_source(self):
        p = filedialog.askopenfilename(
            title="选择数据源表（第一个Excel）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        if p:
            self._source_full = p
            self.source_path.set("📄  " + os.path.basename(p))

    def select_match(self):
        p = filedialog.askopenfilename(
            title="选择匹配表（第二个Excel）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        if p:
            self._match_full = p
            self.match_path.set("📄  " + os.path.basename(p))

    def run(self):
        if not self._source_full or not self._match_full:
            messagebox.showwarning("提示", "请先选择两个 Excel 文件。")
            return

        self.status_var.set("● 读取数据源表...")
        self.update_idletasks()

        try:
            source_index, src_total = load_source(self._source_full)
        except Exception as e:
            self.status_var.set("● 读取失败")
            messagebox.showerror("读取失败", str(e))
            return

        self.status_var.set("● 读取匹配表...")
        self.update_idletasks()

        try:
            match_rows, meta = load_match_table(self._match_full)
        except Exception as e:
            self.status_var.set("● 读取失败")
            messagebox.showerror("读取失败", str(e))
            return

        self.status_var.set("● 匹配中...")
        self.update_idletasks()

        matched, unmatched = process(source_index, match_rows, meta)

        summary = (
            f"处理完成：\n"
            f"  数据源表索引数：{len(source_index)} 条有效记录\n"
            f"  匹配表总行数：{len(match_rows)} 条\n"
            f"  成功匹配：{matched} 条\n"
            f"  未匹配（留空）：{unmatched} 条"
        )
        self.status_var.set(f"● 已处理：成功 {matched} 条 / 未匹配 {unmatched} 条")
        messagebox.showinfo("处理结果", summary)

        save_p = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="身份证号码匹配结果.xlsx")
        if not save_p:
            return
        try:
            write_output(save_p, meta["headers"], match_rows)
        except Exception as e:
            self.status_var.set("● 保存失败")
            messagebox.showerror("保存失败", str(e))
            return
        self.status_var.set(f"● 已保存：{os.path.basename(save_p)}")
        messagebox.showinfo("完成", f"结果已保存至：\n{save_p}")


def main():
    if tk is None:
        sys.stderr.write(
            "缺少 tkinter，GUI 无法启动。请使用标准 Windows Python "
            "(通常自带 tkinter) 运行本程序。\n")
        return
    if openpyxl is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "缺少依赖",
            "未能导入 openpyxl。请先执行：\n\npip install openpyxl\n\n后重新运行本程序。")
        return
    app = MatchIdCardApp()
    app.mainloop()


if __name__ == "__main__":
    main()
