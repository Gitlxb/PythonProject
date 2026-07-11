import logging
import os
import re
import shutil
import tkinter as tk
from copy import copy
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

# 配置日志（仅控制台输出，不生成日志文件）
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


class SplitExcelApp:
    """
    增强版 Excel 表格拆分工具。
    支持嵌入模式：传入 parent 参数时，UI 构建在 parent 上；
    不传 parent 时，自动创建 tk.Tk 顶级窗口（兼容直接运行）。
    """
    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
            self._is_top_level = True
        else:
            self.root = parent.winfo_toplevel()
            self._is_top_level = False

        self.parent = parent

        self.filepath = ""
        self.wb = None
        self.ws = None
        self.sheet_names = []
        self.headers = []
        self.split_col_idx = None  # 1-based
        self.footer_start = None   # 1-based
        self.merge_map = {}
        self.dept_rows = {}
        self.dept_order = []
        self.custom_groups = []   # [{"name": "组合名", "values": [dept_raw, ...]}, ...]
        self.footer_confirmed = False  # 标记表尾是否已通过"选择拆分"确认
        self.output_dir = ""

        self._build_ui()

        if self._is_top_level:
            self.root.title("Excel表格拆分工具（增强版）")
            self.root.geometry("600x550")
            self.root.resizable(False, False)

    # ------------------------------------------------------------------ #
    #  UI
    # ------------------------------------------------------------------ #
    def _build_ui(self):
        padx = 10
        pady = 5

        # 确定 UI 挂载的父容器
        container = self.parent if self.parent is not None else self.root
        # 嵌入模式下也需要配置容器权重，使日志区域自动扩展
        container.columnconfigure(1, weight=1)
        container.rowconfigure(4, weight=1)

        # 文件选择
        tk.Label(container, text="Excel文件:").grid(row=0, column=0, sticky="e", padx=padx, pady=pady)
        self.entry_file = tk.Entry(container, width=45)
        self.entry_file.grid(row=0, column=1, sticky="ew", padx=padx, pady=pady)
        tk.Button(container, text="浏览", command=self._choose_file).grid(row=0, column=2, padx=padx, pady=pady)

        # Sheet选择
        tk.Label(container, text="工作表:").grid(row=1, column=0, sticky="e", padx=padx, pady=pady)
        self.combo_sheet = ttk.Combobox(container, width=42, state="readonly")
        self.combo_sheet.grid(row=1, column=1, sticky="ew", padx=padx, pady=pady)
        self.combo_sheet.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        # 拆分列选择
        tk.Label(container, text="拆分依据:").grid(row=2, column=0, sticky="e", padx=padx, pady=pady)
        self.combo_col = ttk.Combobox(container, width=42, state="readonly")
        self.combo_col.grid(row=2, column=1, sticky="ew", padx=padx, pady=pady)

        # 操作按钮（居中）
        btn_frame = tk.Frame(container)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=15)
        self.btn_choose_group = tk.Button(btn_frame, text="选择拆分", command=self._choose_split_groups, width=15)
        self.btn_choose_group.pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="开始拆分", command=self._start_split, width=15).pack(side=tk.LEFT, padx=5)
        self.btn_open_dir = tk.Button(btn_frame, text="打开输出目录", command=self._open_output_dir, state="disabled", width=15)
        self.btn_open_dir.pack(side=tk.LEFT, padx=5)

        # 状态/日志（自动填充剩余高度）
        tk.Label(container, text="日志:").grid(row=4, column=0, sticky="ne", padx=padx, pady=pady)
        self.text_log = tk.Text(container, width=65, height=1, state="disabled")
        self.text_log.grid(row=4, column=1, columnspan=2, padx=padx, pady=pady, sticky="nsew")

    def _log(self, msg: str):
        self.text_log.config(state="normal")
        self.text_log.insert(tk.END, msg + "\n")
        self.text_log.see(tk.END)
        self.text_log.config(state="disabled")
        if self.parent is not None:
            self.parent.update_idletasks()
        else:
            self.root.update_idletasks()
        logger.info(msg)

    # ------------------------------------------------------------------ #
    #  文件 / 工作表选择
    # ------------------------------------------------------------------ #
    def _choose_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            self.root.lift()
            self.root.focus_force()
            return
        self.filepath = path
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, path)

        try:
            self.wb = load_workbook(path, data_only=False)
            self.sheet_names = self.wb.sheetnames
            self.combo_sheet["values"] = self.sheet_names
            self.custom_groups = []
            self.footer_confirmed = False
            if self.sheet_names:
                self.combo_sheet.current(0)
                self._on_sheet_selected(None)
            self._log(f"已加载文件: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法加载文件:\n{e}")
        finally:
            self.root.lift()
            self.root.focus_force()

    def _on_sheet_selected(self, _):
        if not self.wb:
            return
        sheet_name = self.combo_sheet.get()
        if not sheet_name:
            return
        self.ws = self.wb[sheet_name]
        self.custom_groups = []
        self.footer_confirmed = False

        headers = []
        for idx, cell in enumerate(self.ws[2], 1):
            val = cell.value if cell.value is not None else ""
            headers.append((idx, str(val).strip()))
        self.headers = headers
        self.combo_col["values"] = [f"{col_idx}. {name}" for col_idx, name in headers]
        if headers:
            self.combo_col.current(0)
        self._log(f"已选择工作表: {sheet_name}, 第2行共 {len(headers)} 个表头")

    # ------------------------------------------------------------------ #
    #  选择拆分（自定义组合）
    # ------------------------------------------------------------------ #
    def _choose_split_groups(self):
        """弹出自定义组合拆分对话框（使用 SplitGroupDialog）"""
        if not self.filepath or not self.ws:
            messagebox.showwarning("提示", "请先选择Excel文件和工作表")
            return
        col_text = self.combo_col.get()
        if not col_text:
            messagebox.showwarning("提示", "请先选择拆分依据列")
            return

        self.split_col_idx = int(col_text.split(".")[0])

        # 检测表尾
        suggested_footer = self._detect_footer(self.ws, self.split_col_idx)
        max_row = self.ws.max_row
        if suggested_footer > max_row:
            self.footer_start = max_row + 1
        else:
            confirm = messagebox.askyesno(
                "确认表尾",
                f"检测到数据区域: 第3行 ~ 第{suggested_footer - 1}行\n"
                f"表尾起始行: 第{suggested_footer}行 (共 {max_row - suggested_footer + 1} 行)\n\n"
                f"是否确认？"
            )
            if not confirm:
                return
            self.footer_start = suggested_footer

        self.footer_confirmed = True

        # 构建部门列表（用于对话框展示）
        merge_map = self._build_merge_map(self.ws, self.split_col_idx)
        dept_order_tmp = []
        seen = set()
        for r in range(3, self.footer_start):
            val = merge_map.get(r)
            if val is None:
                val = self.ws.cell(r, self.split_col_idx).value
            if val is not None:
                dept = str(val).strip()
                if dept and dept not in seen:
                    seen.add(dept)
                    dept_order_tmp.append(dept)

        if not dept_order_tmp:
            messagebox.showwarning("提示", "未找到有效的拆分值")
            return

        # 使用 SplitGroupDialog 弹窗
        from hr.cf_gzbcf_hyp_simple import SplitGroupDialog

        dialog = SplitGroupDialog(self.root, dept_order_tmp)
        self.root.wait_window(dialog)

        if dialog.result is not None:
            # 将 SplitGroupDialog 的结果 (list of lists) 转换为 custom_groups 格式
            self.custom_groups = []
            for i, members in enumerate(dialog.result):
                self.custom_groups.append({
                    "name": f"组合{i + 1}",
                    "values": members,
                })
            group_desc = []
            for g in self.custom_groups:
                group_desc.append(f"{g['name']}: {', '.join(g['values'])}")
            self._log(f"已设置组合: {'; '.join(group_desc)}")
            # 确认后直接开始拆分
            self._start_split()

    # ------------------------------------------------------------------ #
    #  开始拆分（主逻辑）
    # ------------------------------------------------------------------ #
    def _start_split(self):
        if not self.filepath or not self.ws:
            messagebox.showwarning("提示", "请先选择Excel文件和工作表")
            return

        self.btn_open_dir.config(state="disabled")
        self.output_dir = ""

        col_text = self.combo_col.get()
        if not col_text:
            messagebox.showwarning("提示", "请选择拆分依据列")
            return

        self.split_col_idx = int(col_text.split(".")[0])
        max_row = self.ws.max_row

        # 若来自"选择拆分"弹窗，表尾已确认，跳过检测
        if not self.footer_confirmed:
            suggested_footer = self._detect_footer(self.ws, self.split_col_idx)
            if suggested_footer > max_row:
                messagebox.showinfo("提示", "未检测到表尾，所有行将被视为数据行")
                self.footer_start = max_row + 1
            else:
                confirm = messagebox.askyesno(
                    "确认表尾",
                    f"检测到数据区域: 第3行 ~ 第{suggested_footer - 1}行\n"
                    f"表尾起始行: 第{suggested_footer}行 (共 {max_row - suggested_footer + 1} 行)\n\n"
                    f"是否确认并开始拆分？"
                )
                if not confirm:
                    return
                self.footer_start = suggested_footer

        self.footer_confirmed = False

        # 构建部门映射
        self.merge_map = self._build_merge_map(self.ws, self.split_col_idx)
        self.dept_rows = {}
        self.dept_order = []
        seen = set()
        for r in range(3, self.footer_start):
            val = self.merge_map.get(r)
            if val is None:
                val = self.ws.cell(r, self.split_col_idx).value
            if val is not None:
                dept = str(val).strip()
                if dept and dept not in seen:
                    seen.add(dept)
                    self.dept_order.append(dept)
                    self.dept_rows[dept] = []
                if dept in self.dept_rows:
                    self.dept_rows[dept].append(r)

        if not self.dept_order:
            messagebox.showwarning("提示", "未找到有效的拆分值")
            return

        self._log(f"检测到 {len(self.dept_order)} 个拆分对象: {', '.join(self.dept_order)}")
        if self.custom_groups:
            self._log(f"自定义组合: {len(self.custom_groups)} 个")

        def _sanitize_filename(name: str) -> str:
            name = name.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()
            for ch in '\\/:*?"<>|':
                name = name.replace(ch, '_')
            return name

        # 预先计算实际有内容的表尾行
        actual_footer_rows = []
        for r in range(self.footer_start, max_row + 1):
            if any(self.ws.cell(r, c).value is not None and str(self.ws.cell(r, c).value).strip()
                   for c in range(1, self.ws.max_column + 1)):
                actual_footer_rows.append(r)

        self._log(f"  表尾有效行数: {len(actual_footer_rows)}")

        # 准备输出目录（若已存在则自动加序号后缀）
        base_dir = os.path.dirname(self.filepath)
        base_name = os.path.splitext(os.path.basename(self.filepath))[0]
        out_base = os.path.join(base_dir, f"{base_name}_拆分结果")
        self.output_dir = out_base
        suffix = 1
        while os.path.exists(self.output_dir):
            self.output_dir = f"{out_base}_{suffix}"
            suffix += 1
        os.makedirs(self.output_dir)

        old_last_data_row = self.footer_start - 1
        max_col = self.ws.max_column
        original_sheet_name = self.ws.title

        # ---- 构建任务列表 ----
        grouped_depts = set()
        tasks = []
        for g in self.custom_groups:
            tasks.append({"type": "group", "name": g["name"], "values": g["values"]})
            grouped_depts.update(g["values"])

        for dept in self.dept_order:
            if dept not in grouped_depts:
                tasks.append({"type": "single", "name": dept, "values": [dept]})

        # ---- 执行任务 ----
        for task in tasks:
            self._generate_file(
                task["values"], task["name"],
                old_last_data_row, max_col, actual_footer_rows,
                original_sheet_name, base_name, _sanitize_filename
            )

        self._log("拆分完成!")
        messagebox.showinfo("完成", f"拆分完成，共生成 {len(tasks)} 个文件，保存在:\n{self.output_dir}")
        self.btn_open_dir.config(state="normal")

    # ------------------------------------------------------------------ #
    #  生成单个输出文件（供任务和单独部门共用）
    # ------------------------------------------------------------------ #
    def _generate_file(self, dept_list, output_name,
                       old_last_data_row, max_col, actual_footer_rows,
                       original_sheet_name, base_name, sanitize_fn):
        """为一组部门（或单个部门）生成一个输出文件"""
        logger.info(f"{'='*60}")
        logger.info(f"开始生成文件: {output_name}")
        logger.info(f"  部门列表: {dept_list}")
        logger.info(f"  原工作表: {original_sheet_name}")
        logger.info(f"  数据行范围: 3 ~ {old_last_data_row}")

        # 合并所有行的索引
        all_rows = []
        for d in dept_list:
            rows = self.dept_rows.get(d, [])
            # 过滤幽灵行
            filtered = []
            for old_r in rows:
                if any(self.ws.cell(old_r, c).value is not None
                       and str(self.ws.cell(old_r, c).value).strip()
                       for c in range(1, max_col + 1) if c != self.split_col_idx):
                    filtered.append(old_r)
            all_rows.extend(filtered)
            logger.debug(f"  部门 '{d}' 的原始行号: {rows}")
            logger.debug(f"  部门 '{d}' 过滤后的行号: {filtered}")

        if not all_rows:
            self._log(f"  跳过 {output_name}: 无有效数据行")
            logger.warning(f"跳过 {output_name}: 无有效数据行")
            return

        all_rows.sort()
        keep_rows = [1, 2] + all_rows + actual_footer_rows
        keep_set = set(keep_rows)
        row_map = {old_r: new_r for new_r, old_r in enumerate(keep_rows, 1)}
        new_data_last_row = row_map[all_rows[-1]]
        old_first_data_row = all_rows[0]
        new_first_data_row = row_map[old_first_data_row]

        logger.info(f"  保留的行号: {keep_rows}")
        logger.info(f"  行号映射关系: {row_map}")
        logger.info(f"  新文件数据最后行: {new_data_last_row}")

        safe_name = sanitize_fn(output_name)
        new_filename = f"{base_name}_{safe_name}.xlsx"
        new_filepath = os.path.join(self.output_dir, new_filename)

        if os.path.exists(new_filepath):
            os.remove(new_filepath)
        shutil.copy(self.filepath, new_filepath)
        logger.info(f"  已创建新文件: {new_filepath}")

        # === 读取原文件：公式版（直接用已加载的self.ws）+ 计算值版 ===
        wb_src_value = load_workbook(self.filepath, data_only=True)    # 原文件（含计算值）
        ws_src_value = wb_src_value[original_sheet_name]
        logger.info(f"  已加载原文件计算值版")

        # === 新文件：复制后写入 ===
        wb_out = load_workbook(new_filepath, data_only=False)
        ws_out_old = wb_out[original_sheet_name]

        # 创建新Sheet（临时名称）
        temp_name = self._get_safe_sheet_name(wb_out, "__temp_split__")
        ws_new = wb_out.create_sheet(title=temp_name, index=0)
        logger.info(f"  已创建新Sheet: {temp_name}")

        # ---- 逐行复制：公式（来自self.ws）+ 计算值（来自wb_src_value） ----
        dispimg_count = 0
        formula_count = 0
        value_count = 0
        # 存储公式及其计算值，避免_value hack导致公式丢失
        formula_cells = {}  # {(new_r, c): (formula, calc_value)}

        for new_r, old_r in enumerate(keep_rows, 1):
            for c in range(1, max_col + 1):
                # 从self.ws读取公式（程序已加载，一定含公式）
                cell_formula = self.ws.cell(old_r, c)
                # 从wb_src_value读取计算值
                cell_value = ws_src_value.cell(old_r, c)
                new_cell = ws_new.cell(new_r, c)

                if isinstance(cell_formula.value, str) and cell_formula.value.startswith("="):
                    formula_count += 1
                    # DISPIMG 绑定原Sheet图片，删除旧Sheet后必#REF!，直接清空
                    if "DISPIMG" in cell_formula.value:
                        new_cell.value = None
                        dispimg_count += 1
                        logger.debug(f"    行{new_r}列{c}: 清除DISPIMG公式")
                    else:
                        # 保留原公式文本
                        original_formula = cell_formula.value
                        new_cell.value = original_formula
                        formula_cells[(new_r, c)] = (original_formula, cell_value.value)
                        logger.debug(f"    行{new_r}列{c}: 公式='{original_formula}' (计算值将由Excel自动计算)")
                else:
                    value_count += 1
                    new_cell.value = cell_formula.value
                    if cell_formula.value is not None and str(cell_formula.value).strip():
                        logger.debug(f"    行{new_r}列{c}: 值='{cell_formula.value}'")

                self._copy_cell_style(cell_formula, new_cell)

            # 复制行高
            if old_r in self.ws.row_dimensions:
                ws_new.row_dimensions[new_r].height = self.ws.row_dimensions[old_r].height

        logger.info(f"  单元格复制完成: 公式单元格={formula_count}, 值单元格={value_count}")
        logger.info(f"  已缓存 {len(formula_cells)} 个公式单元格供后续调整")

        # ---- 复制列宽 ----
        for col_letter in self.ws.column_dimensions:
            ws_new.column_dimensions[col_letter].width = self.ws.column_dimensions[col_letter].width
        logger.info(f"  列宽复制完成")

        # ---- 重建合并单元格 ----
        merged_count = 0
        for merged_range in self.ws.merged_cells.ranges:
            mc_min_col, mc_min_row, mc_max_col, mc_max_row = merged_range.bounds
            if set(range(mc_min_row, mc_max_row + 1)).issubset(keep_set):
                ws_new.merge_cells(
                    start_row=row_map[mc_min_row], start_column=mc_min_col,
                    end_row=row_map[mc_max_row], end_column=mc_max_col
                )
                merged_count += 1
                logger.debug(f"    重建合并单元格: {merged_range.coord} -> 新行{row_map[mc_min_row]}:{row_map[mc_max_row]}")
        logger.info(f"  合并单元格重建完成: {merged_count}个")

        # ---- 调整公式行号 ----
        actual_max_row = len(keep_rows)
        adjusted_formula_count = 0
        # 使用缓存的公式进行调整
        for (new_r, c), (original_formula, calc_value) in formula_cells.items():
            new_formula = self._adjust_formula(
                original_formula, row_map, old_last_data_row, new_data_last_row, actual_max_row,
                old_first_data_row, new_first_data_row
            )
            cell = ws_new.cell(new_r, c)
            cell.value = new_formula
            adjusted_formula_count += 1
            if original_formula != new_formula:
                logger.debug(f"    调整公式 行{new_r}列{c}: '{original_formula}' -> '{new_formula}'")
            # 不再使用_value hack，让Excel打开时自动计算

        # 额外检查：扫描所有单元格，看是否有遗漏的公式
        extra_formula_count = 0
        for r in range(1, actual_max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws_new.cell(r, c)
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    # 检查是否已经在formula_cells中
                    if (r, c) not in formula_cells:
                        # 这是一个遗漏的公式，需要调整
                        new_formula = self._adjust_formula(
                            cell.value, row_map, old_last_data_row, new_data_last_row, actual_max_row,
                            old_first_data_row, new_first_data_row
                        )
                        cell.value = new_formula
                        extra_formula_count += 1
                        logger.warning(f"    发现遗漏公式 行{r}列{c}: '{cell.value}' -> '{new_formula}'")

        logger.info(f"  公式调整完成: 缓存调整={adjusted_formula_count}, 额外发现={extra_formula_count}")
        logger.info(f"  注意: 拆分后的文件包含公式，Excel打开时会自动计算显示正确值")

        # ---- 步骤5.5：填充固定公式模板（个税列 + 社保单位扣款列） ----
        # 构建列名->列索引映射
        col_name_to_indices = {}
        for c in range(1, max_col + 1):
            val = ws_new.cell(2, c).value
            if val is not None:
                name = str(val).strip()
                if name not in col_name_to_indices:
                    col_name_to_indices[name] = []
                col_name_to_indices[name].append(c)

        # 5.5.1 个税列强制填充
        tax_filled = 0
        if "个税" in col_name_to_indices:
            tax_col = col_name_to_indices["个税"][0]  # 取第一个
            for r in range(3, new_data_last_row + 1):
                formula = f"=-ROUND(MAX((AK{r}-5000-AJ{r}-AT{r}+AM{r}+AN{r})*{{3,10,20,25,30,35,45}}%-5*{{0,42,282,532,882,1432,3032}},),2)"
                ws_new.cell(r, tax_col).value = formula
                tax_filled += 1
            logger.info(f"  个税公式填充: {tax_filled} 个单元格 (列{tax_col})")

        # ---- 步骤5.6：社保处理（根据个人扣款填充AS列单位扣款） ----
        ss_personal_cols = col_name_to_indices.get("社保个人扣款", [])
        ss_company_cols = col_name_to_indices.get("社保单位扣款", [])
        ss_filled = 0

        if ss_personal_cols and ss_company_cols:
            ss_personal_col = ss_personal_cols[0]          # 社保个人扣款列（AM）
            ss_company_as_col = ss_company_cols[-1]        # AS列（右边那列，取最后一个）

            for r in range(3, new_data_last_row + 1):
                # 检测社保个人扣款是否有数据（包括0，只要不为None）
                personal_val = ws_new.cell(r, ss_personal_col).value
                has_data = personal_val is not None  # 0也算有数据

                if has_data:
                    # 在AS列填充单位扣款公式
                    formula_company = f"=XLOOKUP(E{r},社保!$A$2:$A$110,-社保!$C$2:$C$110,0)"
                    ws_new.cell(r, ss_company_as_col).value = formula_company
                    # 在个人扣款列填充公式（覆盖原值）
                    formula_personal = f"=XLOOKUP(E{r},社保!$A$2:$A$110,-社保!$D$2:$D$110,0)"
                    ws_new.cell(r, ss_personal_col).value = formula_personal
                    ss_filled += 1
                    logger.debug(f"    行{r}: 个人扣款有数据 → AS列+个人扣款列均填充公式")

            logger.info(f"  社保填充: {ss_filled} 行 (根据个人扣款列{ss_personal_col} → AS列{ss_company_as_col}+个人扣款列均填充公式)")
        else:
            logger.info(f"  社保AS列填充: 跳过 (个人扣款列={len(ss_personal_cols)}列, 单位扣款列={len(ss_company_cols)}列)")

        # ---- 步骤5.7：公积金公式回填（原表哪里用公式，拆分后哪里填充） ----
        # 先打印所有含"公积金"的列名，诊断匹配情况
        gjj_headers_found = {}
        for c in range(1, max_col + 1):
            header_val = self.ws.cell(2, c).value
            if header_val is not None:
                name = str(header_val).strip()
                if "公积金" in name:
                    gjj_headers_found[c] = repr(name)
        logger.info(f"  公积金列检测: 找到{len(gjj_headers_found)}列 -> {gjj_headers_found}")

        # 扫描原表公积金列，记录有公式的行号和公式文本
        # 注意：原表可能使用动态数组公式（ArrayFormula），需要同时处理 str 和 ArrayFormula
        from openpyxl.worksheet.formula import ArrayFormula
        gjj_formula_records = []  # [(col_idx, old_row, formula), ...]
        logger.info(f"  公积金扫描范围: 行3 ~ 行{self.footer_start-1} (footer_start={self.footer_start})")
        for c in range(1, max_col + 1):
            header_val = self.ws.cell(2, c).value
            if header_val is not None:
                name = str(header_val).strip()
                if name in ("公积金扣款-单位", "公积金扣款-个人"):
                    found_in_col = 0
                    for r in range(3, self.footer_start):
                        cell_val = self.ws.cell(r, c).value
                        # 同时处理普通公式(str)和动态数组公式(ArrayFormula)
                        if isinstance(cell_val, ArrayFormula):
                            # ArrayFormula.text 才是真正的公式文本，str()返回的是对象repr
                            formula_text = cell_val.text or ""
                            # 去除 _xlfn. 前缀，否则 Excel 无法识别函数名
                            formula_text = formula_text.replace("=_xlfn.", "=")
                        elif isinstance(cell_val, str) and cell_val.startswith("="):
                            formula_text = cell_val.replace("=_xlfn.", "=")
                        else:
                            formula_text = ""
                        if formula_text:
                            gjj_formula_records.append((c, r, formula_text))
                            found_in_col += 1
                    logger.info(f"  公积金列{c}扫描: 找到{found_in_col}个公式")

        gjj_filled = 0
        for col_idx, old_r, old_formula in gjj_formula_records:
            if old_r in row_map:
                new_r = row_map[old_r]
                new_formula = self._adjust_formula(
                    old_formula, row_map, old_last_data_row, new_data_last_row, actual_max_row,
                    old_first_data_row, new_first_data_row
                )
                ws_new.cell(new_r, col_idx).value = new_formula
                gjj_filled += 1
                if old_formula != new_formula:
                    logger.debug(f"  公积金回填: 列{col_idx} 新行{new_r} '{old_formula}' -> '{new_formula}'")

        logger.info(f"  公积金公式回填: 扫描到{len(gjj_formula_records)}个, 实际填充{gjj_filled}个")

        # ---- 清理并保存 ----
        # 删除旧Sheet前，先清理其上的所有图片对象，避免删除后残留 #REF!
        img_count = 0
        if hasattr(ws_out_old, '_images'):
            img_count = len(ws_out_old._images)
            ws_out_old._images.clear()
        if img_count:
            logger.info(f"  已清理旧Sheet上的 {img_count} 个图片对象")
        wb_out.remove(ws_out_old)       # 删除旧Sheet
        logger.info(f"  已删除旧Sheet: {original_sheet_name}")
        wb_src_value.close()             # 关闭原文件（计算值版）

        # 将新Sheet重命名为原工作簿名称
        final_sheet_name = self._get_safe_sheet_name(wb_out, original_sheet_name)
        ws_new.title = final_sheet_name
        logger.info(f"  新Sheet重命名为: {final_sheet_name}")

        for sheet_name in ["各公司支出", "未发工资的社保扣除人员"]:
            if sheet_name in wb_out.sheetnames:
                wb_out.remove(wb_out[sheet_name])
                logger.info(f"  已删除额外Sheet: {sheet_name}")

        wb_out.save(new_filepath)
        wb_out.close()
        logger.info(f"  文件已保存: {new_filepath}")

        self._log(f"已生成: {new_filename} (数据行 {len(all_rows)} 条, 来源: {', '.join(dept_list)})")
        logger.info(f"文件生成完成: {new_filename}")
        logger.info(f"{'='*60}\n")

    # ------------------------------------------------------------------ #
    #  打开输出目录
    # ------------------------------------------------------------------ #
    def _open_output_dir(self):
        if self.output_dir and os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showwarning("提示", "输出目录不存在")

    # ------------------------------------------------------------------ #
    #  工具方法（合并单元格 / 表尾检测 / 样式复制 / 公式调整 / Sheet命名）
    # ------------------------------------------------------------------ #
    def _build_merge_map(self, ws, col_idx: int):
        merge_map = {}
        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            if min_col <= col_idx <= max_col:
                val = ws.cell(min_row, min_col).value
                for r in range(min_row, max_row + 1):
                    merge_map[r] = val
        return merge_map

    def _detect_footer(self, ws, col_idx: int):
        merge_map = self._build_merge_map(ws, col_idx)
        max_row = ws.max_row
        last_data_row = 2
        for r in range(3, max_row + 1):
            val = merge_map.get(r)
            if val is None:
                val = ws.cell(r, col_idx).value
            if val is not None and str(val).strip():
                txt = str(val).strip()
                if txt not in ("合计", "总计", "汇总", "小计", "平均"):
                    last_data_row = r
        return last_data_row + 1

    def _copy_cell_style(self, src, dst):
        if not src.has_style:
            return
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

    def _adjust_formula(self, formula: str, row_map: dict,
                        old_last_data_row: int, new_last_data_row: int,
                        new_max_row: int,
                        old_first_data_row: int = 0, new_first_data_row: int = 0) -> str:
        """调整公式中的行号引用，保护跨 Sheet 引用不被误改。
        采用分段处理策略：将公式切分为「跨Sheet引用」和「本Sheet部分」，
        只对本Sheet部分做行号替换，彻底避免占位符污染。

        Args:
            old_first_data_row / new_first_data_row: 跨部门下界 → 映射到本部门首行
            new_max_row: 当行号超出原数据区时映射到此行
        """
        logger.debug(f"  [_adjust] 入参: old_last={old_last_data_row}, new_last={new_last_data_row}, "
                     f"new_max={new_max_row}, old_first={old_first_data_row}, new_first={new_first_data_row}")
        logger.debug(f"  [_adjust] 原始公式: '{formula}'")

        # DISPIMG 绑定原Sheet图片对象，删除旧Sheet后必#REF!，直接清空
        if "DISPIMG" in formula:
            logger.debug(f"  [_adjust] 检测到DISPIMG公式，返回空字符串")
            return ""

        # ---- 阶段1: 匹配跨 Sheet 引用，切分公式为段落 ----
        cross_pattern = re.compile(
            r"('[^']+'|[A-Za-z0-9_\u4e00-\u9fa5]+)!"
            r"(\$?[A-Za-z]{1,3}\$?\d+"
            r"(:\$?[A-Za-z]{1,3}\$?\d+)?)"
        )

        segments = []  # (is_cross: bool, text: str)
        pos = 0
        for match in cross_pattern.finditer(formula):
            if pos < match.start():
                segments.append((False, formula[pos:match.start()]))
            segments.append((True, match.group()))
            pos = match.end()
        if pos < len(formula):
            segments.append((False, formula[pos:]))

        logger.debug(f"  [_adjust] 公式切分为 {len(segments)} 段 "
                     f"(跨Sheet={sum(1 for s in segments if s[0])}, "
                     f"本Sheet={sum(1 for s in segments if not s[0])})")

        # ---- 阶段2: 只对本Sheet段做行号替换 ----
        local_pattern = re.compile(r'(\$?[A-Za-z]{1,3}\$?)(\d+)')

        def repl(m):
            col_part = m.group(1)
            row_num = int(m.group(2))
            if row_num in row_map:
                new_num = row_map[row_num]
                logger.debug(f"  [_adjust]     {col_part}{row_num} -> {col_part}{new_num} (映射)")
            elif row_num == old_last_data_row:
                new_num = new_last_data_row
                logger.debug(f"  [_adjust]     {col_part}{row_num} -> {col_part}{new_num} (边界调整)")
            elif row_num > old_last_data_row:
                new_num = new_max_row
                logger.debug(f"  [_adjust]     {col_part}{row_num} -> {col_part}{new_num} (超出数据区)")
            elif old_first_data_row > 0 and row_num < old_first_data_row:
                new_num = new_first_data_row
                logger.debug(f"  [_adjust]     {col_part}{row_num} -> {col_part}{new_num} (下界映射)")
            elif 3 <= row_num <= old_last_data_row and row_num not in row_map:
                # 其他部门的数据行（不在本部门保留行中，但在原数据范围内）
                new_num = new_first_data_row
                logger.debug(f"  [_adjust]     {col_part}{row_num} -> {col_part}{new_num} (其他部门数据行)")
            else:
                new_num = row_num
                logger.debug(f"  [_adjust]     {col_part}{row_num} -> {col_part}{new_num} (不变)")
            return f"{col_part}{new_num}"

        result_parts = []
        replaced = 0
        for is_cross, text in segments:
            if is_cross:
                result_parts.append(text)
            else:
                old_text = text
                new_text = local_pattern.sub(repl, text)
                if old_text != new_text:
                    replaced += 1
                result_parts.append(new_text)

        result = ''.join(result_parts)

        if result != formula:
            logger.info(f"  [_adjust] 已调整: {len(result_parts)}段 替换{replaced}处 → "
                        f"'{formula[:100]}...' -> '{result[:100]}...'")
        else:
            logger.debug(f"  [_adjust] 无需调整: {len(result_parts)}段")

        return result

    def _get_safe_sheet_name(self, wb, base_name: str) -> str:
        name = base_name
        idx = 1
        while name in wb.sheetnames:
            name = f"{base_name}_{idx}"
            idx += 1
        return name


# ------------------------------------------------------------------ #
#  main
# ------------------------------------------------------------------ #
def main():
    root = tk.Tk()
    app = SplitExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
