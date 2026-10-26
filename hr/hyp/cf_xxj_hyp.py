# -*- coding: utf-8 -*-
"""
孝心金拆分 — 核心逻辑层（纯 openpyxl，不含任何 UI 依赖）

职责
    1. 加载「孝心金」类 Excel，识别第 2 行表头，自动检测表尾区间
    2. 按指定拆分列把数据区切分为「段」——段边界由「拆分列合并块」+「值变化」共同决定
    3. 每一段（或用户组合后的若干段）输出为一个独立 Excel 文件，
       保留标题行 / 表头行 / 表尾（空行、金额、个数、汇总公式），列宽与单元格样式按源文件复制

表结构约定（与源表「26年第三季度孝心金.xlsx」一致）
    第 1 行    标题（A1:H1 合并），输出时原文照抄
    第 2 行    表头，拆分依据列下拉列表的来源
    第 3 行起  数据区（每行一位员工）
    空行之后   表尾：汇总标签行（H=金额 / I=个数）+ 汇总公式行（=SUM(...) / =H{行}/500）

与「工资表拆分」的差异
    - 不含个税 / 社保 / 公积金 / 通讯费等业务公式回填，汇总区只有 =SUM 与 =H{r}/500
    - 拆分单元是「拆分列的合并块」而非「拆分列的值」：同名但不连续的段分别输出为独立文件
      （例：源表「浙江锦途」出现在第 3-46 行与第 125-137 行，会生成 浙江锦途.xlsx 与 浙江锦途_2.xlsx）
    - 列宽一律按源文件复制
"""
import os
import re
import logging
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from numbers import Number

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------- 结构常量
TITLE_ROW = 1           # 标题行
HEADER_ROW = 2          # 表头区起始行（拆分依据列表头的来源）
DATA_START_ROW = 3      # 数据起始行的兜底值（实际由 _detect_data_start_row 动态识别）

TMP_SHEET_NAME = "__xxj_tmp__"
MAX_SHEET_NAME_LEN = 31

# 表尾标签行常见词（避免被误判为数据行）
SUMMARY_KEYWORDS = ("合计", "总计", "汇总", "小计", "平均", "金额", "个数", "人数")

# 金额列的候选表头名（用于预览区展示每段金额合计）
AMOUNT_HEADER_CANDIDATES = ("金额", "金额(元)", "金额（元）", "补贴金额", "孝心金")

_INVALID_FILENAME_CHARS = '\\/:*?"<>|'
_WIN_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_MAX_FILENAME_LEN = 120


def _display_width(text) -> int:
    """按终端显示宽度计算字符串长度（中日韩字符算 2 格）"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def _ljust_display(text, width: int) -> str:
    """按显示宽度左对齐补空格，让日志里的文件名列整齐"""
    s = str(text)
    return s + " " * max(width - _display_width(s), 1)


def sanitize_filename(name: str) -> str:
    """清理 Windows 文件名非法字符 / 保留名 / 结尾点号空格"""
    name = str(name)
    for ch in ("\r", "\n", "\t"):
        name = name.replace(ch, " ")
    for ch in _INVALID_FILENAME_CHARS:
        name = name.replace(ch, "_")
    name = name.strip().rstrip(". ")
    if not name:
        name = "未命名"
    if name.upper() in _WIN_RESERVED_NAMES:
        name += "_"
    if len(name) > _MAX_FILENAME_LEN:
        name = name[:_MAX_FILENAME_LEN].rstrip(". ")
    return name


@dataclass
class SplitSegment:
    """一个拆分分组：拆分列取值相同的所有数据行（同名不连续时也并入同一组）"""
    value: str
    rows: list = field(default_factory=list)
    amount: float = 0.0

    @property
    def start(self) -> int:
        return self.rows[0] if self.rows else 0

    @property
    def end(self) -> int:
        return self.rows[-1] if self.rows else 0

    @property
    def count(self) -> int:
        return len(self.rows)

    @property
    def spans(self):
        """把行号列表压成连续区间 [(起行, 止行), ...]（1-based）"""
        out = []
        for r in sorted(self.rows):
            if out and r == out[-1][1] + 1:
                out[-1][1] = r
            else:
                out.append([r, r])
        return [(a, b) for a, b in out]

    @property
    def range_text(self) -> str:
        """界面展示用的行范围文本：该组在源表中的首行 ~ 末行（同名值跨多处时也只显示一个范围）"""
        if not self.rows:
            return ""
        return f"第{self.start}-{self.end}行"

    @property
    def label(self) -> str:
        """用于界面展示：分组值 + 源文件行范围"""
        return f"{self.value}    {self.range_text}"


class XXJSplitter:
    """孝心金拆分核心逻辑（与 UI 解耦，可独立调用 / 测试）"""

    _CROSS_SHEET_RE = re.compile(
        r"('[^']+'|[A-Za-z0-9_\u4e00-\u9fa5]+)!"
        r"(\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?)"
    )
    _LOCAL_REF_RE = re.compile(r"(\$?[A-Za-z]{1,3}\$?)(\d+)")

    def __init__(self, log=None, debug=None):
        # log   -> 关键进度信息（界面日志区显示）
        # debug -> 细节信息（公式重排等，界面默认不显示）
        self.log = log if callable(log) else (lambda _msg: None)
        self.debug = debug if callable(debug) else (lambda _msg: None)

        self.filepath = None
        self.sheet_name = None
        self.wb = None
        self.ws = None

        self.headers = []            # [(col_idx, header_name), ...]  仅非空表头
        self.header_map = {}         # header_name -> col_idx
        self.split_col_idx = None    # 拆分依据列的列号（1-based）
        self.amount_col_idx = None   # 金额列列号（找不到为 None）

        self.data_start_row = DATA_START_ROW   # 数据起始行（加载后动态识别）
        self.header_rows = [HEADER_ROW]        # 表头区行号列表，如 [2, 3, 4]

        self.footer_start = None     # 表尾起始行
        self.footer_end = None       # 表尾最后一个非空行（< footer_start 表示无表尾）
        self.effective_max_col = None

        self.segments = []           # [SplitSegment, ...]

    # ================================================== 加载与解析
    def load(self, filepath: str, sheet_name: str = None):
        """加载工作簿与工作表，识别表头区并解析列名"""
        self.filepath = filepath
        self.wb = load_workbook(filepath, data_only=False)
        self.sheet_name = sheet_name or self.wb.sheetnames[0]
        if self.sheet_name not in self.wb.sheetnames:
            self.sheet_name = self.wb.sheetnames[0]
        self.ws = self.wb[self.sheet_name]

        # 先定位「表头区 / 数据起始行」，再解析列名（列名可能分布在表头区多行里）
        self.data_start_row = self._detect_data_start_row()
        self.header_rows = list(range(HEADER_ROW, self.data_start_row))

        self._parse_headers()
        self._resolve_amount_col()
        # 列号与表尾相关的状态在切换拆分列时重建
        self.split_col_idx = None
        self.footer_start = self.footer_end = None
        self.effective_max_col = None
        self.segments = []

        self.log(f"已加载：{os.path.basename(filepath)} / 工作表「{self.sheet_name}」")
        if len(self.header_rows) == 1:
            self.log(f"表头：第 {self.header_rows[0]} 行，数据从第 {self.data_start_row} 行开始")
        else:
            self.log(f"表头：第 {self.header_rows[0]}-{self.header_rows[-1]} 行（共 {len(self.header_rows)} 行），"
                     f"数据从第 {self.data_start_row} 行开始")
        self.log(f"识别到 {len(self.headers)} 个有效表头："
                 + "、".join(name for _, name in self.headers))
        return self

    def _detect_data_start_row(self) -> int:
        """动态识别数据起始行（兼容单行表头与多行表头两种格式）。

        判定：从第 3 行起，第一个满足「非空单元格 >= 3」且「含非文本值（日期/数值）」的行。
        表头行即便有值也只是文本或公式，因此不会被命中。
        """
        max_row = self.ws.max_row
        max_col = self.ws.max_column
        for r in range(HEADER_ROW + 1, max_row + 1):
            values = []
            for c in range(1, max_col + 1):
                v = self.ws.cell(r, c).value
                if v is not None and str(v).strip():
                    values.append(v)
            if len(values) < 3:
                continue
            if any(isinstance(v, (datetime, date))
                   or (isinstance(v, Number) and not isinstance(v, bool))
                   for v in values):
                return r
        return HEADER_ROW + 1   # 兜底：表头只有一行

    def _header_area_max_col(self) -> int:
        """表头区内最后一个非空列"""
        max_col = 0
        for r in self.header_rows:
            for c in range(1, self.ws.max_column + 1):
                v = self.ws.cell(r, c).value
                if v is not None and str(v).strip():
                    max_col = max(max_col, c)
        return max_col

    def _parse_headers(self):
        """解析表头区，逐列取「最后一个非空、非公式」的值作为列名。

        新版表的列名并非都在同一行：A~F 在第 2 行（靠跨行合并撑着），
        G/H 在第 4 行 —— 第 2/3 行的「总金额」「个数」及其公式都不算列名。
        """
        self.headers = []
        self.header_map = {}
        for idx in range(1, self._header_area_max_col() + 1):
            name = ""
            for r in self.header_rows:
                raw = self.ws.cell(r, idx).value
                if raw is None:
                    continue
                txt = str(raw).strip()
                if not txt or txt.startswith("="):
                    continue
                name = txt          # 持续覆盖，最终留下表头区里最靠下的那个文本
            if name:
                self.headers.append((idx, name))
                self.header_map.setdefault(name, idx)

    def _resolve_amount_col(self):
        self.amount_col_idx = None
        for name in AMOUNT_HEADER_CANDIDATES:
            if name in self.header_map:
                self.amount_col_idx = self.header_map[name]
                break

    # ================================================== 拆分准备
    def prepare(self, split_col_idx: int):
        """按拆分列解析表尾、有效列范围，并切分段"""
        if self.ws is None:
            raise RuntimeError("尚未加载文件")

        self.split_col_idx = int(split_col_idx)
        self._resolve_footer(self.split_col_idx)
        self._resolve_effective_max_col()
        self._build_segments(self.split_col_idx)

        info = f"数据区：第 {self.data_start_row}-{self.footer_start - 1} 行"
        if self.footer_end >= self.footer_start:
            info += f"，表尾：第 {self.footer_start}-{self.footer_end} 行"
        else:
            info += "，未检测到表尾"
        self.log(info)
        self.log(f"按「{self.header_name_of(self.split_col_idx)}」拆分，共识别 {len(self.segments)} 个分组")
        return self.segments

    def header_name_of(self, col_idx: int) -> str:
        for idx, name in self.headers:
            if idx == col_idx:
                return name
        return f"第{col_idx}列"

    # ---------------------------------------------- 表尾 / 列范围
    def _resolve_footer(self, col_idx: int):
        """定位表尾区间。

        数据区最后一行 = 拆分列上最后一个「有效值」所在行；
        有效值需排除：空值、公式、模板文字（合计/金额/个数等）、与表头同名的标签。
        表尾末行 = 表尾区间内最后一个非空行（保证分隔空行被一并保留）。
        """
        merge_map = self._build_merge_map(col_idx)
        header_names = {name for _, name in self.headers}

        last_data_row = self.header_rows[-1]
        for r in range(self.data_start_row, self.ws.max_row + 1):
            val = merge_map.get(r)
            if val is None:
                val = self.ws.cell(r, col_idx).value
            if val is None:
                continue
            txt = str(val).strip()
            if not txt or txt.startswith("=") or txt in SUMMARY_KEYWORDS or txt in header_names:
                continue
            last_data_row = r

        footer_start = last_data_row + 1

        footer_end = footer_start - 1
        for r in range(footer_start, self.ws.max_row + 1):
            if any(
                self.ws.cell(r, c).value is not None and str(self.ws.cell(r, c).value).strip()
                for c in range(1, self.ws.max_column + 1)
            ):
                footer_end = r

        self.footer_start = footer_start
        self.footer_end = footer_end

    def _resolve_effective_max_col(self):
        """有效列范围 = 表头区最后一个非空列，与表尾区间涉及的列取较大值。

        新版表头区在 G/H 列给出「总金额/个数/联系电话/金额」，有效列为 A~H；
        旧版表头只到 H、表尾的「个数」落在 I 列，此时取并集得到 A~I。
        """
        max_col = self._header_area_max_col() or self.ws.max_column
        if self.footer_end >= self.footer_start:
            for r in range(self.footer_start, min(self.footer_end, self.ws.max_row) + 1):
                for c in range(1, self.ws.max_column + 1):
                    v = self.ws.cell(r, c).value
                    if v is not None and str(v).strip():
                        max_col = max(max_col, c)
        self.effective_max_col = max_col

    # ---------------------------------------------- 合并块 / 分段
    def _build_merge_map(self, col_idx: int) -> dict:
        """返回 {行号: 该行在拆分列上的有效值}（合并块内非首行取首行值）"""
        merge_map = {}
        for mr in self.ws.merged_cells.ranges:
            if mr.min_col <= col_idx <= mr.max_col:
                val = self.ws.cell(mr.min_row, mr.min_col).value
                for r in range(mr.min_row, mr.max_row + 1):
                    merge_map[r] = val
        return merge_map

    def _build_segments(self, col_idx: int):
        """按拆分列的取值聚合分组。

        同名值无论出现在几个不连续的位置，都并入同一个分组（输出为同一个文件），
        分组顺序按该值在源表中首次出现的顺序。
        """
        merge_map = self._build_merge_map(col_idx)
        groups = {}      # 值 -> SplitSegment
        order = []       # 保持首次出现顺序
        last_val = ""

        for r in range(self.data_start_row, self.footer_start):
            val = merge_map.get(r)
            if val is None:
                val = self.ws.cell(r, col_idx).value
            txt = "" if val is None else str(val).strip()
            if not txt:
                txt = last_val          # 拆分列为空时沿用上一行的值
            if not txt:
                continue
            if txt not in groups:
                groups[txt] = SplitSegment(txt, [])
                order.append(txt)
            groups[txt].rows.append(r)
            last_val = txt

        segments = [groups[v] for v in order]
        for seg in segments:
            seg.rows.sort()
            seg.amount = self._sum_amount(seg.rows)

        self.segments = segments
        return segments

    def _sum_amount(self, rows) -> float:
        if not self.amount_col_idx:
            return 0.0
        total = 0.0
        for r in rows:
            v = self.ws.cell(r, self.amount_col_idx).value
            if isinstance(v, Number) and not isinstance(v, bool):
                total += float(v)
        return total

    # ================================================== 输出
    def make_output_dir(self) -> str:
        """在源文件同级目录下建立「{源名}_拆分结果」文件夹（重名自动加序号）"""
        base_dir = os.path.dirname(os.path.abspath(self.filepath))
        base_name = os.path.splitext(os.path.basename(self.filepath))[0]
        out = os.path.join(base_dir, f"{base_name}_拆分结果")
        idx = 1
        while os.path.exists(out):
            out = os.path.join(base_dir, f"{base_name}_拆分结果_{idx}")
            idx += 1
        os.makedirs(out)
        return out

    def split(self, output_dir: str = None, groups=None):
        """执行拆分。

        groups 为 None                -> 每个段各自输出一个文件（直接拆分）
        groups 为 [[seg, seg], ...]   -> 组内段合并为一个文件，未入组的段各自输出（组合拆分）

        返回 ([(输出路径, [SplitSegment, ...]), ...], 输出目录)
        """
        if not self.segments:
            raise RuntimeError("尚未准备拆分数据，请先调用 prepare()")

        output_dir = output_dir or self.make_output_dir()

        tasks = []
        if groups:
            assigned = set()
            for grp in groups:
                members = [s for s in grp if s is not None]
                if not members:
                    continue
                tasks.append(members)
                assigned.update(id(s) for s in members)
            for seg in self.segments:
                if id(seg) not in assigned:
                    tasks.append([seg])
        else:
            tasks = [[seg] for seg in self.segments]

        used_names = set()
        results = []
        total = len(tasks)
        for i, segs in enumerate(tasks, 1):
            base = self._compose_name(segs)
            filename = self._unique_filename(base, used_names)
            path = os.path.join(output_dir, filename)
            self._generate_one(segs, path)
            row_total = sum(s.count for s in segs)
            amount_total = sum(s.amount for s in segs)
            amount_txt = (f"{self._fmt_amount(amount_total):>9,} 元   " if self.amount_col_idx else "")
            source = "源 " + "、".join(f"第{s.start}-{s.end}行" for s in segs)
            self.log(f"[{i:>3}/{total}] {_ljust_display(filename, 24)}{row_total:>3} 行   "
                     f"{amount_txt}{source}")
            results.append((path, segs))

        self.log(f"完成：共生成 {len(results)} 个文件 → {output_dir}")
        return results, output_dir

    @staticmethod
    def _fmt_amount(v):
        """金额去掉无意义的小数尾巴"""
        v = float(v)
        return int(v) if v.is_integer() else round(v, 2)

    @staticmethod
    def _compose_name(segs) -> str:
        """组合文件名：成员值按顺序用下划线连接，连续重复值只保留一次"""
        uniq = []
        for s in segs:
            if not uniq or uniq[-1] != s.value:
                uniq.append(s.value)
        return sanitize_filename("_".join(uniq))

    @staticmethod
    def _unique_filename(base: str, used_names: set) -> str:
        """同名输出自动加序号：浙江锦途.xlsx -> 浙江锦途_2.xlsx"""
        name = base
        idx = 2
        while name.lower() in used_names:
            name = sanitize_filename(f"{base}_{idx}")
            idx += 1
        used_names.add(name.lower())
        return f"{name}.xlsx"

    # ---------------------------------------------- 单文件生成
    def _generate_one(self, segs, output_path: str):
        max_col = self.effective_max_col
        split_col = self.split_col_idx

        rows = []
        seen = set()
        for seg in segs:
            for r in seg.rows:
                if r not in seen:
                    seen.add(r)
                    rows.append(r)

        # 过滤幽灵行：拆分列以外的列全空则视为无效行
        rows = [
            r for r in rows
            if any(
                self.ws.cell(r, c).value is not None
                and str(self.ws.cell(r, c).value).strip()
                for c in range(1, max_col + 1)
                if c != split_col
            )
        ]
        if not rows:
            self.log(f"跳过（无有效数据行）：{os.path.basename(output_path)}")
            return
        rows.sort()

        footer_rows = (
            list(range(self.footer_start, self.footer_end + 1))
            if self.footer_end >= self.footer_start else []
        )
        # 标题行 + 完整表头区（单行表头=第2行，多行表头=第2~N行）+ 数据行 + 表尾
        keep_rows = list(range(TITLE_ROW, self.data_start_row)) + rows + footer_rows

        # 去重并保持顺序（容错）
        ordered, seen_row = [], set()
        for r in keep_rows:
            if r not in seen_row:
                seen_row.add(r)
                ordered.append(r)
        keep_rows = ordered

        row_map = {old: new for new, old in enumerate(keep_rows, 1)}

        old_last_data_row = self.footer_start - 1
        new_last_data_row = row_map[rows[-1]]
        new_max_row = len(keep_rows)
        old_first_data_row = rows[0]
        new_first_data_row = row_map[rows[0]]

        wb_out = load_workbook(self.filepath, data_only=False)
        ws_out = wb_out.create_sheet(title=TMP_SHEET_NAME, index=0)

        # ---- 逐行复制：值 + 公式 + 单元格样式 + 行高 ----
        for new_r, old_r in enumerate(keep_rows, 1):
            for c in range(1, max_col + 1):
                src = self.ws.cell(old_r, c)
                dst = ws_out.cell(new_r, c)
                val = src.value
                if isinstance(val, str) and val.startswith("="):
                    # DISPIMG 公式绑定原 Sheet 的图片对象，脱离后会变成 #REF!
                    dst.value = None if "DISPIMG" in val else val
                else:
                    dst.value = val
                self._copy_cell_style(src, dst)

            src_dim = self.ws.row_dimensions.get(old_r)
            if src_dim is not None and src_dim.height:
                ws_out.row_dimensions[new_r].height = src_dim.height

        # ---- 列宽：按源文件复制 ----
        for col_letter, dim in self.ws.column_dimensions.items():
            if dim.width:
                ws_out.column_dimensions[col_letter].width = dim.width

        # ---- 重建合并单元格（仅当整个合并区都在保留行内）----
        # 拆分列单独处理：该列在这一组的所有数据行上取值相同，源表里可能分成多块
        # （同名值在源表中不连续时），这里统一合并成一个整块，避免出现两个同名单元格。
        for mr in self.ws.merged_cells.ranges:
            if mr.min_col > max_col:
                continue
            # 只跳过「数据区、且仅覆盖拆分列」的合并块（交给下面整组统一合并）。
            # 判断必须同时限定单列 + 数据区：早期写成 min_col <= split_col <= max_col，
            # 会把 A1:H1 这种跨列标题合并也误跳过，导致输出文件标题塌进 A1。
            if (mr.min_col == mr.max_col == split_col
                    and mr.min_row >= self.data_start_row):
                continue
            if all(r in seen_row for r in range(mr.min_row, mr.max_row + 1)):
                ws_out.merge_cells(
                    start_row=row_map[mr.min_row], start_column=mr.min_col,
                    end_row=row_map[mr.max_row], end_column=mr.max_col,
                )

        # 拆分列：整组数据行合并为一个单元格（行在输出中连续时才合并）
        first_new, last_new = row_map[rows[0]], row_map[rows[-1]]
        if last_new > first_new and (last_new - first_new + 1) == len(rows):
            ws_out.merge_cells(
                start_row=first_new, start_column=split_col,
                end_row=last_new, end_column=split_col,
            )

        # ---- 汇总公式行号重排（=SUM(H3:H137) -> =SUM(H3:H15) 等）----
        for row in ws_out.iter_rows(min_row=1, max_row=new_max_row, max_col=max_col):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    if "DISPIMG" in v:
                        cell.value = None
                        continue
                    new_formula = self._adjust_formula(
                        v, row_map, old_last_data_row, new_last_data_row, new_max_row,
                        old_first_data_row, new_first_data_row,
                    )
                    if new_formula != v:
                        cell.value = new_formula
                        self.debug(f"公式重排 {cell.coordinate}：{v}  ->  {new_formula}")

        # ---- 只保留新 Sheet，并沿用源工作表名 ----
        for name in list(wb_out.sheetnames):
            if name != ws_out.title:
                wb_out.remove(wb_out[name])
        ws_out.title = self.sheet_name[:MAX_SHEET_NAME_LEN]

        wb_out.save(output_path)
        wb_out.close()

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
                        old_last_data_row: int, new_last_data_row: int, new_max_row: int,
                        old_first_data_row: int = 0, new_first_data_row: int = 0) -> str:
        """重排公式中的行号引用。

        先把公式切成「跨 Sheet 引用段」与「本 Sheet 段」，只对后者做行号替换，
        避免把 '员工信息-7月'!$C:$P 这类引用改坏。
        """
        segments = []
        pos = 0
        for m in self._CROSS_SHEET_RE.finditer(formula):
            if pos < m.start():
                segments.append((False, formula[pos:m.start()]))
            segments.append((True, m.group()))
            pos = m.end()
        if pos < len(formula):
            segments.append((False, formula[pos:]))

        def repl(m):
            col_part, row_num = m.group(1), int(m.group(2))
            if row_num in row_map:
                new_num = row_map[row_num]
            elif row_num == old_last_data_row:
                new_num = new_last_data_row
            elif row_num > old_last_data_row:
                new_num = new_max_row
            elif old_first_data_row and row_num < old_first_data_row:
                new_num = new_first_data_row
            elif self.data_start_row <= row_num <= old_last_data_row:
                # 引用了落在本段之外的其他分组数据行，收敛到本段首行
                new_num = new_first_data_row
            else:
                new_num = row_num
            return f"{col_part}{new_num}"

        return "".join(
            text if is_cross else self._LOCAL_REF_RE.sub(repl, text)
            for is_cross, text in segments
        )

    # ================================================== 预览
    def preview_rows(self):
        """界面预览数据：[{name, range, count, amount}, ...]"""
        out = []
        for seg in self.segments:
            out.append({
                "name": seg.value,
                "range": seg.range_text,
                "count": seg.count,
                "amount": self._fmt_amount(seg.amount) if self.amount_col_idx else "",
            })
        return out


# ==================================================================== 自测
def _smoke_test(src_path: str, col_idx: int = 1):
    """命令行冒烟测试：python cf_xxj_hyp.py <xlsx 路径> [拆分列号]"""
    print("=" * 72)
    splitter = XXJSplitter(log=print)
    splitter.load(src_path)
    segments = splitter.prepare(col_idx)
    print("-" * 72)
    for seg in segments:
        print(f"  {seg.value:<14} {seg.range_text:<16} {seg.count:>3} 行   金额 {seg.amount}")
    print("-" * 72)
    results, out_dir = splitter.split()
    print(f"输出目录：{out_dir}")
    for path, segs in results:
        print(f"  {os.path.basename(path)}")


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    default_src = r"C:\Users\jintu32\Desktop\26年第三季度孝心金.xlsx"
    src = sys.argv[1] if len(sys.argv) > 1 else default_src
    col = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    _smoke_test(src, col)
