# -*- coding: utf-8 -*-
"""
Word 文档合并工具 — 分类层级合并引擎（word_core_hyp2.py）

本模块是 Issue 2 的全新实现：
  1. 关键词提取（带 parent 一级标题）
     - 0-X 分数模式: 提取关键词 + 左侧合并单元格中的"类别名"
     - 主观回退: 提取关键词 + 左侧"类别名" + 顿号拆分
  2. 智能搜索（精准 → 模糊 两级）
     - 先精准匹配（文件名包含完整关键词）
     - 未命中则模糊匹配（4字分块组合）
  3. 分类合并
     - 按 parent 分组，插入"一、X"格式一级标题
     - 分类间插入分节符

依赖:
  - word_core_hyp1:  文件检测、安全打开、合并引擎（clone_element_to_doc,
                      copy_all_content, insert_section_break 等）
"""

import os
import re

from .word_core_hyp1 import (
    _safe_open_document,
    clone_element_to_doc,
    copy_all_content,
    insert_section_break,
    _build_style_map,
    _extract_base_hf_rids,
    _preload_image_rels,
    _merge_numbering,
)


# ============================================================
#  中文数字（序号用）
# ============================================================
_CHINESE_NUMS = [
    "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
    "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
]


# ============================================================
#  顿号拆分
# ============================================================

def _split_by_dunhao(text):
    """
    按顿号（、）拆分关键词。

    规则：
      1. 无顿号 → 原样返回 [text]
      2. 有顿号 → 按顿号拆分
      3. 最后一个元素若含逗号（，）→ 去掉逗号及之后的内容

    例如:
      "根据投标人的管理运作制度、人员福利激励机制、内部岗位责任制度、
       管理人员考核制度及标准、人员信息及工作情况档案建立与管理制度，进行综合打分"
      → ["根据投标人的管理运作制度", "人员福利激励机制",
         "内部岗位责任制度", "管理人员考核制度及标准",
         "人员信息及工作情况档案建立与管理制度"]

    返回:
        list[str]: 拆分后的子关键词列表
    """
    if '、' not in text:
        return [text.strip()]

    parts = text.split('、')
    result = []

    for i, part in enumerate(parts):
        part = part.strip()
        if not part:
            continue

        # 最后一个元素: 去掉逗号及之后的内容
        if i == len(parts) - 1 and '，' in part:
            comma_idx = part.index('，')
            part = part[:comma_idx].strip()

        if part:
            result.append(part)

    return result if result else [text.strip()]


# ============================================================
#  4字分块模糊搜索
# ============================================================

def _fuzzy_chunks_4char(keyword, min_chunk=4):
    """
    将关键词按4字一组分块，余数合并到前一块。

    规则:
      - 每4个字符一组
      - 最后一块不足4字 → 合并到前一块
      - 返回: [完整关键词, 4字块1, 4字块2, ...]

    例如:
      "售后服务能力" (6字)
      → chunk1: "售后" (2字) → 不足4, 继续
      → chunk1: "售后服务" (4字) ✓
      → chunk2: "能力" (2字) → 合并 → "售后服务能力" (6字)
      → 返回: ["售后服务能力", "售后服务"]

      "临时劳务派遣员工解决方案" (10字)
      → "临时劳务" (4字) → "派遣员工" (4字) → "解决方案" (4字)
      → 返回: ["临时劳务派遣员工解决方案", "临时劳务", "派遣员工", "解决方案"]
    """
    n = len(keyword)
    if n <= min_chunk:
        return [keyword]

    # 按 min_chunk 分块
    chunks = []
    i = 0
    while i < n:
        chunk = keyword[i:i + min_chunk]
        chunks.append(chunk)
        i += min_chunk

    # 有效块: 长度 >= min_chunk 的
    valid = [c for c in chunks if len(c) >= min_chunk]

    # 最后一块如果短于 min_chunk，合并到前一有效块
    if len(chunks) > 1 and len(chunks[-1]) < min_chunk:
        merged = chunks[-2] + chunks[-1]
        if merged != keyword and merged not in valid:
            valid.append(merged)

    # 返回: 完整关键词 + 有效块（去重）
    result = [keyword]
    for ch in valid:
        if ch != keyword and ch not in result:
            result.append(ch)
    return result


# ============================================================
#  Parent 列自动检测
# ============================================================

def _detect_parent_col(row, keyword_col_index, keyword_is_subjective=False):
    """
    在表格行中自 keyword_col 向左查找"类别名"所在的列。

    "类别名"特征（越靠左侧优先级越高）:
      - 文本长度 2 ~ 20 字
      - 不含冒号（：或 :）
      - 不含 0-X 分数模式
      - 不含"主观"二字
      - 不含纯数字分数（如 "14分"、"6"）
      - 优先取最靠近关键词、且最短的

    参数:
        row:                   表格行对象
        keyword_col_index:     关键词所在列号
        keyword_is_subjective: 关键词是否来自主观检测
    返回:
        parent_col_index: int | None → 类别名列号
    """
    candidates = []

    for ci in range(keyword_col_index - 1, -1, -1):
        cell_text = row.cells[ci].text.strip()
        if not cell_text:
            continue

        # 排除：含冒号的行（那是关键词本身或描述）
        if '：' in cell_text or ':' in cell_text:
            continue

        # 排除：含 0-X 分数模式
        if re.search(r'0-\d+', cell_text):
            continue

        # 排除：含"主观"
        if '主观' in cell_text:
            continue

        # 排除：纯数字分数（如 "14分"、"6"）
        if re.match(r'^\s*\d+\s*(分)?\s*$', cell_text):
            continue

        # 类别名特征：短文本（2-20字）
        if 2 <= len(cell_text) <= 20:
            candidates.append((ci, len(cell_text)))

    if not candidates:
        return None

    # 优选：越靠左、文字越短的（越高优先级）
    # 取文字最短的；同长取最靠左的
    candidates.sort(key=lambda x: (x[1], x[0]))
    return candidates[0][0]


# ============================================================
#  关键词提取 v2 — 0-X 分数模式（带 parent）
# ============================================================

def _extract_score_categories(doc, log_callback=None):
    """
    一级检测 v2：从表格中匹配 "0-X分" 模式，并提取左侧的类别名(parent)。

    核心改进（相比 word_core_hyp1.py 的 _extract_by_score_pattern）:
      1. 保留表格行结构（不扁平化），找到关键词的列位置
      2. 从关键词列向左自动检测 parent 列
      3. 支持合并单元格（空 cell 继承上一行的 parent）
      4. 返回 [{parent, type:"score", keywords}]

    参数:
        doc:          已加载的 Document 对象
        log_callback: 可选的日志回调
    返回:
        categories: list[dict] — [{"parent":..., "type":"score", "keywords":[...]}, ...]
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    # 用于合并 parent 的关键词
    parent_kw_map = {}   # {parent: set(keywords)}  (保持插入顺序)
    parent_order = []    # [parent, ...]

    # 严格匹配: "关键词：0-X分" 的正则
    score_pattern = re.compile(
        r'^\s*'
        r'(?:\d+[、.\uFF0E]\s*)?'
        r'(.+?)'
        r'[：:]\s*0-\d+'
    )

    for table in doc.tables:
        # 第一遍：遍历每行，找 current_parent
        current_parent_info = {}  # {table_idx: {parent_col_idx: str}}

        for ri, row in enumerate(table.rows):
            cells = row.cells
            detected_in_row = False

            for ci, cell in enumerate(cells):
                text = cell.text.strip()
                if not text:
                    continue

                # 只处理含 "：0-X" 的行（快速过滤）
                if not re.search(r'[：:]\s*0-\d+', text):
                    continue

                # 逐行扫描（同一单元格可能多行）
                for line in text.split('\n'):
                    line = line.strip()
                    if not line:
                        continue

                    match = score_pattern.match(line)
                    if not match:
                        continue

                    keyword = match.group(1).strip()
                    keyword = re.sub(r'[；;、，,。\.]+$', '', keyword)
                    if not keyword or len(keyword) < 2:
                        continue

                    # ---- 自动检测 parent 列 ----
                    parent_col = None
                    table_key = id(table)
                    if table_key not in current_parent_info:
                        # 第一次遇到此 table，探测 parent 列
                        parent_col = _detect_parent_col(row, ci, keyword_is_subjective=False)
                        if parent_col is not None:
                            current_parent_info[table_key] = parent_col
                    else:
                        parent_col = current_parent_info[table_key]

                    # ---- 获取 parent 值（支持合并单元格）----
                    parent_text = None
                    if parent_col is not None:
                        parent_text = cells[parent_col].text.strip()

                    # 合并单元格回退：当前行为空则从上一行继承
                    if not parent_text and ri > 0:
                        for pr in range(ri - 1, -1, -1):
                            prev_text = table.rows[pr].cells[parent_col].text.strip() if parent_col is not None else ''
                            if prev_text:
                                parent_text = prev_text
                                break

                    # 若仍未找到 parent，用关键词自身
                    if not parent_text:
                        parent_text = keyword

                    # ---- 归类 ----
                    if parent_text not in parent_kw_map:
                        parent_kw_map[parent_text] = set()
                        parent_order.append(parent_text)

                    parent_kw_map[parent_text].add(keyword)
                    log(f"  ✓ [一级·分类] {parent_text} ← {keyword}")
                    detected_in_row = True
                    break  # 该行已命中，下一个 cell

                if detected_in_row:
                    break  # 下一个 row

    # 组装返回
    categories = []
    for parent in parent_order:
        categories.append({
            "parent": parent,
            "type": "score",
            "keywords": list(parent_kw_map[parent]),
        })

    if categories:
        total_kws = sum(len(c["keywords"]) for c in categories)
        log(f"  [一级·分类] 共 {len(categories)} 个类别, {total_kws} 个关键词")
    else:
        log("  [一级·分类] 未命中")

    return categories


# ============================================================
#  关键词提取 v2 — 主观回退（带 parent + 顿号拆分）
# ============================================================

def _extract_subjective_categories(doc, log_callback=None):
    """
    二级检测 v2：在表格行级结构中查找"主观"，提取关键词 + parent。

    核心改进（相比 _extract_by_subjective_fallback）:
      1. 提取左侧类别名(parent)
      2. 支持顿号拆分（_split_by_dunhao）
      3. 返回 [{parent, type, keywords}]

    参数:
        doc:          已加载的 Document 对象
        log_callback: 可选的日志回调
    返回:
        categories: list[dict]
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    # 匹配 "X.X 关键词：" 或 "关键词：" 格式
    keyword_pattern = re.compile(
        r'^\s*'
        r'(?:\d+[.\uFF0E]\d+\s*)?'   # 可选序号前缀
        r'(.+?)'                      # 关键词文本（非贪婪）
        r'[：:]'                      # 冒号
    )

    parent_kw_map = {}   # {parent: set(keywords)}
    parent_order = []

    # 用于缓存已探测的 parent 列位置 {table_id: parent_col_idx}
    col_cache = {}

    for table in doc.tables:
        table_id = id(table)
        cached_parent_col = col_cache.get(table_id)

        for ri, row in enumerate(table.rows):
            cells = row.cells

            # ---- 找到含"主观"的列 ----
            subjective_col = -1
            for ci, cell in enumerate(cells):
                if '主观' in cell.text:
                    subjective_col = ci
                    break
            if subjective_col == -1:
                continue

            # ---- 在主观列左侧搜索关键词 ----
            for k in range(subjective_col - 1, -1, -1):
                cell_text = cells[k].text.strip()
                if not cell_text:
                    continue

                found_in_cell = False
                bare_keywords = []  # 该单元格中提取的所有裸关键词（未拆分）

                for line in cell_text.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    match = keyword_pattern.match(line)
                    if match:
                        kw = match.group(1).strip()
                        kw = re.sub(r'[；;、，,。\.]+$', '', kw)
                        if kw and len(kw) >= 2:
                            bare_keywords.append(kw)

                # 若正则没命中，但文本够长（>10字）且含顿号，则当作关键词全文
                if not bare_keywords and len(cell_text) > 10 and '、' in cell_text:
                    bare_keywords = [cell_text]

                if not bare_keywords:
                    continue

                # ---- 自动检测 parent 列（若未缓存）----
                parent_col = cached_parent_col
                if parent_col is None:
                    parent_col = _detect_parent_col(row, k, keyword_is_subjective=True)
                    if parent_col is not None:
                        col_cache[table_id] = parent_col
                        cached_parent_col = parent_col

                # ---- 获取 parent 值 ----
                parent_text = None
                if parent_col is not None:
                    parent_text = cells[parent_col].text.strip()

                # 合并单元格回退
                if not parent_text and ri > 0:
                    for pr in range(ri - 1, -1, -1):
                        ptext = table.rows[pr].cells[parent_col].text.strip() if parent_col is not None else ''
                        if ptext:
                            parent_text = ptext
                            break

                # ---- 对每个关键词处理顿号拆分 ----
                for bare_kw in bare_keywords:
                    sub_kws = _split_by_dunhao(bare_kw)

                    # 确定该关键词的 parent
                    actual_parent = parent_text if parent_text else bare_kw

                    if actual_parent not in parent_kw_map:
                        parent_kw_map[actual_parent] = set()
                        parent_order.append(actual_parent)

                    for sk in sub_kws:
                        parent_kw_map[actual_parent].add(sk)
                        log(f"  ✓ [二级·分类] {actual_parent} ← {sk}")

                found_in_cell = True
                break  # 该行已找到关键词，离开列回退循环

            # 主观列已处理，继续下一行

    # 组装返回
    categories = []
    for parent in parent_order:
        all_kws = list(parent_kw_map[parent])
        # 判断类型
        has_split = any('、' in kw for kw in all_kws) or any(
            _split_by_dunhao(kw) != [kw] for kw in all_kws
        )
        cat_type = "subjective_split" if has_split else "subjective_direct"
        categories.append({
            "parent": parent,
            "type": cat_type,
            "keywords": all_kws,
        })

    if categories:
        total_kws = sum(len(c["keywords"]) for c in categories)
        log(f"  [二级·分类] 共 {len(categories)} 个类别, {total_kws} 个关键词")
    else:
        log("  [二级·分类] 未命中")

    return categories


# ============================================================
#  关键词提取 v2 — 主入口
# ============================================================

def extract_categories_from_document(doc_path, log_callback=None):
    """
    从评标办法文档中提取分类关键词。

    检测顺序:
      一级: 0-X 分数模式 → 提取 parent + keywords
      二级: 主观回退      → 提取 parent + keywords（含顿号拆分）
    两级独立执行，结果合并。

    返回:
        (success: bool, categories: list[dict] | error_msg: str)
        categories = [{"parent":..., "type":..., "keywords":[...]}, ...]
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    doc_path = os.path.normpath(doc_path)
    log(f"正在提取分类关键词: {os.path.basename(doc_path)}")

    ok, result = _safe_open_document(doc_path, log_callback=log)
    if not ok:
        return False, result
    doc = result

    # ---- 一级: 0-X 分数模式 ----
    log("  [一级] 匹配 '0-X分' 模式（分类提取）")
    score_cats = _extract_score_categories(doc, log_callback=log)

    # ---- 二级: 主观回退 ----
    log("  [二级] '主观'回退（分类提取）")
    subjective_cats = _extract_subjective_categories(doc, log_callback=log)

    # ---- 合并两级结果（按 parent 合并）----
    all_cats = {}
    cat_order = []
    for cat in score_cats + subjective_cats:
        parent = cat["parent"]
        if parent not in all_cats:
            all_cats[parent] = {"parent": parent, "type": cat["type"], "keywords": set()}
            cat_order.append(parent)
        # 合并关键词
        all_cats[parent]["keywords"].update(cat["keywords"])

    categories = []
    for parent in cat_order:
        categories.append({
            "parent": parent,
            "type": all_cats[parent]["type"],
            "keywords": list(all_cats[parent]["keywords"]),
        })

    if not categories:
        log("  [警告] 两级检测均未找到可用关键词")
        return False, '文档中未找到包含"0-X分"或"主观"评分标准的行，无法提取关键词'

    total_kws = sum(len(c["keywords"]) for c in categories)
    log(f"  共提取 {len(categories)} 个类别, {total_kws} 个不重复关键词")
    return True, categories


def extract_categories_from_documents(doc_paths, log_callback=None):
    """
    从多个评标办法文档中提取分类关键词（跨文档合并）。

    同一 parent 在不同文档中的关键词自动合并去重。

    返回:
        (success: bool, categories: list[dict] | error_msg: str)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    merged = {}
    order = []
    failed = []

    for doc_path in doc_paths:
        log(f"\n处理文档: {os.path.basename(doc_path)}")
        success, result = extract_categories_from_document(doc_path, log_callback=log)
        if success:
            for cat in result:
                parent = cat["parent"]
                if parent not in merged:
                    merged[parent] = {"parent": parent, "type": cat["type"], "keywords": set()}
                    order.append(parent)
                merged[parent]["keywords"].update(cat["keywords"])
        else:
            failed.append((doc_path, result))

    if not merged:
        return False, f"所有文档均未能提取到关键词：{failed}"

    categories = []
    for parent in order:
        categories.append({
            "parent": parent,
            "type": merged[parent]["type"],
            "keywords": list(merged[parent]["keywords"]),
        })

    total_kws = sum(len(c["keywords"]) for c in categories)
    log(f"\n  跨 {len(doc_paths)} 个文档共 {len(categories)} 个类别, {total_kws} 个关键词")
    return True, categories


# ============================================================
#  分类搜索（精准 → 模糊 两级）
# ============================================================

def categorized_search(categories, folder_path, log_callback=None):
    """
    对每个分类的每个关键词执行精准→模糊两级搜索。

    策略:
      1. 精准匹配（文件名包含完整关键词）→ 命中 → 加入结果
      2. 未命中 → 模糊匹配（4字分块）→ 命中 → 加入结果
      3. 均未命中 → 跳过该关键词

    结果以分类为单位去重。

    参数:
        categories:  [{"parent":..., "keywords":[...]}, ...]
        folder_path: 搜索文件夹路径
        log_callback: 日志回调

    返回:
        categories: 原地填充 "matched_files" 和 "search_details"
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    log("=" * 50)
    log("分类搜索开始（精准 → 模糊 两级）")
    log(f"  搜索文件夹: {folder_path}")
    log("=" * 50)

    # 扫描文件夹中的 Word 文件
    all_files = []
    for f in sorted(os.listdir(folder_path)):
        fl = f.lower()
        if fl.endswith('.docx'):
            all_files.append(f)

    log(f"  文件夹内文档总数: {len(all_files)}")

    for cat in categories:
        parent = cat["parent"]
        keywords = cat["keywords"]
        log(f"\n--- {parent}（{len(keywords)} 个关键词）---")

        matched = []
        search_details = []  # [(keyword, file, method), ...]

        for kw in keywords:
            # ---- Step A: 精准匹配 ----
            precision_hits = []
            for f in all_files:
                if kw.lower() in f.lower():
                    full_path = os.path.join(folder_path, f)
                    precision_hits.append(full_path)

            if precision_hits:
                for fp in precision_hits:
                    matched.append(fp)
                    search_details.append((kw, os.path.basename(fp), "精准"))
                log(f"  ✓ [{kw}] 精准匹配 → {len(precision_hits)} 个文件")
                continue

            # ---- Step B: 模糊匹配（4字分块）----
            fuzzy_kws = _fuzzy_chunks_4char(kw, min_chunk=4)
            fuzzy_hits = []
            for fw in fuzzy_kws:
                for f in all_files:
                    fp = os.path.join(folder_path, f)
                    if fp not in fuzzy_hits and fw.lower() in f.lower():
                        fuzzy_hits.append(fp)

            if fuzzy_hits:
                for fp in fuzzy_hits:
                    if fp not in matched:
                        matched.append(fp)
                        search_details.append((kw, os.path.basename(fp), "模糊"))
                log(f"  ◐ [{kw}] 模糊匹配 → {len(fuzzy_hits)} 个文件")
            else:
                log(f"  ✗ [{kw}] 未找到匹配文件")

        # 去重
        cat["matched_files"] = list(dict.fromkeys(matched))
        cat["search_details"] = search_details

    total_matched = sum(len(c.get("matched_files", [])) for c in categories)
    log(f"\n{'=' * 50}")
    log(f"搜索完成: 共 {len(categories)} 个类别, 匹配 {total_matched} 个文件(含跨类重复)")
    log(f"{'=' * 50}")

    return categories


# ============================================================
#  分类合并
# ============================================================

def merge_documents_categorized(categories, output_path, log_callback=None):
    """
    按分类层级合并 Word 文档。

    输出结构:
      一、服务能力
        第一章 售后服务能力.docx
        第二章 人员储备能力.docx
        ...
      二、管理制度
        第一章 管理运作制度.docx
        ...

    每个分类:
      - 先插入"X、类别名"一级标题（Heading 1 样式）
      - 然后按顺序合并该分类下的文件
      - 文件间插入"下一页分节符"
      - 分类间也插入分节

    参数:
        categories:   [{"parent":..., "matched_files":[...]}, ...]
                       matched_files 需已填充（由 categorized_search 完成）
        output_path:  输出文件路径
        log_callback: 日志回调

    返回:
        (success: bool, message: str)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    # ---- 全局预去重：同一文件只保留在第一个出现的分类中 ----
    # 使用大小写不敏感比较（Windows 文件系统大小写不敏感）
    seen_paths = set()
    for cat in categories:
        deduped = []
        for fp in cat.get("matched_files", []):
            fp = os.path.normpath(fp)
            fp_lower = fp.lower()  # Windows 大小写不敏感
            if fp_lower not in seen_paths:
                seen_paths.add(fp_lower)
                deduped.append(fp)
            else:
                log(f"  [去重] 跨类重复文件: {os.path.basename(fp)}")
        cat["matched_files"] = deduped

    # 收集所有要合并的文件
    all_files = []
    for cat in categories:
        for f in cat.get("matched_files", []):
            if f not in all_files:
                all_files.append(f)

    if not all_files:
        return False, "没有匹配到任何文件"

    # 规范化路径
    output_path = os.path.normpath(output_path)

    log("=" * 50)
    log("分类合并开始")
    log(f"  类别数: {len(categories)}")
    log(f"  文件数: {len(all_files)}")
    log(f"  输出路径: {output_path}")
    log("=" * 50)

    # ---- 以第一个文件为基础文档 ----
    first_file = all_files[0]
    first_name = os.path.basename(first_file)

    log(f"\n  [基础] 以「{first_name}」为基础文档")

    from docx import Document as Doc
    ok, result = _safe_open_document(first_file, log_callback=log)
    if not ok:
        return False, f"无法读取首个文件：{result}"
    dst_doc = result
    log(f"  ✓ 已加载基础文档")

    # 设置页边距（使用统一函数，与后处理保持一致）
    _normalize_section_margins(dst_doc, log, verbose=True)

    # 提取基础文档的页眉页脚
    _extract_base_hf_rids(dst_doc)

    # ---- 辅助：动态查找 Heading 1 样式 ID ----
    def _find_heading1_id(doc):
        """在目标文档中查找 Heading 1 的 style_id（不硬编码'1'）。"""
        for style in doc.styles:
            if style.name == 'Heading 1' and style.type is not None \
                    and style.type.name == 'PARAGRAPH':
                return style.style_id
        # 回退：尝试列表 1, 10, 11 等常见 Heading 1 ID
        fallback_ids = ['1', '10', '11', 'Heading1', 'heading1', 'Heading 1']
        for fid in fallback_ids:
            try:
                style = doc.styles[fid]
                if style and style.type is not None and style.type.name == 'PARAGRAPH':
                    return fid
            except Exception:
                continue
        return '1'  # 最终回退

    _heading1_id = _find_heading1_id(dst_doc)
    log(f"  [样式] 目标文档 Heading 1 style_id = '{_heading1_id}'")

    # ---- 清空基础文档 body 内容（保留 sectPr），避免原始游离内容与分类标题重复 ----
    _clear_body_content(dst_doc, log)
    log(f"  [清理] 已清空基础文档正文，准备按分类重建")

    def _insert_category_heading(doc, title_text):
        """在文档末尾插入一个 Heading 1 级别的标题段落（左对齐，禁用自动编号）。"""
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _OxmlElement

        p = _OxmlElement('w:p')
        pPr = _OxmlElement('w:pPr')

        # 样式：Heading 1
        pStyle = _OxmlElement('w:pStyle')
        pStyle.set(_qn('w:val'), _heading1_id)
        pPr.append(pStyle)

        # 对齐方式：左对齐
        jc = _OxmlElement('w:jc')
        jc.set(_qn('w:val'), 'left')
        pPr.append(jc)

        # 禁用多级列表自动编号（避免 Heading 1 样式自带编号导致乱码）
        numPr = _OxmlElement('w:numPr')
        ilvl = _OxmlElement('w:ilvl')
        ilvl.set(_qn('w:val'), '0')
        numId = _OxmlElement('w:numId')
        numId.set(_qn('w:val'), '0')
        numPr.append(ilvl)
        numPr.append(numId)
        pPr.append(numPr)

        p.append(pPr)

        r = _OxmlElement('w:r')
        t = _OxmlElement('w:t')
        t.set(_qn('xml:space'), 'preserve')
        t.text = title_text
        r.append(t)
        p.append(r)

        doc.element.body.append(p)

    # ---- 按分类合并 ----
    processed_files = set()  # 存储小写路径用于去重
    total_copied = 0
    category_index = 0

    for cat in categories:
        parent = cat["parent"]
        files = [f for f in cat.get("matched_files", [])
                 if f.lower() not in processed_files]
        if not files:
            continue

        category_index += 1

        num = _CHINESE_NUMS[category_index - 1] if category_index <= len(_CHINESE_NUMS) else str(category_index)
        heading = f"{num}、{parent}"
        log(f"\n  [{num}] 一级标题: {heading}")
        log(f"       包含 {len(files)} 个文件")

        # 插入分节符（第一个分类不加，直接从第 1 页开始）
        _insert_section_marker(dst_doc, first=(category_index == 1))

        # 插入一级标题
        _insert_category_heading(dst_doc, heading)
        log(f"    → 已插入一级标题段落")

        # 合并该分类的文件
        for file_idx, file_path in enumerate(files):
            file_name = os.path.basename(file_path)
            log(f"    [{file_idx + 1}/{len(files)}] {file_name}")

            try:
                ok, result = _safe_open_document(file_path, log_callback=log)
                if not ok:
                    log(f"    [跳过] {file_name}: {result}")
                    continue
                src_doc = result
            except Exception as e:
                log(f"    [错误] 无法读取: {e}")
                continue

            # 预加载源文档的图片关系到目标文档（确保图片能正确复制）
            _preload_image_rels(src_doc, dst_doc, log)

            # 合并 numbering.xml 列表定义（避免自定义编号格式丢失，如"项目 1"→"•"）
            num_id_map = _merge_numbering(src_doc, dst_doc, log)

            # 文件间插入分节符
            if file_idx > 0:
                insert_section_break(dst_doc)
                log(f"    → 已插入分节符")

            # 样式映射
            style_map = _build_style_map(src_doc, dst_doc)

            # 复制内容
            try:
                n = copy_all_content(src_doc, dst_doc, log_callback=log,
                                     style_map=style_map if style_map else None,
                                     num_id_map=num_id_map if num_id_map else None)
                total_copied += n
            except Exception as e:
                log(f"    [错误] 复制内容失败: {e}")
                continue

            processed_files.add(file_path.lower())

    # ---- 统一页边距 ----
    _normalize_section_margins(dst_doc, log)

    # ---- 剥离标题自动编号（标题用纯文本，不使用 Word 自动编号）----
    _strip_heading_auto_numbering(dst_doc, log)

    # ---- 保存 ----
    try:
        dst_doc.save(output_path)
    except Exception as e:
        return False, f"保存失败：{e}"

    log(f"\n{'=' * 50}")
    log(f"合并完成 → {os.path.basename(output_path)}")
    log(f"共 {len(categories)} 个类别, {len(processed_files)} 个文件")
    log(f"{'=' * 50}")
    return True, output_path


def _clear_body_content(doc, log):
    """
    清空文档 body 中所有段落和表格，只保留 sectPr（页面设置）。
    这样可以让后续按分类添加的内容成为文档的唯一正文，
    不会残留基础文档原始的游离内容（避免同一文件出现两次）。
    """
    body = doc.element.body
    to_remove = []
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('p', 'tbl'):
            to_remove.append(child)
    for el in to_remove:
        body.remove(el)
    if to_remove:
        log(f"    → 已清空 {len(to_remove)} 个元素（段落+表格），保留 sectPr")


def _strip_heading_auto_numbering(doc, log):
    """
    后处理：剥离所有标题的 Word 自动编号并回填静态文字编号前缀。

    背景：
      合并后基础文档的"标题 N"样式绑定多级列表（w:numPr 写在 styles.xml
      的样式定义里）。所有源文件标题经样式映射后共用同一套列表计数器，
      出现"第五章 第一章..."双重编号；表格内标题的计数器也跨文件累积。

    本函数做法：
      ① 对【正文 + 表格内】的标题段落操作：
         - 在其 w:pPr 内写入 w:numPr/w:numId='0' 覆盖样式级自动编号
         - 回填静态文字编号前缀（第X章 /（N）/ N.）
      ② 计数器在 H1（分类）处重置，H2 按分类内连续编号；
         H3 每遇 H2 重置，H4 每遇 H3 重置。
      ③ H5（标题5）同样纳入处理（numId='0' + 回填"（N）"），
        每遇 H1/H2/H3/H4 重置计数器。

    编号规则：
      H1: 由 merge 生成，形如"一、xxx"，不回填（仅作 H2/H5 计数器重置信号）
      H2: "第N章 xxx"   —— 每遇 H1 重置，N 用中文数字，分类内连续
      H3: "（N）xxx"     —— 每遇 H2 重置，N 用中文数字
      H4: "N. xxx"       —— 每遇 H3 重置，N 用阿拉伯数字
      H5: "（N）xxx"     —— 每遇 H1/H2/H3/H4 重置，N 用中文数字
    """
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OxmlElement

    # ---- 标题样式判定：返回级别 1~4，非标题返回 0 ----
    def _heading_level(style):
        try:
            sname = (style.name or '').strip()
        except Exception:
            return 0
        m = re.match(r'^(Heading|标题)\s*(\d+)$', sname, re.IGNORECASE)
        if not m:
            return 0
        try:
            lvl = int(m.group(2))
        except Exception:
            return 0
        return lvl if 1 <= lvl <= 5 else 0

    # ---- 0) 收集标题 style_id → level（注意：不修改样式定义！）----
    # 说明：若在此处删除样式级 w:numPr，会影响所有使用该样式的段落
    #       （含表格内标题），导致表格编号被一并删除且无回填补救。
    #       因此改为只对【正文段落】写入 numId='0' 覆盖样式级编号，
    #       表格段落完全不动、保持其原始样式绑定。
    heading_style_ids = {}   # style_id -> level
    for style in doc.styles:
        lvl = _heading_level(style)
        if lvl == 0:
            continue
        try:
            heading_style_ids[style.style_id] = lvl
        except Exception:
            pass

    if not heading_style_ids:
        log(f"  [去编号] 未发现任何标题样式，跳过回填")
        return

    # ---- 中文数字映射 ----
    chinese_nums = [
        "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
        "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
        "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十",
    ]

    def _cn(n):
        return chinese_nums[n - 1] if 1 <= n <= len(chinese_nums) else str(n)

    # ---- 预编译"已有编号前缀"检测正则（避免重复回填）----
    h2_existing = re.compile(r'^第[一二三四五六七八九十百零\d]+章\s*')
    h3_existing = re.compile(r'^[（(][一二三四五六七八九十百零\d]+[)）]\s*')
    h4_existing = re.compile(r'^\d+[\.\、]\s*')
    h5_existing = re.compile(r'^[（(][一二三四五六七八九十百零\d]+[)）]\s*')  # H5 与 H3 相同格式

    # ---- 计数器 ----
    chapter_counter = 0   # H2（每 H1 重置）
    section_counter = 0   # H3（每 H2 重置）
    item_counter = 0      # H4（每 H3 重置）
    h5_counter = 0        # H5（每 H1/H2/H3/H4 重置）

    # ---- 遍历 body 下所有段落（含表格单元格内的段落）----
    def _iter_all_paragraphs(body):
        """递归产出 body 下所有 w:p 元素（含表格、列表等任意嵌套深度）。"""
        for p in body.iter(_qn('w:p')):
            yield p

    def _get_full_text(p):
        texts = p.findall('.//' + _qn('w:t'))
        return ''.join(t.text or '' for t in texts)

    def _prepend_text(p, prefix):
        """在段落第一个 w:t 文本前插入 prefix（不改变其他格式）。"""
        texts = p.findall('.//' + _qn('w:t'))
        if texts:
            first = texts[0]
            first.text = prefix + (first.text or '')
        else:
            # 段落无 w:t（极端情况）：新建一个 run 写入 prefix
            r = _OxmlElement('w:r')
            t = _OxmlElement('w:t')
            t.set(_qn('xml:space'), 'preserve')
            t.text = prefix
            r.append(t)
            p.append(r)

    def _disable_para_numbering(pPr):
        """
        用 numId='0' 覆盖该段落的样式级自动编号。

        写入/修改段落 pPr 内的 w:numPr/w:numId='0'，使该段落不再参与
        多级列表计数（保留表格不受影响，因为表格段落不进入本函数）。
        """
        num_pr = pPr.find(_qn('w:numPr'))
        if num_pr is None:
            num_pr = _OxmlElement('w:numPr')
            pPr.append(num_pr)
        numId = num_pr.find(_qn('w:numId'))
        if numId is None:
            numId = _OxmlElement('w:numId')
            num_pr.append(numId)
        numId.set(_qn('w:val'), '0')

    # ---- 主扫描：剥离段落内联编号 + 回填静态编号 ----
    body = doc.element.body
    stats = {'h1': 0, 'h2': 0, 'h3': 0, 'h4': 0, 'h5': 0,
             'h2_fill': 0, 'h3_fill': 0, 'h4_fill': 0, 'h5_fill': 0}

    for p in _iter_all_paragraphs(body):
        pPr = p.find(_qn('w:pPr'))
        if pPr is None:
            continue

        ps_el = pPr.find(_qn('w:pStyle'))
        if ps_el is None:
            continue
        sid = ps_el.get(_qn('w:val'))
        if not sid or sid not in heading_style_ids:
            continue

        lvl = heading_style_ids[sid]

        # ① 用 numId='0' 覆盖样式级自动编号（仅在正文段落；表格不进入此函数）
        _disable_para_numbering(pPr)

        full = _get_full_text(p).strip()
        if not full:
            continue

        if lvl == 1:
            # H1：仅重置下级计数器（由 merge 生成的"一、xxx"不回填）
            chapter_counter = 0
            section_counter = 0
            item_counter = 0
            h5_counter = 0
            stats['h1'] += 1
            continue

        if lvl == 2:
            chapter_counter += 1
            section_counter = 0
            item_counter = 0
            h5_counter = 0
            stats['h2'] += 1
            if h2_existing.match(full):
                continue    # 已有"第X章"前缀，跳过避免重复
            _prepend_text(p, f'第{_cn(chapter_counter)}章 ')
            stats['h2_fill'] += 1

        elif lvl == 3:
            section_counter += 1
            item_counter = 0
            h5_counter = 0
            stats['h3'] += 1
            if h3_existing.match(full):
                continue
            _prepend_text(p, f'（{_cn(section_counter)}）')
            stats['h3_fill'] += 1

        elif lvl == 4:
            item_counter += 1
            h5_counter = 0
            stats['h4'] += 1
            if h4_existing.match(full):
                continue
            _prepend_text(p, f'{item_counter}. ')
            stats['h4_fill'] += 1

        elif lvl == 5:
            h5_counter += 1
            stats['h5'] += 1
            if h5_existing.match(full):
                continue
            _prepend_text(p, f'（{h5_counter}）')
            stats['h5_fill'] += 1

    log(f"  [去编号] 完成：H1={stats['h1']}, "
        f"H2={stats['h2']}(填空{stats['h2_fill']}), "
        f"H3={stats['h3']}(填空{stats['h3_fill']}), "
        f"H4={stats['h4']}(填空{stats['h4_fill']}), "
        f"H5={stats['h5']}(填空{stats['h5_fill']}) —— "
        f"编号前缀已回填为静态文字，标题现以纯文本显示")


def _insert_section_marker(dst_doc, first=False):
    """在文档末尾插入分节标记（分类间用）。"""
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OxmlElement

    if first:
        return  # 第一个分类不加分节符

    insert_section_break(dst_doc)


def _normalize_section_margins(dst_doc, log, verbose=False):
    """统一所有 section 的页边距。"""
    try:
        from docx.shared import Cm
        for sect in dst_doc.sections:
            sect.top_margin = Cm(2)
            sect.bottom_margin = Cm(2)
            sect.left_margin = Cm(1.8)
            sect.right_margin = Cm(1.8)
            sect.header_distance = Cm(2)
            sect.footer_distance = Cm(1.5)
            try:
                sect.gutter = Cm(0)
            except Exception:
                pass
        if verbose:
            log(f"  ✓ 已设置页边距")
    except Exception:
        pass  # 静默失败
