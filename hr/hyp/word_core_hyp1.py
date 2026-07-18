# -*- coding: utf-8 -*-
"""
Word 文档合并工具 — 关键词提取与文件搜索模块（word_core_hyp1.py）

本模块专注于「分析」侧逻辑：
  - 关键词提取（两级检测）
        · 一级: "0-X分" 分数模式（0-1 ~ 0-5 等）
        · 二级: "主观" 回退检测（表格行级，向左搜索 X.X关键字：）
  - 文件搜索（模糊 / 精确）

合并引擎（样式复制、XML 克隆、分节符）已合并至本文件。
分类层级合并在 word_core_hyp2.py。
"""

import os
import re
import zipfile
from docx import Document


# ============================================================
#  文件格式检测（仅检测，不自动转换）
# ============================================================

def _is_doc_format(filepath):
    """
    检测文件是否为旧版 .doc 格式（Word 97-2003 OLE2/Compound Document）。

    判断方式：
      - .doc 文件头: D0 CF 11 E0（OLE2 复合文档签名）
      - .docx 文件头: PK（ZIP 包签名）
    返回 True=旧版.doc格式, False=新版.docx格式或无法判断
    """
    try:
        with open(filepath, 'rb') as f:
            header = f.read(4)
            if header[:4] == b'\xd0\xcf\x11\xe0':
                return True
            if header[:2] == b'PK':
                return False
        return False
    except Exception:
        return False


def _check_docx_valid(filepath):
    """
    验证文件是否为有效的 .docx（ZIP 包）。
    返回 (ok: bool, error_msg: str)。
    """
    if not os.path.exists(filepath):
        return False, f"文件不存在：{filepath}"
    if _is_doc_format(filepath):
        return False, (
            f"文件「{os.path.basename(filepath)}」是旧版 .doc 格式（Word 97-2003），"
            f"\n不支持直接读取。\n\n"
            f"解决方法：用 Word 打开该文件，\n"
            f"「文件 → 另存为 → Word 文档(*.docx)」后重试。"
        )
    try:
        with zipfile.ZipFile(filepath, 'r'):
            pass
        return True, ""
    except zipfile.BadZipFile:
        return False, (
            f"文件「{os.path.basename(filepath)}」不是有效的 .docx 文件（ZIP 包损坏），"
            f"\n请确认文件是否已正确转换为 .docx 格式。"
        )
    except Exception as e:
        return False, f"无法读取文件「{os.path.basename(filepath)}」：{e}"


def _safe_open_document(filepath, log_callback=None):
    """
    安全打开 Word 文档，带格式检测和友好错误提示。

    返回 (success: bool, document: Document | error_msg: str)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    ok, err = _check_docx_valid(filepath)
    if not ok:
        return False, err
    try:
        return True, Document(filepath)
    except Exception as e:
        return False, f"无法读取文档「{os.path.basename(filepath)}」：{e}"

import copy
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement


# Word 使用 twips 单位：1 cm = 567 twips（Word 内部标准）
_TWIPS_PER_CM = 567


def _dedup_image_relate(dst_doc, image_target_part):
    """
    将图片注册到目标文档，自动去重 + 防同名碰撞（v3：手动创建 ImagePart）。

    背景：不同源文档可能有同名但不同内容的图片（如 image2.png），
    python-docx 的 relate_to 在处理跨文档 ImagePart 时，内部的
    get_or_add_part 可能忽略 _partname 的临时修改，导致图片丢失
    或 ZIP 重复条目。

    本函数绕过 python-docx 的 partname 管理，改为手动创建全新
    ImagePart 并赋予全局唯一名称，彻底杜绝同名冲突。

    流程：
      1. 对图片 blob 做 MD5 指纹，相同内容复用已有 rId
      2. 未命中时：用 dst_doc._image_global_counter 生成唯一 partname
      3. 手动创建 ImagePart(unique_partname, content_type, blob)
      4. 通过 dst_part.relate_to 注册到目标包并建立关系
    """
    import hashlib
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.parts.image import ImagePart

    dst_part = dst_doc.part

    # ---- 懒初始化 ----
    if not hasattr(dst_doc, '_image_dedup_cache'):
        dst_doc._image_dedup_cache = {}
        dst_doc._image_global_counter = 0
    cache = dst_doc._image_dedup_cache

    # ---- 计算 MD5 指纹 ----
    try:
        blob = image_target_part.blob
        fingerprint = hashlib.md5(blob).hexdigest()
    except Exception:
        # 无法读取 blob 时回退到原始 relate_to
        return dst_part.relate_to(image_target_part, RT.IMAGE)

    # ---- 缓存命中：直接复用已有 rId ----
    if fingerprint in cache:
        return cache[fingerprint]

    # ---- 未命中：创建全新 ImagePart，使用全局唯一 partname ----
    dst_doc._image_global_counter += 1

    # 从源 partname 提取扩展名（保留原始格式）
    src_partname = image_target_part.partname
    if isinstance(src_partname, str):
        ext = src_partname.rsplit('.', 1)[-1] if '.' in src_partname else 'png'
    else:
        ext = 'png'

    # 使用 img_ 前缀 + 6 位全局计数器，确保与 Word 自动命名的 imageN.xxx 不冲突
    unique_partname = PackURI(f'/word/media/img_{dst_doc._image_global_counter:06d}.{ext}')
    content_type = image_target_part.content_type

    # 在目标文档包中创建全新 ImagePart（不传入 package，由 relate_to 在 get_or_add_part 中关联）
    new_image_part = ImagePart(unique_partname, content_type, blob)

    # 通过 relate_to 将新 Part 注册到目标包并建立关系
    new_rid = dst_part.relate_to(new_image_part, RT.IMAGE)

    cache[fingerprint] = new_rid
    return new_rid


def _get_body_sectPr(doc):
    """获取 body 的 sectPr 元素（若存在）。"""
    body = doc.element.body
    sectPrs = body.findall(qn('w:sectPr'))
    if sectPrs:
        return sectPrs[-1]
    # 回退：尝试从最后一个 section 获取
    try:
        return doc.sections[-1]._sectPr
    except Exception:
        return None


def _preload_image_rels(src_doc, dst_doc, log):
    """
    预先将源文档中的所有图片关系注册到目标文档。
    这样在后续 clone_element_to_doc 时可以直接查找已注册的关系。
    """
    src_part = src_doc.part

    preload_count = 0
    for rel_id, rel in src_part.rels.items():
        if "image" in rel.reltype:
            try:
                image_part = rel.target_part
                # 去重注册：相同内容复用已有 rId
                new_rId = _dedup_image_relate(dst_doc, image_part)
                preload_count += 1
            except Exception:
                pass  # 可能已经注册过

    if preload_count > 0:
        log(f"  ✓ 已预注册 {preload_count} 个图片关系到目标文档")


# ============================================================
#  XML 元素克隆引擎（保留所有格式 + 图片重映射）
# ============================================================

def clone_element_to_doc(src_element, src_doc, dst_doc):
    """
    将源文档中的 body 子元素（段落/表格）完整复制到目标文档：
    1. deepcopy XML 保留所有格式（字体/颜色/大小/对齐等）
    2. 重新映射图片引用：
       - a:blip（r:embed / r:link） — DrawingML 格式（包括形状填充图）
       - v:imagedata（r:id） — 旧版 VML 兼容格式
       - w:pict 包裹的 VML 图片
       - mc:AlternateContent 兼容性图片
    """
    new_el = copy.deepcopy(src_element)

    src_part = src_doc.part

    # ========== 1) a:blip → r:embed / r:link（DrawingML，包括形状填充图）==========
    for blip in new_el.findall('.//' + qn('a:blip')):
        rid = blip.get(qn('r:embed')) or blip.get(qn('r:link'))
        if not rid or rid not in src_part.rels:
            continue
        rel = src_part.rels[rid]
        if "image" not in rel.reltype:
            continue
        try:
            new_rid = _dedup_image_relate(dst_doc, rel.target_part)
            blip.set(qn('r:embed'), new_rid)
            # 清除 r:link（如果是外部链接）
            blip.attrib.pop(qn('r:link'), None)
        except (KeyError, Exception):
            pass

    # ========== 2) v:imagedata → r:id / o:relid（VML 图片）==========
    # 注：VML 图片可能使用 r:id（标准）或 o:relid（OLE 兼容），
    # 也可能包裹在 mc:AlternateContent/mc:Fallback 中。
    # 'v' 和 'o' 命名空间前缀不一定在所有文档中都注册，需要用 try/except 保护。
    try:
        _qn_o_relid = qn('o:relid')
    except KeyError:
        _qn_o_relid = None

    try:
        _qn_v_imagedata = qn('v:imagedata')
    except KeyError:
        _qn_v_imagedata = None

    if _qn_v_imagedata is not None:
        for img_data in new_el.findall('.//' + _qn_v_imagedata):
            # 尝试多种 VML 图片引用属性：r:id → o:relid
            rid = None
            attr_name = None  # 记录实际使用的属性名，写入时复用
            for attr in (qn('r:id'), _qn_o_relid):
                if attr is None:
                    continue
                val = img_data.get(attr)
                if val and val in src_part.rels:
                    rid = val
                    attr_name = attr
                    break
            if not rid:
                continue
            rel = src_part.rels[rid]
            if "image" not in rel.reltype:
                continue
            try:
                new_rid = _dedup_image_relate(dst_doc, rel.target_part)
                img_data.set(attr_name, new_rid)
            except (KeyError, Exception):
                pass

    # ========== 3) w:pict 包裹的 VML 图片（v:shape 内的 v:imagedata）==========
    # 注：section 2 的 './/v:imagedata' 已覆盖 w:pict 内的普通情况，
    # 此处补充处理 section 2 未覆盖到的边缘格式。
    try:
        for pict in new_el.findall('.//' + qn('w:pict')):
            for img_data in pict.iter():
                tag = img_data.tag.split('}')[-1] if '}' in img_data.tag else img_data.tag
                if tag != 'imagedata':
                    continue
                # 跳过已被 section 2 处理的（有 r:id 或 o:relid）
                rid = img_data.get(qn('r:id'))
                if not rid and _qn_o_relid is not None:
                    rid = img_data.get(_qn_o_relid)
                if not rid or rid not in src_part.rels:
                    continue
                rel = src_part.rels[rid]
                if "image" not in rel.reltype:
                    continue
                try:
                    new_rid = _dedup_image_relate(dst_doc, rel.target_part)
                    # 用实际使用的属性写回
                    if img_data.get(qn('r:id')):
                        img_data.set(qn('r:id'), new_rid)
                    elif _qn_o_relid is not None and img_data.get(_qn_o_relid):
                        img_data.set(_qn_o_relid, new_rid)
                except (KeyError, Exception):
                    pass
    except Exception:
        pass

    return new_el


def copy_all_content(src_doc, dst_doc, log_callback=None, style_map=None, num_id_map=None):
    """
    将源文档的全部内容（段落、表格）原样复制到目标文档末尾。
    保留所有格式、图片、表格、字体、颜色。

    如果提供了 style_map，则会将源文档中的样式 ID 映射到
    目标文档的对应样式 ID（方案 A：强制使用基础文档的样式体系）。
    如果提供了 num_id_map，则对段落中的 numPr/numId 做重映射。

    参数:
        src_doc:       源 Document 对象
        dst_doc:       目标 Document 对象（内容将追加到其 body 末尾）
        log_callback:  可选的日志回调
        style_map:     样式 ID 映射表 {源 styleId: 目标 styleId}
        num_id_map:    编号 ID 映射表 {源 numId: 目标 numId}
    返回:
        复制的元素总数（段落 + 表格）
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    import copy as _cpy

    def _promote_para(src_p, dst_p):
        """把【源样式级】 w:numPr 提升到 dst 段落直接格式。

        仅当 src 段落直接格式无 numPr、且源样式定义含有效 numId 时，
        将源样式里的 w:numPr 深拷贝到 dst 段落 w:pPr（pStyle 之后）。
        dst 若已直接含 numPr（含显式 numId=0 关闭编号）则不覆盖，
        以确保源文档中"已显式关闭编号"的段落（如表格内"渠道类型"）
        保持无编号、不被误提升。
        """
        _src_pPr = src_p.find(qn('w:pPr'))
        if _src_pPr is None:
            return
        _src_pStyle = _src_pPr.find(qn('w:pStyle'))
        if _src_pStyle is None:
            return
        _src_sid = _src_pStyle.get(qn('w:val'))
        if not _src_sid:
            return
        # 直接从 XML 层面查找源样式定义（版本无关）：
        # 部分 python-docx 版本的 Styles.get_by_id 需要 style_type 第二参数，
        # 单参数调用会抛 TypeError（被 except 吞掉），导致样式级 numPr 提升从未执行。
        _src_style_el = None
        try:
            _styles_root = src_doc.styles.element
            for _st in _styles_root.findall(qn('w:style')):
                if _st.get(qn('w:styleId')) == _src_sid:
                    _src_style_el = _st
                    break
        except Exception:
            return
        if _src_style_el is None:
            return
        _src_style_pPr = _src_style_el.find(qn('w:pPr'))
        if _src_style_pPr is None:
            return
        _src_num_pr = _src_style_pPr.find(qn('w:numPr'))
        if _src_num_pr is None:
            return
        _src_nid = _src_num_pr.find(qn('w:numId'))
        if _src_nid is None or _src_nid.get(qn('w:val')) in (None, '0'):
            return
        _dst_pPr = dst_p.find(qn('w:pPr'))
        if _dst_pPr is None:
            return
        # 目标段落直接格式已含 numPr（含 numId=0 显式关闭）则保留，不覆盖
        if _dst_pPr.find(qn('w:numPr')) is not None:
            return
        _cloned_np = _cpy.deepcopy(_src_num_pr)
        _dst_pStyle = _dst_pPr.find(qn('w:pStyle'))
        if _dst_pStyle is not None:
            _dst_pStyle.addnext(_cloned_np)
        else:
            _dst_pPr.insert(0, _cloned_np)

    dst_body = dst_doc.element.body
    count = 0

    for child in src_doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        # 跳过 sectPr（章节属性，不复制，避免章节属性混乱）
        if tag == 'sectPr':
            continue

        if tag in ('p', 'tbl'):
            try:
                new_el = clone_element_to_doc(child, src_doc, dst_doc)

                # 应用样式映射（方案 A）
                if style_map:
                    _apply_style_map(new_el, style_map)

                # 把【源样式级】 w:numPr 提升到段落直接格式：
                # 正文段落直接处理；表格则递归处理其内部所有 w:p 段落。
                if tag == 'p':
                    _promote_para(child, new_el)
                else:  # tag == 'tbl'
                    _src_paras = child.findall('.//' + qn('w:p'))
                    _dst_paras = new_el.findall('.//' + qn('w:p'))
                    for _sp, _dp in zip(_src_paras, _dst_paras):
                        _promote_para(_sp, _dp)

                # 应用编号 ID 重映射（numbering.xml 合并后的 numId 对齐）
                if num_id_map and len(new_el) > 0:
                    all_numpr = new_el.findall('.//' + qn('w:numPr'))
                    # ---- 诊断：统计 numPr 结构（确认是内联 numId 还是样式级编号）----
                    diag_total = 0
                    diag_has_numid = 0
                    diag_has_ilvl = 0
                    diag_remapped = 0
                    diag_ilvl_only = 0   # 仅含 ilvl 不含 numId（可能走样式编号）
                    diag_missed = 0       # 有 numId 但不在 num_id_map 中
                    for np in all_numpr:
                        diag_total += 1
                        ni = np.find(qn('w:numId'))
                        ilvl_el = np.find(qn('w:ilvl'))
                        has_numid = ni is not None and ni.get(qn('w:val'))
                        has_ilvl = ilvl_el is not None
                        if has_numid:
                            diag_has_numid += 1
                            old_nid = ni.get(qn('w:val'))
                            if old_nid in num_id_map:
                                ni.set(qn('w:val'), num_id_map[old_nid])
                                diag_remapped += 1
                            else:
                                diag_missed += 1
                        if has_ilvl:
                            diag_has_ilvl += 1
                        if not has_numid and has_ilvl:
                            diag_ilvl_only += 1
                    if diag_total > 0 and tag == 'tbl':
                        log(f"    [诊断·表格] numPr 共 {diag_total} 个, "
                            f"含 numId: {diag_has_numid}, 含 ilvl: {diag_has_ilvl}, "
                            f"已重映射: {diag_remapped}, 未映射: {diag_missed}, "
                            f"仅 ilvl 无 numId: {diag_ilvl_only}")
                    elif diag_total > 0:
                        log(f"    [诊断·段落] numPr 共 {diag_total} 个, "
                            f"含 numId: {diag_has_numid}, 含 ilvl: {diag_has_ilvl}, "
                            f"已重映射: {diag_remapped}, 未映射: {diag_missed}, "
                            f"仅 ilvl 无 numId: {diag_ilvl_only}")

                dst_body.append(new_el)
                count += 1
            except Exception as e:
                log(f"    [警告] 无法复制元素: {e}")

    log(f"    → 已复制 {count} 个元素（段落+表格）")
    return count


def insert_section_break(dst_doc):
    """
    在目标文档 body 末尾插入一个「下一页分节符」。

    每个分节符都包含：
      - 下一页分页类型
      - 纸张大小锁定为 A4（避免源文档纸张大小不一致导致内容区域宽度异常）
      - 统一页边距（上2cm/下2cm/左1.8cm/右1.8cm）
      - 装订线边距清零（w:gutter="0"，避免源文档装订线导致额外空白）
      - 页眉距离顶部2cm
      - 页脚距离底部1.5cm
      - 基础文档的页眉页脚引用（确保所有 section 共享同一个页眉页脚）
    """
    # Word 使用 twips 单位：1 cm = 567 twips（Word 内部标准）
    # 注意：python-docx 的 Cm() 也使用 twips

    break_p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')

    sectPr = OxmlElement('w:sectPr')
    
    # 分节类型：下一页
    sect_type = OxmlElement('w:type')
    sect_type.set(qn('w:val'), 'nextPage')
    sectPr.append(sect_type)

    # 纸张大小：锁定为 A4（避免源文档纸张大小不一致导致内容区域宽度异常）
    pgSz = OxmlElement('w:pgSz')
    pgSz.set(qn('w:w'), str(int(21 * _TWIPS_PER_CM)))   # A4 宽度  21cm
    pgSz.set(qn('w:h'), str(int(29.7 * _TWIPS_PER_CM))) # A4 高度  29.7cm
    pgSz.set(qn('w:orient'), 'portrait')
    sectPr.append(pgSz)

    # 页边距
    pgMar = OxmlElement('w:pgMar')
    pgMar.set(qn('w:top'), str(int(2 * _TWIPS_PER_CM)))       # 上边距 2cm
    pgMar.set(qn('w:bottom'), str(int(2 * _TWIPS_PER_CM)))     # 下边距 2cm
    pgMar.set(qn('w:left'), str(int(1.8 * _TWIPS_PER_CM)))     # 左边距 1.8cm
    pgMar.set(qn('w:right'), str(int(1.8 * _TWIPS_PER_CM)))    # 右边距 1.8cm
    pgMar.set(qn('w:gutter'), '0')                               # 装订线 = 0（关键！避免右侧额外空白）
    pgMar.set(qn('w:header'), str(int(2 * _TWIPS_PER_CM)))       # 页眉上边距 2cm
    pgMar.set(qn('w:footer'), str(int(1.5 * _TWIPS_PER_CM)))     # 页脚下边距 1.5cm
    sectPr.append(pgMar)

    # 如果已记录基础文档的页眉页脚 rId，则在新分节符中引用它们
    base_hf = getattr(dst_doc, '_base_hf_rids', None)
    if base_hf:
        # headerReference（页眉）
        for h_type, r_id in base_hf.get('headers', {}).items():
            ref = OxmlElement('w:headerReference')
            ref.set(qn('w:type'), h_type)
            ref.set(qn('r:id'), r_id)
            sectPr.append(ref)

        # footerReference（页脚）
        for f_type, r_id in base_hf.get('footers', {}).items():
            ref = OxmlElement('w:footerReference')
            ref.set(qn('w:type'), f_type)
            ref.set(qn('r:id'), r_id)
            sectPr.append(ref)

    pPr.append(sectPr)
    break_p.append(pPr)
    dst_doc.element.body.append(break_p)


def _extract_base_hf_rids(dst_doc):
    """
    从目标文档（基础文档）中提取页眉页脚的 rId，存储到 dst_doc._base_hf_rids。

    这样 insert_section_break() 可以在新分节符中引用这些 rId，
    确保所有 section 共享同一个页眉页脚，页码自动连续。
    """
    base_hf = {'headers': {}, 'footers': {}}

    # 从基础文档的主 sectPr 中提取 headerReference / footerReference
    body_sectPr = _get_body_sectPr(dst_doc)
    if body_sectPr is None:
        return

    # 提取 headerReference
    for ref in body_sectPr.findall(qn('w:headerReference')):
        r_id = ref.get(qn('r:id'))
        h_type = ref.get(qn('w:type'), 'default')
        if r_id:
            base_hf['headers'][h_type] = r_id

    # 提取 footerReference
    for ref in body_sectPr.findall(qn('w:footerReference')):
        r_id = ref.get(qn('r:id'))
        f_type = ref.get(qn('w:type'), 'default')
        if r_id:
            base_hf['footers'][f_type] = r_id

    dst_doc._base_hf_rids = base_hf


def _build_style_map(src_doc, dst_doc):
    """
    构建源文档 → 目标文档的样式 ID 映射表（方案 A：强制使用基础文档的样式体系）。

    映射规则：
      1. 按样式名称（本地名称，如"标题 2"、"Heading 2"）进行匹配
      2. 如果目标文档中有同名样式，则映射到目标的 styleId
      3. 处理中英文 Word 版本差异（"标题 2" ↔ "Heading 2"）

    返回:
        style_id_map: {源 styleId: 目标 styleId}
    """
    # 内置样式名称映射（处理中英文 Word 版本差异）
    BUILTIN_MAP = {
        '标题 1': ['Heading 1', '标题1'],
        '标题 2': ['Heading 2', '标题2'],
        '标题 3': ['Heading 3', '标题3'],
        '标题 4': ['Heading 4', '标题4'],
        '标题 5': ['Heading 5', '标题5'],
        '标题 6': ['Heading 6', '标题6'],
        '标题 7': ['Heading 7', '标题7'],
        '标题 8': ['Heading 8', '标题8'],
        '标题 9': ['Heading 9', '标题9'],
        'Heading 1': ['标题 1', '标题1'],
        'Heading 2': ['标题 2', '标题2'],
        'Heading 3': ['标题 3', '标题3'],
        'Heading 4': ['标题 4', '标题4'],
        'Heading 5': ['标题 5', '标题5'],
        'Heading 6': ['标题 6', '标题6'],
        'Heading 7': ['标题 7', '标题7'],
        'Heading 8': ['标题 8', '标题8'],
        'Heading 9': ['标题 9', '标题9'],
    }

    # 收集目标文档的样式：styleId → 样式名称
    dst_styles = {}
    try:
        dst_styles_el = dst_doc.styles.element
        for s in dst_styles_el.findall(qn('w:style')):
            sid = s.get(qn('w:styleId'))
            # 获取样式名称（本地名称）
            name_el = s.find(qn('w:name'))
            sname = name_el.get(qn('w:val')) if name_el is not None else sid
            if sid:
                dst_styles[sid] = sname
    except Exception:
        pass

    # 收集源文档的样式：styleId → 样式名称
    src_styles = {}
    try:
        src_styles_el = src_doc.styles.element
        for s in src_styles_el.findall(qn('w:style')):
            sid = s.get(qn('w:styleId'))
            name_el = s.find(qn('w:name'))
            sname = name_el.get(qn('w:val')) if name_el is not None else sid
            if sid:
                src_styles[sid] = sname
    except Exception:
        pass

    # 构建映射：源 styleId → 目标 styleId
    style_id_map = {}

    for src_sid, src_sname in src_styles.items():
        # 尝试在目标文档中找到同名样式
        matched_dst_sid = None

        # 方法1：直接匹配样式名称
        for dst_sid, dst_sname in dst_styles.items():
            if src_sname == dst_sname:
                matched_dst_sid = dst_sid
                break

        # 方法2：使用内置映射表（处理中英文差异）
        if matched_dst_sid is None:
            for canonical, aliases in BUILTIN_MAP.items():
                if src_sname == canonical or src_sname in aliases:
                    # 在目标文档中查找 canonical 或 aliases
                    for dst_sid, dst_sname in dst_styles.items():
                        if dst_sname == canonical or dst_sname in aliases:
                            matched_dst_sid = dst_sid
                            break
                    if matched_dst_sid:
                        break

        # 如果找到匹配，建立映射
        if matched_dst_sid:
            style_id_map[src_sid] = matched_dst_sid

    return style_id_map


def _apply_style_map(element, style_id_map):
    """
    递归地将 element 中所有 w:pStyle / w:tblStyle 的 w:val
    替换为映射后的 styleId。
    """
    if not style_id_map:
        return

    # 处理段落样式引用
    for pStyle in element.findall('.//' + qn('w:pStyle')):
        old_val = pStyle.get(qn('w:val'))
        if old_val and old_val in style_id_map:
            pStyle.set(qn('w:val'), style_id_map[old_val])


def _merge_numbering(src_doc, dst_doc, log=None):
    """
    将源文档 numbering.xml 中缺失的列表格式定义复制到目标文档。

    背景：clone_element_to_doc 会复制段落的 w:numPr（含 numId），
    但目标文档 numbering.xml 可能没有对应的抽象编号格式定义，
    导致自定义编号（如"项目 1"）变成默认 bullet "•"。

    做法：
      1. 遍历源文档 abstractNum，目标缺失的则 deepcopy 追加
      2. 遍历源文档 num，处理 numId 冲突（冲突时分配新 ID）
      3. 返回 numId 重映射表 {src_numId: dst_numId}

    不影响已写入 numId='0' 的标题段落（numId='0' 不参与任何列表）。
    """
    from lxml import etree
    import copy as _copy_module

    def _l(msg):
        if log:
            log(msg)

    try:
        dst_num_el = dst_doc.part.numbering_part._element
        src_num_el = src_doc.part.numbering_part._element
    except Exception:
        return {}

    # 收集目标已有的 abstractNum ID → 元素映射（用于内容比较）
    existing_abs = {}  # aid -> element
    max_abs_id = 0
    for an in dst_num_el.findall(qn('w:abstractNum')):
        aid = an.get(qn('w:abstractNumId'))
        if aid:
            existing_abs[aid] = an
            max_abs_id = max(max_abs_id, int(aid))

    # ---- 诊断辅助：提取 abstractNum 中各级别的编号格式 ----
    def _extract_numfmts(an_el):
        """提取 abstractNum 中所有 lvl 的 numFmt，用于诊断。"""
        fmts = []
        for lvl in an_el.findall(qn('w:lvl')):
            ilvl = lvl.get(qn('w:ilvl'), '?')
            nf = lvl.find(qn('w:numFmt'))
            fmt_val = nf.get(qn('w:val')) if nf is not None else '?'
            fmts.append(f"lvl{ilvl}={fmt_val}")
        return ', '.join(fmts) if fmts else '(无 lvl)'

    # 复制 abstractNum：ID 不存在则直接复制；存在但内容不同则分配新 ID 复制
    abs_id_map = {}   # src_abstractNumId -> dst_abstractNumId
    copied_abs = 0
    for an in src_num_el.findall(qn('w:abstractNum')):
        aid = an.get(qn('w:abstractNumId'))
        if not aid:
            continue
        src_fmts = _extract_numfmts(an)
        if aid not in existing_abs:
            # 直接复制（无冲突）
            dst_num_el.append(_copy_module.deepcopy(an))
            existing_abs[aid] = an
            abs_id_map[aid] = aid
            copied_abs += 1
            _l(f"    [诊断·编号] abstractNumId={aid}（新增） 格式: {src_fmts}")
        else:
            # ID 冲突：比较 XML 内容是否相同
            dst_an = existing_abs[aid]
            dst_fmts = _extract_numfmts(dst_an)
            if etree.tostring(an) != etree.tostring(dst_an):
                # 内容不同：分配新 abstractNumId
                max_abs_id += 1
                new_aid = str(max_abs_id)
                new_an = _copy_module.deepcopy(an)
                new_an.set(qn('w:abstractNumId'), new_aid)
                dst_num_el.append(new_an)
                existing_abs[new_aid] = new_an
                abs_id_map[aid] = new_aid
                copied_abs += 1
                _l(f"    [诊断·编号] abstractNumId={aid}→{new_aid}（内容冲突） 源格式: {src_fmts} | 目标格式: {dst_fmts}")
            else:
                abs_id_map[aid] = aid
                _l(f"    [诊断·编号] abstractNumId={aid}（已存在且相同） 格式: {src_fmts}")

    # ---- 辅助：读写 w:num 下 w:abstractNumId 子元素的 w:val ----
    # 注意：w:abstractNumId 是 w:num 的【子元素】，不是属性！
    # 结构：<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    def _get_num_abs_id(num_el):
        abs_el = num_el.find(qn('w:abstractNumId'))
        return abs_el.get(qn('w:val')) if abs_el is not None else None

    def _set_num_abs_id(num_el, abs_id):
        abs_el = num_el.find(qn('w:abstractNumId'))
        if abs_el is not None:
            abs_el.set(qn('w:val'), abs_id)

    # 收集目标已有的 numId → abstractNumId 映射
    existing_nums = {}
    max_num_id = 0
    for n in dst_num_el.findall(qn('w:num')):
        nid_str = n.get(qn('w:numId'))
        if nid_str:
            max_num_id = max(max_num_id, int(nid_str))
            existing_nums[nid_str] = _get_num_abs_id(n)

    # 复制源文档的 num，处理 numId 冲突 + abstractNumId 重映射
    num_id_map = {}  # src_numId -> dst_numId
    copied_nums = 0
    for n in src_num_el.findall(qn('w:num')):
        src_nid = n.get(qn('w:numId'))
        if not src_nid:
            continue
        src_abs_id = _get_num_abs_id(n)
        new_abs_id = abs_id_map.get(src_abs_id, src_abs_id) if src_abs_id else src_abs_id
        if src_nid in existing_nums:
            # numId 冲突：分配新 ID
            max_num_id += 1
            new_nid = str(max_num_id)
            new_n = _copy_module.deepcopy(n)
            new_n.set(qn('w:numId'), new_nid)
            if new_abs_id:
                _set_num_abs_id(new_n, new_abs_id)
            dst_num_el.append(new_n)
            num_id_map[src_nid] = new_nid
            existing_nums[new_nid] = new_abs_id
            _l(f"    [诊断·numId] numId={src_nid}→{new_nid}（冲突） absId={src_abs_id}→{new_abs_id}")
        else:
            new_n = _copy_module.deepcopy(n)
            if new_abs_id:
                _set_num_abs_id(new_n, new_abs_id)
            dst_num_el.append(new_n)
            existing_nums[src_nid] = new_abs_id
            if new_abs_id != src_abs_id:
                _l(f"    [诊断·numId] numId={src_nid}（无冲突，仅重映射 absId: {src_abs_id}→{new_abs_id}")
        copied_nums += 1

    if copied_abs > 0 or copied_nums > 0:
        _l(f"    → 已合并 numbering：{copied_abs} 个列表格式, "
           f"{copied_nums} 个列表实例"
           + (f"（{len(num_id_map)} 个 numId 冲突已重映射）" if num_id_map else ""))

    return num_id_map
