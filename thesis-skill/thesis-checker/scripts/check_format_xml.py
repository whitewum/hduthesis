#!/usr/bin/env python3
"""
毕业论文格式检查脚本（XML模式）
基于 docx skill 的 unpack 方法，直接解析 Word XML，无需第三方依赖。
杭州电子科技大学信息工程学院本科毕业设计格式规范
"""

import sys
import json
import re
import os
from pathlib import Path
from xml.etree import ElementTree as ET

# Word XML 命名空间
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}

# ─── 字号换算 ─────────────────────────────────────────────────
# Word XML 中字号单位是半磅（half-points），即 24 = 12pt = 小四
FONT_SIZES_HALF = {
    "28号": 56,
    "小三": 30,
    "三号": 32,
    "四号": 28,
    "小四": 24,
    "五号": 21,
}
TOLERANCE_HALF = 1  # 半磅误差容忍

def half_to_pt(h): return h / 2.0
def pt_to_half(p): return int(p * 2)


# ─── XML 辅助函数 ─────────────────────────────────────────────
def w(tag): return f'{{{NS["w"]}}}{tag}'


def get_rpr_val(rPr, tag):
    """从 <w:rPr> 中获取某个元素的 w:val 属性"""
    if rPr is None:
        return None
    el = rPr.find(w(tag))
    if el is None:
        return None
    return el.get(w('val'))


def get_run_font_size(rPr):
    """从 <w:rPr> 读取字号（半磅）"""
    if rPr is None:
        return None
    sz = rPr.find(w('sz'))
    if sz is not None:
        val = sz.get(w('val'))
        if val:
            return int(val)
    return None


def get_run_east_asian_font(rPr):
    """读取中文字体（w:eastAsia）"""
    if rPr is None:
        return None
    rFonts = rPr.find(w('rFonts'))
    if rFonts is None:
        return None
    return rFonts.get(w('eastAsia'))


def get_run_ascii_font(rPr):
    """读取西文字体（w:ascii）"""
    if rPr is None:
        return None
    rFonts = rPr.find(w('rFonts'))
    if rFonts is None:
        return None
    return rFonts.get(w('ascii'))


def get_para_text(para):
    """获取段落纯文本"""
    parts = []
    for r in para.findall('.//' + w('t')):
        parts.append(r.text or '')
    return ''.join(parts)


def get_pPr(para):
    return para.find(w('pPr'))


def get_alignment(pPr):
    """读取对齐方式"""
    if pPr is None:
        return None
    jc = pPr.find(w('jc'))
    if jc is None:
        return None
    return jc.get(w('val'))  # 'center', 'left', 'right', 'both'


def get_spacing(pPr):
    """读取行距设置，返回 (lineRule, line_half_pt)"""
    if pPr is None:
        return None, None
    spacing = pPr.find(w('spacing'))
    if spacing is None:
        return None, None
    line_rule = spacing.get(w('lineRule'))  # 'exact', 'atLeast', 'auto'
    line = spacing.get(w('line'))           # 240 = 单倍, exact值
    return line_rule, int(line) if line else None


def get_space_before_after(pPr):
    """读取段前段后（TWIPs: 1磅=20 TWIPs）"""
    if pPr is None:
        return None, None
    spacing = pPr.find(w('spacing'))
    if spacing is None:
        return None, None
    before = spacing.get(w('before'))
    after = spacing.get(w('after'))
    return (int(before) if before else None), (int(after) if after else None)


def get_first_line_indent(pPr):
    """读取首行缩进（TWIPs）"""
    if pPr is None:
        return None
    ind = pPr.find(w('ind'))
    if ind is None:
        return None
    fi = ind.get(w('firstLine'))
    return int(fi) if fi else None


def get_para_style(pPr):
    """读取段落样式名"""
    if pPr is None:
        return None
    pStyle = pPr.find(w('pStyle'))
    if pStyle is None:
        return None
    return pStyle.get(w('val'))


def para_rpr(para):
    """获取段落下第一个非空 run 的 rPr（优先段落级 rPr）"""
    pPr = get_pPr(para)
    if pPr is not None:
        rPr = pPr.find(w('rPr'))
        if rPr is not None:
            return rPr
    for r in para.findall(w('r')):
        rPr = r.find(w('rPr'))
        t = r.find(w('t'))
        if rPr is not None and t is not None and (t.text or '').strip():
            return rPr
    return None


# ─── 标签归一化（处理 "摘 要" / "目　录" / 大小写 等变体） ───────
def normalize_label(s):
    """去掉所有空白字符（半角空格、全角空格、制表符等）并转小写"""
    if s is None:
        return ''
    return re.sub(r'\s+', '', s).lower()


def matches_label(text, label):
    """判断 text 是否就是 label（忽略中间空格 / 大小写）"""
    return normalize_label(text) == normalize_label(label)


# ─── 样式继承：解析 styles.xml 的 pPr / rPr 并提供属性回溯 ─────
def _build_style_table(styles_xml_path):
    """
    返回 (styles_dict, default_pPr, default_rPr)
        styles_dict[styleId] = {'name': str, 'basedOn': str|None,
                                'pPr': Element|None, 'rPr': Element|None}
    """
    styles = {}
    default_pPr = None
    default_rPr = None
    if not Path(styles_xml_path).exists():
        return styles, default_pPr, default_rPr
    tree = ET.parse(str(styles_xml_path))
    root = tree.getroot()
    # docDefaults
    dd = root.find(w('docDefaults'))
    if dd is not None:
        pPrD = dd.find(w('pPrDefault'))
        if pPrD is not None:
            default_pPr = pPrD.find(w('pPr'))
        rPrD = dd.find(w('rPrDefault'))
        if rPrD is not None:
            default_rPr = rPrD.find(w('rPr'))
    # styles
    for st in root.findall(w('style')):
        sid = st.get(w('styleId'))
        if not sid:
            continue
        name_el = st.find(w('name'))
        name = name_el.get(w('val')) if name_el is not None else ''
        basedOn_el = st.find(w('basedOn'))
        basedOn = basedOn_el.get(w('val')) if basedOn_el is not None else None
        styles[sid] = {
            'name': name or '',
            'basedOn': basedOn,
            'pPr': st.find(w('pPr')),
            'rPr': st.find(w('rPr')),
        }
    return styles, default_pPr, default_rPr


def _walk_style_chain(styles, style_id):
    """生成器：依次产出 style_id, style_id.basedOn, ..."""
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = styles.get(style_id)
        if not st:
            return
        yield st
        style_id = st.get('basedOn')


def get_pPr_chain(para, styles, default_pPr):
    """段落 pPr 的继承链：段落 → 样式链 → docDefaults"""
    chain = []
    pPr = para.find(w('pPr'))
    style_id = None
    if pPr is not None:
        chain.append(pPr)
        pStyle = pPr.find(w('pStyle'))
        if pStyle is not None:
            style_id = pStyle.get(w('val'))
    for st in _walk_style_chain(styles, style_id):
        if st['pPr'] is not None:
            chain.append(st['pPr'])
    if default_pPr is not None:
        chain.append(default_pPr)
    return chain


def get_rPr_chain(para, styles, default_rPr):
    """run rPr 继承链：第一个有内容 run 的 rPr → 段落级 rPr → 样式链 → docDefaults"""
    chain = []
    # 第一个非空 run 的 rPr
    for r in para.findall(w('r')):
        rPr = r.find(w('rPr'))
        t = r.find(w('t'))
        if rPr is not None and t is not None and (t.text or '').strip():
            chain.append(rPr)
            break
    pPr = para.find(w('pPr'))
    style_id = None
    if pPr is not None:
        para_rPr_el = pPr.find(w('rPr'))
        if para_rPr_el is not None and para_rPr_el not in chain:
            chain.append(para_rPr_el)
        pStyle = pPr.find(w('pStyle'))
        if pStyle is not None:
            style_id = pStyle.get(w('val'))
    for st in _walk_style_chain(styles, style_id):
        if st['rPr'] is not None:
            chain.append(st['rPr'])
    if default_rPr is not None:
        chain.append(default_rPr)
    return chain


def resolve_alignment(chain):
    """从继承链获取对齐方式，找到第一个 <w:jc w:val="..."/>"""
    for pPr in chain:
        jc = pPr.find(w('jc'))
        if jc is not None:
            v = jc.get(w('val'))
            if v:
                return v
    return None


def resolve_spacing(chain):
    """从继承链获取行距（line, lineRule），属性级回溯"""
    rule = None
    line = None
    for pPr in chain:
        sp = pPr.find(w('spacing'))
        if sp is None:
            continue
        if rule is None:
            rule = sp.get(w('lineRule'))
        if line is None:
            v = sp.get(w('line'))
            if v:
                line = int(v)
        if rule is not None and line is not None:
            break
    return rule, line


def resolve_first_line_indent(chain):
    """读取首行缩进（TWIPs），属性级回溯。优先 firstLine；其次按 firstLineChars 推算"""
    for pPr in chain:
        ind = pPr.find(w('ind'))
        if ind is None:
            continue
        fi = ind.get(w('firstLine'))
        if fi:
            return int(fi)
        fic = ind.get(w('firstLineChars'))
        if fic:
            # firstLineChars 单位为 1/100 字符，按 12pt 字号 1 字符≈240 TWIPs 估算
            return int(int(fic) / 100 * 240)
    return None


def resolve_run_font_size(chain):
    """从 rPr 链读取字号（半磅）"""
    for rPr in chain:
        sz = rPr.find(w('sz'))
        if sz is not None:
            v = sz.get(w('val'))
            if v:
                return int(v)
    return None


def resolve_run_east_asia_font(chain):
    """从 rPr 链读取中文字体（w:eastAsia）"""
    for rPr in chain:
        rFonts = rPr.find(w('rFonts'))
        if rFonts is not None:
            v = rFonts.get(w('eastAsia'))
            if v:
                return v
    return None


def resolve_run_ascii_font(chain):
    """从 rPr 链读取西文字体（w:ascii）"""
    for rPr in chain:
        rFonts = rPr.find(w('rFonts'))
        if rFonts is not None:
            v = rFonts.get(w('ascii'))
            if v:
                return v
    return None


def resolve_run_bold(chain):
    """从 rPr 链判断是否加粗"""
    for rPr in chain:
        b = rPr.find(w('b'))
        if b is not None:
            v = b.get(w('val'))
            # 默认有 <w:b/> 即加粗，除非 val=0/false
            if v in (None, '1', 'true', 'on'):
                return True
            return False
    return False


# ─── 主检查器 ─────────────────────────────────────────────────
class ThesisXMLChecker:
    def __init__(self, unpacked_dir):
        self.unpacked = Path(unpacked_dir)
        self.doc_xml = self.unpacked / 'word' / 'document.xml'
        self.styles_xml = self.unpacked / 'word' / 'styles.xml'
        self.issues_critical = []
        self.issues_warning = []
        self.passed = []
        self.content_notes = []

        # 解析主文档
        tree = ET.parse(self.doc_xml)
        self.root = tree.getroot()
        body = self.root.find('.//' + w('body'))
        # 用 recursive find，覆盖 <w:sdt>（TOC/SDT 容器）、<w:tbl> 表格、文本框等嵌套段落
        self.paragraphs = body.findall('.//' + w('p')) if body is not None else []

        # 解析样式表（含 pPr/rPr，用于属性继承回溯）
        self.style_table, self.default_pPr, self.default_rPr = _build_style_table(self.styles_xml)
        # 兼容旧字段：styleId -> styleName
        self.styles = {sid: meta['name'] for sid, meta in self.style_table.items()}

    def c(self, msg): self.issues_critical.append(msg)
    def w_(self, msg): self.issues_warning.append(msg)
    def p(self, msg): self.passed.append(msg)

    # ─── 继承链便捷方法 ──────────────────────────────────────
    def pPr_chain(self, para):
        return get_pPr_chain(para, self.style_table, self.default_pPr)

    def rPr_chain(self, para):
        return get_rPr_chain(para, self.style_table, self.default_rPr)

    def heading_level(self, style_id):
        """根据样式ID判断标题级别（1/2/3/None）"""
        if style_id is None:
            return None
        # 直接ID匹配
        if style_id in ('Heading1', '1', 'heading1', '标题1', '1级标题'):
            return 1
        if style_id in ('Heading2', '2', 'heading2', '标题2', '2级标题'):
            return 2
        if style_id in ('Heading3', '3', 'heading3', '标题3', '3级标题'):
            return 3
        # 通过样式名匹配
        sname = self.styles.get(style_id, '').lower()
        if 'heading 1' in sname or '标题 1' in sname:
            return 1
        if 'heading 2' in sname or '标题 2' in sname:
            return 2
        if 'heading 3' in sname or '标题 3' in sname:
            return 3
        return None

    # ─── 页眉检查 ─────────────────────────────────────────────
    def check_header(self):
        # 允许「本科毕业论文」或「本科毕业设计」，允许其间含空格
        header_pattern = re.compile(
            r'杭州电子科技大学信息工程学院本科毕业(论文|设计)'
        )
        header_files = list((self.unpacked / 'word').glob('header*.xml'))
        if not header_files:
            self.c("[页眉] 未找到页眉文件，封面后所有页面须有页眉")
            return

        # 跨多个 header 文件去重，同一问题只报一次
        seen_keys = set()
        any_text_found = False

        for hf in sorted(header_files):
            tree = ET.parse(str(hf))
            root = tree.getroot()
            for para in root.findall('.//' + w('p')):
                text = get_para_text(para)
                if not text.strip():
                    continue
                any_text_found = True

                # 归一化：去掉所有空白后做正则匹配
                norm = re.sub(r'\s+', '', text)
                key_content = ('content', norm)
                if not header_pattern.search(norm):
                    if ('content_bad', norm) not in seen_keys:
                        seen_keys.add(('content_bad', norm))
                        self.c(
                            f"[页眉] 内容不正确，应为「杭州电子科技大学信息工程学院本科毕业论文」"
                            f"或「杭州电子科技大学信息工程学院本科毕业设计」，当前为「{text.strip()}」"
                        )
                else:
                    if 'content_ok' not in seen_keys:
                        seen_keys.add('content_ok')
                        self.p("页眉内容正确（含「本科毕业论文/设计」）")

                # 对齐：使用样式继承
                align = resolve_alignment(self.pPr_chain(para))
                if align != 'center':
                    if 'align_bad' not in seen_keys:
                        seen_keys.add('align_bad')
                        self.c(f"[页眉] 未居中，当前对齐方式={align}")
                else:
                    if 'align_ok' not in seen_keys:
                        seen_keys.add('align_ok')
                        self.p("页眉居中对齐")

                # 字号 / 字体：使用 rPr 继承链
                rchain = self.rPr_chain(para)
                sz = resolve_run_font_size(rchain)
                if sz and abs(sz - FONT_SIZES_HALF["五号"]) > TOLERANCE_HALF:
                    if 'sz_bad' not in seen_keys:
                        seen_keys.add('sz_bad')
                        self.w_(f"[页眉] 字号应为五号(10.5pt/21半磅)，检测到 {half_to_pt(sz)}pt")
                elif sz and 'sz_ok' not in seen_keys:
                    seen_keys.add('sz_ok')
                    self.p("页眉字号正确（五号 10.5pt）")

                ea = resolve_run_east_asia_font(rchain)
                if ea and '宋体' not in ea and 'SimSun' not in ea:
                    if 'ea_bad' not in seen_keys:
                        seen_keys.add('ea_bad')
                        self.w_(f"[页眉] 中文字体应为宋体，检测到「{ea}」")
                break  # 同一 header 文件只检查首个非空段落
        if not any_text_found:
            self.c("[页眉] 所有页眉文件内容均为空，封面后页面应有页眉文字")

    # ─── 结构完整性 ───────────────────────────────────────────
    def check_structure(self):
        # 章节关键词支持别名（"参考文献" 也可能写作 "References"）
        required = {
            "摘要": ["摘要"],
            "Abstract": ["Abstract"],
            "目录": ["目录"],
            "参考文献": ["参考文献", "References"],
        }
        found = {k: False for k in required}
        for para in self.paragraphs:
            t = get_para_text(para).strip()
            if not t:
                continue
            for canon, aliases in required.items():
                if any(matches_label(t, a) for a in aliases):
                    found[canon] = True
        missing = [k for k, v in found.items() if not v]
        if missing:
            self.c(f"[结构] 缺少必要章节：{'、'.join(missing)}")
        else:
            self.p("论文结构完整（含摘要、Abstract、目录、参考文献）")

    # ─── 摘要格式 ─────────────────────────────────────────────
    def check_abstract(self):
        for para in self.paragraphs:
            text = get_para_text(para).strip()
            if not text:
                continue
            pchain = self.pPr_chain(para)
            rchain = self.rPr_chain(para)

            # "摘要" / "摘 要" / "摘　要" 都识别
            if matches_label(text, '摘要'):
                align = resolve_alignment(pchain)
                if align != 'center':
                    self.c("[摘要] 标题应居中")
                else:
                    self.p("摘要标题居中")
                sz = resolve_run_font_size(rchain)
                if sz and abs(sz - FONT_SIZES_HALF["三号"]) > TOLERANCE_HALF:
                    self.w_(f"[摘要] 字号应为三号(16pt)，检测到 {half_to_pt(sz)}pt")
                elif sz:
                    self.p("摘要标题字号正确（三号 16pt）")
                ea = resolve_run_east_asia_font(rchain)
                if ea and '黑体' not in ea:
                    self.w_(f"[摘要] 标题字体应为黑体，检测到「{ea}」")

            if matches_label(text, 'Abstract'):
                align = resolve_alignment(pchain)
                if align != 'center':
                    self.c("[英文摘要] Abstract 标题应居中")
                sz = resolve_run_font_size(rchain)
                if sz and abs(sz - FONT_SIZES_HALF["三号"]) > TOLERANCE_HALF:
                    self.w_(f"[英文摘要] 字号应为三号(16pt)，检测到 {half_to_pt(sz)}pt")
                # 加粗（沿继承链查找）
                if not resolve_run_bold(rchain):
                    self.w_("[英文摘要] Abstract 标题应加粗")

            # 目录
            if matches_label(text, '目录'):
                align = resolve_alignment(pchain)
                if align != 'center':
                    self.c("[目录] 标题应居中")
                else:
                    self.p("目录标题居中")
                sz = resolve_run_font_size(rchain)
                if sz and abs(sz - FONT_SIZES_HALF["三号"]) > TOLERANCE_HALF:
                    self.w_(f"[目录] 字号应为三号(16pt)，检测到 {half_to_pt(sz)}pt")
                elif sz:
                    self.p("目录标题字号正确（三号 16pt）")
                ea = resolve_run_east_asia_font(rchain)
                if ea and '黑体' not in ea:
                    self.w_(f"[目录] 标题字体应为黑体，检测到「{ea}」")

            if text.startswith('关键词') or text.startswith('关键字') or text.lower().startswith('keywords'):
                ind = resolve_first_line_indent(pchain)
                if ind and ind > 0:
                    self.w_("[摘要] 关键词行应顶格，不要首行缩进")
                else:
                    self.p("关键词顶格书写")
                if '；' in text or ';' in text:
                    self.p("关键词使用分号分割")
                else:
                    self.w_("[摘要] 关键词应使用分号（；）分割")

    # ─── 标题格式 ─────────────────────────────────────────────
    def check_headings(self):
        counts = {1: 0, 2: 0, 3: 0}
        expected_sz = {1: FONT_SIZES_HALF["三号"], 2: FONT_SIZES_HALF["四号"], 3: FONT_SIZES_HALF["小四"]}
        # 一级标题居中检查的"全部通过"汇总
        h1_align_ok = 0
        h1_align_bad = 0

        for para in self.paragraphs:
            pPr = get_pPr(para)
            style_id = get_para_style(pPr)
            level = self.heading_level(style_id)
            if level not in (1, 2, 3):
                continue

            text = get_para_text(para).strip()
            if not text:
                continue
            counts[level] += 1
            prefix = f"[{level}级标题「{text[:20]}」]"

            # 使用属性继承链
            pchain = self.pPr_chain(para)
            rchain = self.rPr_chain(para)

            # 字号
            sz = resolve_run_font_size(rchain)
            if sz is not None:
                if abs(sz - expected_sz[level]) > TOLERANCE_HALF:
                    self.w_(f"{prefix} 字号应为 {half_to_pt(expected_sz[level])}pt，检测到 {half_to_pt(sz)}pt")

            # 中文字体（黑体）
            ea = resolve_run_east_asia_font(rchain)
            if ea and '黑体' not in ea and 'SimHei' not in ea:
                self.w_(f"{prefix} 中文字体应为黑体，检测到「{ea}」")

            # 对齐 — 一级标题居中
            align = resolve_alignment(pchain)
            rule, line = resolve_spacing(pchain)
            if level == 1:
                if align != 'center':
                    h1_align_bad += 1
                    self.c(f"{prefix} 一级标题应居中，当前={align}")
                else:
                    h1_align_ok += 1
                # 行距（单倍 lineRule=auto，line=240）
                if rule == 'exact':
                    self.w_(f"{prefix} 一级标题行距应为单倍行距，不应用固定值")
            else:
                if align == 'center':
                    self.c(f"{prefix} {level}级标题应居左顶格，不能居中")
                # 行距（固定值20磅 = 400 twips, lineRule=exact）
                if rule != 'exact':
                    self.w_(f"{prefix} 行距应为固定值20磅（lineRule=exact），当前 rule={rule}")
                elif line and abs(line - 400) > 10:
                    self.c(f"{prefix} 固定行距应为20磅(400 twips)，当前={line} twips({line/20:.1f}pt)")

            # 编号格式
            if level == 1:
                if re.match(r'^[一二三四五六七八九十]', text):
                    self.w_(f"{prefix} 编号应为阿拉伯数字(1,2,3...)，不要用中文数字")
            elif level == 2:
                if not re.match(r'^\d+\.\d+', text):
                    self.w_(f"{prefix} 编号格式应为 1.1, 1.2...，请使用多级列表自动编号")
            elif level == 3:
                if not re.match(r'^\d+\.\d+\.\d+', text):
                    self.w_(f"{prefix} 编号格式应为 1.1.1...，请使用多级列表自动编号")

        if counts[1] > 0:
            self.p(f"检测到 {counts[1]} 个一级标题，{counts[2]} 个二级标题，{counts[3]} 个三级标题")
            if h1_align_bad == 0 and h1_align_ok > 0:
                self.p(f"全部 {h1_align_ok} 个一级标题居中（含样式继承）")
        else:
            self.w_("[标题] 未检测到标题样式段落，请确认使用了「标题1/2/3」样式")

    # ─── 正文格式 ─────────────────────────────────────────────
    def _is_caption_or_label(self, text):
        """图题、表题、关键词、参考文献条目等应排除在正文统计外"""
        if not text:
            return True
        t = text.strip()
        if re.match(r'^(图|表)\s*\d+[-－]?\d*', t):
            return True
        if t.startswith('关键词') or t.startswith('关键字') or t.lower().startswith('keywords'):
            return True
        if matches_label(t, '摘要') or matches_label(t, 'Abstract'):
            return True
        if matches_label(t, '目录') or matches_label(t, '参考文献'):
            return True
        # 参考文献条目：通常以 [N] 开头
        if re.match(r'^\[\d+\]', t):
            return True
        return False

    @staticmethod
    def _is_chinese_dominant(text, threshold=0.3):
        """中文字符占比是否够高 — 用于判别是否是中文正文段落
        代码块、纯 ASCII 数据/JSON、程序输出等均为非中文主导，应排除。
        """
        if not text:
            return False
        zh = sum(1 for ch in text if '一' <= ch <= '鿿')
        # 仅统计有意义的可见字符
        total = sum(1 for ch in text if not ch.isspace())
        if total == 0:
            return False
        return zh / total >= threshold

    @staticmethod
    def _looks_like_code(text):
        """启发式判断是否为代码 / 数据片段（即便包含少量中文注释也算）"""
        if not text:
            return False
        t = text.strip()
        # 典型代码关键字 / 符号
        code_keywords = (
            'def ', 'class ', 'import ', 'from ', 'return ', 'self.', 'lambda ',
            'function ', 'const ', 'var ', 'let ', 'public ', 'private ',
            '#include', 'std::', '->', '=>', '!==', '===',
        )
        if any(k in t for k in code_keywords):
            return True
        # 大量代码符号
        code_punct = sum(t.count(c) for c in '{}();[]<>=')
        if code_punct >= 3 and len(t) < 200:
            return True
        return False

    def check_body_text(self):
        # 仅统计"主体"正文：第一个一级标题之后、参考文献之前
        body_paras = []
        seen_first_heading = False
        in_references = False
        for p in self.paragraphs:
            pPr = get_pPr(p)
            style_id = get_para_style(pPr)
            level = self.heading_level(style_id)
            text = get_para_text(p).strip()

            if level == 1:
                seen_first_heading = True
                # 进入参考文献章节后停止收集
                if matches_label(text, '参考文献') or matches_label(text, 'References'):
                    in_references = True
                    continue
                # 一级标题本身不算正文
                continue
            if level in (2, 3):
                continue  # 二/三级标题不算正文
            if not seen_first_heading or in_references:
                continue
            if len(text) <= 10:
                continue
            if self._is_caption_or_label(text):
                continue
            # 跳过代码块 / 数据片段 / 纯 ASCII 段落（非中文正文）
            if self._looks_like_code(text):
                continue
            if not self._is_chinese_dominant(text):
                continue
            # 仅接受 Normal / 正文 / 无样式 的段落
            if style_id is not None and style_id not in (
                'Normal', 'normal', '正文', 'BodyText', 'Body Text', 'a',
            ):
                # 允许样式名（非 ID）也判定
                style_name = (self.styles.get(style_id) or '').lower()
                if not any(k in style_name for k in ('normal', 'body', '正文')):
                    continue
            body_paras.append(p)

        # 兜底：若收集为空（譬如没有标题样式），退回旧逻辑
        if not body_paras:
            for p in self.paragraphs:
                pPr = get_pPr(p)
                style_id = get_para_style(pPr)
                if self.heading_level(style_id) is not None:
                    continue
                text = get_para_text(p).strip()
                if (len(text) > 10
                        and not self._is_caption_or_label(text)
                        and not self._looks_like_code(text)
                        and self._is_chinese_dominant(text)):
                    body_paras.append(p)

        if not body_paras:
            self.w_("[正文] 未检测到正文样式段落")
            return

        # 取靠中部的样本，避开开头的封面/摘要末尾
        sample_size = min(20, len(body_paras))
        start = max(0, (len(body_paras) - sample_size) // 2)
        sample = body_paras[start:start + sample_size]

        sz_issues = spacing_issues = indent_issues = font_issues = 0
        sz_examples = []
        spacing_examples = []
        indent_examples = []

        for para in sample:
            pchain = self.pPr_chain(para)
            rchain = self.rPr_chain(para)

            # 字号（小四 = 24 半磅）
            sz = resolve_run_font_size(rchain)
            if sz and abs(sz - FONT_SIZES_HALF["小四"]) > TOLERANCE_HALF:
                sz_issues += 1
                if len(sz_examples) < 3:
                    sz_examples.append(f"{half_to_pt(sz)}pt")

            # 行距（固定值20磅 = 400 twips, lineRule=exact）— 通过继承链解析
            rule, line = resolve_spacing(pchain)
            if rule != 'exact' or (line is not None and abs(line - 400) > 10):
                spacing_issues += 1
                if len(spacing_examples) < 3:
                    spacing_examples.append(f"rule={rule}, line={line}")

            # 首行缩进（2字符 ≈ 480 twips）
            ind = resolve_first_line_indent(pchain)
            if ind is not None and ind < 300:
                indent_issues += 1
                if len(indent_examples) < 3:
                    indent_examples.append(f"{ind} twips")

            # 中文字体
            ea = resolve_run_east_asia_font(rchain)
            if ea and '宋体' not in ea and 'SimSun' not in ea:
                font_issues += 1

        n = len(sample)
        if sz_issues > n * 0.3:
            self.w_(
                f"[正文] 字号不符合要求（应为小四12pt），约{sz_issues}/{n}个段落有问题"
                + (f"，例：{', '.join(sz_examples)}" if sz_examples else '')
            )
        else:
            self.p(f"正文字号检查通过（小四 12pt，{n}个抽样段落）")

        if spacing_issues > n * 0.3:
            self.c(
                f"[正文] 行距不符合要求（应为固定值20磅），约{spacing_issues}/{n}个段落有问题"
                + (f"，例：{'; '.join(spacing_examples)}" if spacing_examples else '')
            )
        else:
            self.p(f"正文行距检查通过（固定值20磅，{n}个抽样段落）")

        if indent_issues > n * 0.3:
            self.w_(
                f"[正文] 首行缩进可能不足2字符，约{indent_issues}/{n}个段落"
                + (f"，例：{', '.join(indent_examples)}" if indent_examples else '')
            )
        else:
            self.p("正文首行缩进检查通过")

        if font_issues > n * 0.3:
            self.w_(f"[正文] 部分正文中文字体非宋体（约{font_issues}/{n}处），请人工确认")

    # ─── 图表检查 ─────────────────────────────────────────────
    def check_figures_tables(self):
        all_text = '\n'.join(get_para_text(p) for p in self.paragraphs)

        fig_refs = re.findall(r'如图\s*\d+[-－]\d+\s*所示', all_text)
        tab_refs = re.findall(r'如表\s*\d+[-－]\d+\s*所示', all_text)
        fig_captions = [get_para_text(p).strip() for p in self.paragraphs
                        if re.match(r'^图\s*\d+[-－]\d+', get_para_text(p).strip())]
        tab_captions = [get_para_text(p).strip() for p in self.paragraphs
                        if re.match(r'^表\s*\d+[-－]\d+', get_para_text(p).strip())]

        if fig_captions:
            self.p(f"检测到 {len(fig_captions)} 个图题（图X-Y格式）")
            # 图题字号检查
            for para in self.paragraphs:
                t = get_para_text(para).strip()
                if re.match(r'^图\s*\d+[-－]\d+', t):
                    rPr = para_rpr(para)
                    sz = get_run_font_size(rPr)
                    if sz and abs(sz - FONT_SIZES_HALF["五号"]) > TOLERANCE_HALF:
                        self.w_(f"[图题] 「{t[:20]}」字号应为五号(10.5pt)，检测到 {half_to_pt(sz)}pt")
                    break
        else:
            self.w_("[图] 未检测到图题（应为「图X-Y 图题名」格式）")

        if tab_captions:
            self.p(f"检测到 {len(tab_captions)} 个表题（表X-Y格式）")
        else:
            self.w_("[表] 未检测到表题（应为「表X-Y 表题名」格式）")

        if fig_refs:
            self.p(f"检测到 {len(fig_refs)} 处正确的图引用格式「如图X-Y所示」")
        if tab_refs:
            self.p(f"检测到 {len(tab_refs)} 处正确的表引用格式「如表X-Y所示」")

        # 续表
        xu = [get_para_text(p) for p in self.paragraphs if '续表' in get_para_text(p)]
        if tab_captions and not xu:
            self.w_("[表] 若有跨页表格，需在续页加「续表」标注")

    # ─── 参考文献 ─────────────────────────────────────────────
    def check_references(self):
        all_text = '\n'.join(get_para_text(p) for p in self.paragraphs)
        found = False
        for para in self.paragraphs:
            t = get_para_text(para).strip()
            if matches_label(t, '参考文献') or matches_label(t, 'References'):
                found = True
                pchain = self.pPr_chain(para)
                rchain = self.rPr_chain(para)
                if resolve_alignment(pchain) != 'center':
                    self.w_("[参考文献] 标题应居中（与一级标题同格式）")
                sz = resolve_run_font_size(rchain)
                if sz and abs(sz - FONT_SIZES_HALF["三号"]) > TOLERANCE_HALF:
                    self.w_(f"[参考文献] 标题字号应为三号(16pt)，检测到 {half_to_pt(sz)}pt")
                break
        if found:
            self.p("检测到参考文献章节")
        else:
            self.c("[参考文献] 未检测到参考文献章节")

        refs = re.findall(r'\[\d+\]', all_text)
        if refs:
            self.p(f"正文中检测到 {len(set(refs))} 个文献引用编号")
        else:
            self.w_("[参考文献] 正文中未检测到方括号引用格式 [1][2]...")

    # ─── 内容质量扫描 ─────────────────────────────────────────
    def scan_content(self):
        all_text = '\n'.join(get_para_text(p) for p in self.paragraphs)
        total = len(all_text.replace('\n', '').replace(' ', ''))
        self.content_notes.append(f"文档总字符数（估计）：约 {total:,}")

        if any(k in all_text for k in ('结论', '总结', 'Conclusion')):
            self.content_notes.append("✓ 包含结论/总结章节")
        else:
            self.content_notes.append("⚠ 未检测到明确的结论/总结章节")

        refs = re.findall(r'\[\d+\]', all_text)
        if len(refs) < 5:
            self.content_notes.append(f"⚠ 文献引用较少（{len(refs)} 处），建议增加文献支撑")
        else:
            self.content_notes.append(f"✓ 文献引用数量：{len(refs)} 处")

        figs = set(re.findall(r'图\s*\d+[-－]\d+', all_text))
        tabs = set(re.findall(r'表\s*\d+[-－]\d+', all_text))
        self.content_notes.append(f"图：{len(figs)} 个，表：{len(tabs)} 个")

        if not any(k in all_text for k in ('研究方法', '分析方法', '实验', '调研')):
            self.content_notes.append("⚠ 未检测到明确的研究方法描述")

    def run(self):
        self.check_header()
        self.check_structure()
        self.check_abstract()
        self.check_headings()
        self.check_body_text()
        self.check_figures_tables()
        self.check_references()
        self.scan_content()
        return {
            "mode": "XML（docx skill 解包模式）",
            "critical": self.issues_critical,
            "warning": self.issues_warning,
            "passed": self.passed,
            "content_notes": self.content_notes,
            "summary": {
                "critical_count": len(self.issues_critical),
                "warning_count": len(self.issues_warning),
                "passed_count": len(self.passed),
            }
        }


# ─── 入口 ─────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({
            "error": "用法: python check_format_xml.py <已解包的docx目录>\n"
                     "请先用 docx skill 的 unpack.py 解包：\n"
                     "  python <docx-skill>/scripts/office/unpack.py 论文.docx 解包目录/"
        }, ensure_ascii=False))
        sys.exit(1)

    unpacked = sys.argv[1]
    if not Path(unpacked).exists():
        print(json.dumps({"error": f"目录不存在: {unpacked}"}, ensure_ascii=False))
        sys.exit(1)

    if not (Path(unpacked) / 'word' / 'document.xml').exists():
        print(json.dumps({
            "error": f"未找到 word/document.xml，请确认是已解包的 docx 目录"
        }, ensure_ascii=False))
        sys.exit(1)

    try:
        checker = ThesisXMLChecker(unpacked)
        result = checker.run()
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except Exception as e:
        import traceback
        print(json.dumps({"error": str(e), "trace": traceback.format_exc()}, ensure_ascii=False))
        sys.exit(1)
